# writer.py
# Technical Proposal Writer - 技术标书写作Skill
# 版本: 2.0
# 依赖: python-docx, PyMuPDF, mammoth (可选)

import json
import re
import os
from pathlib import Path
from typing import List, Dict, Optional, Union, Tuple
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from dataclasses import dataclass, field
from enum import Enum
import traceback


# ============== 数据模型 ==============

class WritingMode(Enum):
    """写作模式"""
    FULL = "full"               # 全部写作
    SECTION = "section"          # 指定单个章节
    SECTIONS = "sections"        # 指定多个章节
    INCREMENTAL = "incremental"  # 增量修改


@dataclass
class TemplateSection:
    """模板章节结构"""
    title: str                  # 章节标题（清理后）
    raw_title: str              # 原始标题（含标记）
    level: int                  # 标题级别（1-9）
    number: str                 # 章节编号，如"3.4"
    full_path: str              # 完整路径，如"3.4 技术路线"
    parent_number: str          # 父章节编号
    instructions: str           # 写作指示
    must_write: bool            # 是否需要生成内容
    original_index: int         # 在原文档中的段落索引
    paragraph_element: object = None  # 原始XML元素引用

    def matches(self, target: str) -> bool:
        """检查是否匹配目标章节"""
        target = target.strip().lower()
        return (
            self.number == target or
            self.full_path.lower() == target or
            self.title.lower() == target or
            target in self.full_path.lower()
        )


@dataclass
class Reference:
    """参考文献"""
    id: str
    text: str = ""
    authors: str = ""
    title: str = ""
    source: str = ""
    year: str = ""
    doi: str = ""

    def format_citation(self) -> str:
        """格式化为引用标记"""
        return f"[{self.id}]"

    def format_bibliography(self) -> str:
        """格式化为参考文献条目"""
        parts = []
        if self.authors:
            parts.append(self.authors)
        if self.title:
            parts.append(self.title)
        if self.source:
            parts.append(f"[J]. {self.source}")
        if self.year:
            parts.append(f", {self.year}")
        return "".join(parts) if parts else self.text


@dataclass
class Equation:
    """公式"""
    id: str
    latex: str = ""
    mathml: str = ""
    label: str = ""     # 公式编号，如(3-1)
    caption: str = ""    # 公式说明


@dataclass
class Figure:
    """图表"""
    id: str
    image_path: str = ""
    label: str = ""     # 图编号，如图3-1
    caption: str = ""
    width: float = 14.0  # cm
    source_ref_ids: List[str] = field(default_factory=list)  # 来源参考文献编号
    source_location: str = ""   # 在原论文中的位置，如"第3章 图3-2"
    source_description: str = ""  # 来源说明，如"数据来源"或"图片来源"


@dataclass
class CitationConfig:
    """引用配置——用户对引用处理的偏好"""
    include_in_text: bool = True        # 正文中是否标注引用编号
    include_reference_list: bool = True  # 是否生成参考文献列表
    reference_source: str = "papers"    # "papers"|"template"|"auto"
    target_page_count: int = 30         # 预期篇幅（页），默认30页
    figure_count: int = 0               # 预计插入图片数量，用于页数估算


@dataclass
class FigurePolicy:
    """图片使用策略"""
    allow_self_drawn: bool = False       # 是否允许自行绘制图片（默认仅允许概览性技术路线图）
    prefer_text_over_table: bool = True  # 是否优先用文字替代表格
    max_table_items_for_text: int = 5    # 超过此项数的数据才考虑使用表格
    infer_from_caption: bool = True      # 无法提取图片时，是否从题注推知内容决定插入


# ============== 主类 ==============

class TechnicalProposalWriter:
    """
    技术标书写作Skill

    核心功能：
    1. 解析标书模板，提取写作指示和格式
    2. 从参考论文提取研究内容和引用文献
    3. 按模板要求生成标书章节内容
    4. 支持选择性章节写作和增量修改
    5. 保持引用一致性，图表公式来源可溯

    用法：
        writer = TechnicalProposalWriter(claude_client)

        # 全量写作
        writer.write_proposal(
            template_path="template.docx",
            papers_data=papers,
            output_path="output.docx"
        )

        # 选择性写作
        writer.write_proposal(
            template_path="template.docx",
            papers_data=papers,
            output_path="output.docx",
            write_sections="3.4"
        )

        # 增量修改
        writer.write_proposal(
            template_path="template.docx",
            papers_data=papers,
            output_path="output.docx",
            write_sections="3.5",
            base_document="existing.docx"
        )
    """

    # 需要AI写作的关键词
    WRITING_KEYWORDS = [
        '研究现状', '研究内容', '国内外现状', '文献综述',
        '技术路线', '研究方案', '关键技术', '实施方案',
        '可行性分析', '创新点', '研究目标', '研究思路',
        '研究背景', '发展趋势', '技术比较', '方案设计',
        '总体架构', '功能模块', '性能指标'
    ]

    # 占位符标记
    PLACEHOLDER_MARKERS = [
        '待填写', '请在此处撰写', '此处写入', 'TBD',
        'todo', '待补充', '[待', '【待', '{{待'
    ]

    def __init__(self, claude_client):
        """
        初始化

        Args:
            claude_client: Claude API客户端实例
        """
        self.claude = claude_client
        self.references_pool: Dict[str, Reference] = {}
        self.template_structure: List[TemplateSection] = []
        self.section_map: Dict[str, TemplateSection] = {}
        self.equations: List[Equation] = []
        self.figures: List[Figure] = []
        self.citation_config = CitationConfig()  # 默认引用配置，可由用户覆盖
        self.figure_policy = FigurePolicy()      # 默认图片策略，可由用户覆盖

    # ============== 公共API ==============

    def write_proposal(
        self,
        template_path: str,
        papers_data: List[Dict],
        output_path: str,
        write_sections: Optional[Union[str, List[str]]] = None,
        base_document: Optional[str] = None,
        include_citations: Optional[bool] = None,
        include_reference_list: Optional[bool] = None,
        reference_source: Optional[str] = None,
        allow_self_drawn_figures: Optional[bool] = None,
        prefer_text_over_table: Optional[bool] = None,
        infer_figures_from_caption: Optional[bool] = None,
        target_page_count: Optional[int] = None
    ) -> str:
        """
        写作技术标书

        Args:
            template_path: 标书模板Word文件路径
            papers_data: 参考论文数据列表，格式：
                [{
                    'filename': 'paper1.pdf',
                    'content': '论文完整文本内容',
                    'metadata': {'title': '...', 'authors': '...'}
                }]
            output_path: 输出文件路径
            write_sections: 要写作的章节，支持：
                - None: 写作所有需要生成的章节
                - "3.4": 只写3.4节
                - "研究现状": 按标题匹配
                - "3.1,3.4,研究内容": 多个章节
                - ["3.1", "3.4"]: 列表形式
            base_document: 基础文档路径（增量模式）
            include_citations: 正文中是否标注引用编号（None=使用默认True）
            include_reference_list: 是否生成参考文献列表（None=使用默认True）
            reference_source: 引用编号来源 "papers"|"template"|"auto"（None=使用默认"papers"）
            allow_self_drawn_figures: 是否允许自行绘制图片（None=使用默认False，仅技术路线图例外）
            prefer_text_over_table: 是否优先用文字替代表格（None=使用默认True）
            infer_figures_from_caption: 无法提取图片时是否从题注推知内容（None=使用默认True）
            target_page_count: 预期篇幅页数（None=使用默认30页）

        Returns:
            输出文件路径
        """
        try:
            print("=" * 60)
            print("📋 技术标书写作 Skill v2.0")
            print("=" * 60)

            # 1. 解析目标章节
            target_sections = self._resolve_target_sections(write_sections)
            mode = self._determine_mode(write_sections, base_document)

            # 1.5 应用引用配置
            if include_citations is not None:
                self.citation_config.include_in_text = include_citations
            if include_reference_list is not None:
                self.citation_config.include_reference_list = include_reference_list
            if reference_source is not None:
                self.citation_config.reference_source = reference_source

            print(f"\n🎯 写作模式: {mode.value}")
            print(f"📎 引用配置: 正文引用={'是' if self.citation_config.include_in_text else '否'}, "
                  f"参考文献列表={'是' if self.citation_config.include_reference_list else '否'}, "
                  f"编号来源={self.citation_config.reference_source}")

            # 1.6 应用图片策略
            if allow_self_drawn_figures is not None:
                self.figure_policy.allow_self_drawn = allow_self_drawn_figures
            if prefer_text_over_table is not None:
                self.figure_policy.prefer_text_over_table = prefer_text_over_table
            if infer_figures_from_caption is not None:
                self.figure_policy.infer_from_caption = infer_figures_from_caption

            print(f"🖼️ 图片策略: 自绘={'允许' if self.figure_policy.allow_self_drawn else '仅技术路线图'}, "
                  f"文字替表格={'是' if self.figure_policy.prefer_text_over_table else '否'}, "
                  f"题注推断={'是' if self.figure_policy.infer_from_caption else '否'}")

            # 1.7 应用篇幅配置
            if target_page_count is not None:
                self.citation_config.target_page_count = target_page_count
            # 根据图片数量调整有效文字页数估算：每张图片约占0.3页
            figure_pages = self.citation_config.figure_count * 0.3
            effective_text_pages = self.citation_config.target_page_count - figure_pages
            print(f"📏 预期篇幅: {self.citation_config.target_page_count}页"
                  f"（图片{self.citation_config.figure_count}张约占{figure_pages:.1f}页, "
                  f"文字约{effective_text_pages:.0f}页, "
                  f"预估{int(effective_text_pages * 800)}字）")

            # 2. 解析模板
            print("\n📄 [1/5] 解析标书模板...")
            template_info = self.parse_template(template_path)
            print(f"   ✅ 解析完成，共识别 {len(template_info['sections'])} 个章节")

            # 3. 筛选章节
            print("\n🔍 [2/5] 筛选目标章节...")
            sections_to_write = self._filter_sections(
                template_info['sections'],
                target_sections
            )

            if not sections_to_write:
                print("\n⚠️ 没有找到匹配的章节！")
                self.list_available_sections(template_info['sections'])
                return ""

            print(f"   需要写作 {len(sections_to_write)} 个章节:")
            for sec in sections_to_write:
                instruction_preview = sec.instructions[:60] + "..." if len(sec.instructions) > 60 else sec.instructions
                print(f"   📌 {sec.full_path}")
                if instruction_preview:
                    print(f"      └─ 指示: {instruction_preview}")

            # 4. 提取引用
            print(f"\n📚 [3/5] 提取参考文献...")
            self._build_reference_pool(papers_data)
            print(f"   ✅ 共提取 {len(self.references_pool)} 条参考文献")

            # 5. 生成内容
            print(f"\n✍️ [4/5] 生成章节内容...")
            section_contents = {}
            for i, section in enumerate(sections_to_write, 1):
                print(f"\n   [{i}/{len(sections_to_write)}] 生成: {section.full_path}")

                # 查找相关论文内容
                relevant_papers = self._filter_relevant_papers(papers_data, section)

                try:
                    content = self._generate_section_content(
                        section=section,
                        papers_data=relevant_papers,
                        format_info=template_info['format']
                    )
                    section_contents[section.full_path] = content
                    print(f"   ✅ 完成 ({len(content)} 字符)")
                except Exception as e:
                    print(f"   ❌ 生成失败: {str(e)}")
                    section_contents[section.full_path] = f"[生成失败: {str(e)}]"

            # 6. 输出文档
            print(f"\n📝 [5/5] 生成文档...")
            if base_document and os.path.exists(base_document):
                print(f"   📄 基础文档: {base_document}")
                self._modify_existing_document(base_document, template_info, section_contents, output_path)
            else:
                self._create_new_document(template_info, section_contents, output_path)

            print(f"\n{'=' * 60}")
            print(f"✨ 完成！文件已保存至: {output_path}")
            print(f"{'=' * 60}")

            return output_path

        except Exception as e:
            print(f"\n❌ 错误: {str(e)}")
            print(traceback.format_exc())
            raise

    def list_available_sections(self, sections: List[TemplateSection] = None):
        """列出可写作的章节"""
        if sections is None:
            sections = self.template_structure

        print("\n📋 模板中可写作的章节:")
        writable = [s for s in sections if s.must_write]
        other = [s for s in sections if not s.must_write]

        if writable:
            print("   【需要生成内容】")
            for sec in writable:
                print(f"   • {sec.full_path} (级别{sec.level})")
                if sec.instructions:
                    print(f"     └─ {sec.instructions[:80]}...")

        if other:
            print(f"\n   【已有内容/模板章节】(共{len(other)}个)")
            for sec in other[:5]:  # 只显示前5个
                print(f"   • {sec.full_path}")
            if len(other) > 5:
                print(f"   ... 还有 {len(other) - 5} 个章节")

    def extract_references_from_section(self, section_number: str) -> List[Reference]:
        """获取指定章节相关的参考文献"""
        section = self.section_map.get(section_number)
        if not section:
            # 尝试模糊匹配
            for num, sec in self.section_map.items():
                if section_number in num or section_number in sec.title:
                    section = sec
                    break

        return list(self.references_pool.values()) if section else []

    # ============== 模板解析 ==============

    def parse_template(self, template_path: str) -> Dict:
        """
        解析标书模板Word文档

        Returns:
            {
                'sections': [TemplateSection, ...],
                'format': {...},
                'raw_document': Document
            }
        """
        doc = Document(template_path)

        template_info = {
            'sections': [],
            'format': self._extract_document_format(doc),
            'raw_document': doc
        }

        # 追踪章节计数器
        section_counters = {}
        heading_styles = self._detect_heading_styles(doc)

        for idx, para in enumerate(doc.paragraphs):
            is_heading, level = self._check_heading(para, heading_styles)

            if is_heading:
                # 计算章节编号
                number = self._calculate_section_number(level, section_counters)
                clean_title = self._clean_title(para.text)

                section = TemplateSection(
                    title=clean_title,
                    raw_title=para.text,
                    level=level,
                    number=number,
                    full_path=f"{number} {clean_title}",
                    parent_number='.'.join(number.split('.')[:-1]) if '.' in number else '0',
                    instructions=self._extract_instructions(para, doc, idx),
                    must_write=self._check_writing_needed(clean_title),
                    original_index=idx,
                    paragraph_element=para._element
                )

                template_info['sections'].append(section)
                self.section_map[number] = section
                self.section_map[clean_title] = section

        self.template_structure = template_info['sections']
        return template_info

    def _detect_heading_styles(self, doc: Document) -> set:
        """检测文档中使用的标题样式"""
        heading_styles = set()
        for para in doc.paragraphs[:100]:  # 检查前100个段落
            style_name = para.style.name if para.style else ''
            style_name_lower = style_name.lower()
            if any(kw in style_name_lower for kw in ['heading', '标题', 'title']):
                heading_styles.add(style_name)
        return heading_styles

    def _check_heading(self, para, heading_styles: set) -> Tuple[bool, int]:
        """检查段落是否为标题，返回(是否标题, 级别)"""
        style_name = para.style.name if para.style else ''
        style_name_lower = style_name.lower()

        # 标准标题样式
        if 'heading' in style_name_lower or '标题' in style_name:
            for i in range(1, 10):
                if str(i) in style_name:
                    return True, i
            return True, 1

        # 自定义标题样式
        if style_name in heading_styles:
            return True, 1

        # 通过格式判断（大字、加粗、居中）
        if para.runs:
            run = para.runs[0]
            font_size = run.font.size
            is_bold = run.font.bold
            alignment = para.paragraph_format.alignment

            if font_size and font_size >= Pt(14) and is_bold:
                return True, 1
            if font_size and font_size >= Pt(13) and is_bold and alignment == WD_ALIGN_PARAGRAPH.CENTER:
                return True, 1

        return False, 0

    def _calculate_section_number(self, level: int, counters: Dict) -> str:
        """计算章节编号（如3.4, 3.4.1）"""
        # 清除更深层级的计数器
        for l in list(counters.keys()):
            if l > level:
                del counters[l]

        # 递增当前层级
        if level not in counters:
            counters[level] = 0
        counters[level] += 1

        # 构建编号
        parts = []
        for l in range(1, level + 1):
            if l in counters:
                parts.append(str(counters[l]))

        return '.'.join(parts)

    def _clean_title(self, text: str) -> str:
        """清理标题文本，移除指示标记"""
        # 移除【】标记
        text = re.sub(r'【.*?】', '', text)
        # 移除{{}}标记
        text = re.sub(r'\{\{.*?\}\}', '', text)
        # 移除多余空格
        text = re.sub(r'\s+', ' ', text)
        return text.strip()

    def _check_writing_needed(self, title: str) -> bool:
        """判断章节是否需要AI写作"""
        return any(kw in title for kw in self.WRITING_KEYWORDS)

    def _extract_instructions(self, para, doc: Document, idx: int) -> str:
        """提取写作指示"""
        instructions = []
        text = para.text

        # 1. 【】标记的指示
        brackets = re.findall(r'【(.*?)】', text)
        instructions.extend(brackets)

        # 2. {{}}标记的指示
        curlies = re.findall(r'\{\{(.*?)\}\}', text)
        instructions.extend(curlies)

        # 3. 括号内的写作要求
        parens = re.findall(r'[（(](.*?)[）)]', text)
        for p in parens:
            if any(kw in p for kw in ['需要', '必须', '要求', '包含', '包括', '不少于', '至少']):
                instructions.append(p)

        # 4. 检查后续段落是否是小字注释
        try:
            if idx + 1 < len(doc.paragraphs):
                next_para = doc.paragraphs[idx + 1]
                if next_para.runs:
                    font_size = next_para.runs[0].font.size
                    if font_size and font_size < Pt(10):
                        comment = next_para.text.strip()
                        if comment:
                            instructions.append(f"[格式注释] {comment}")
        except:
            pass

        return '；'.join(instructions) if instructions else ''

    def _is_placeholder(self, text: str) -> bool:
        """检查文本是否为占位符"""
        text_lower = text.strip().lower()
        return any(marker.lower() in text_lower for marker in self.PLACEHOLDER_MARKERS)

    # ============== 格式提取 ==============

    def _extract_document_format(self, doc: Document) -> Dict:
        """提取文档格式信息"""
        format_info = {
            'page': {
                'width_cm': 21.0,
                'height_cm': 29.7,
                'margin_top_cm': 2.54,
                'margin_bottom_cm': 2.54,
                'margin_left_cm': 3.18,
                'margin_right_cm': 3.18
            },
            'body_font': {
                'name': '宋体',
                'size_pt': 12,
                'name_ascii': 'Times New Roman'
            },
            'heading_font': {
                'name': '黑体',
                'size_pt': 14
            },
            'line_spacing': 1.5,
            'first_line_indent_cm': 0.74,
            'alignment': 'justify'
        }

        try:
            # 页面设置
            section = doc.sections[0]
            if section.page_width:
                format_info['page']['width_cm'] = round(section.page_width.cm, 1)
            if section.page_height:
                format_info['page']['height_cm'] = round(section.page_height.cm, 1)
            if section.top_margin:
                format_info['page']['margin_top_cm'] = round(section.top_margin.cm, 1)
            if section.bottom_margin:
                format_info['page']['margin_bottom_cm'] = round(section.bottom_margin.cm, 1)
            if section.left_margin:
                format_info['page']['margin_left_cm'] = round(section.left_margin.cm, 1)
            if section.right_margin:
                format_info['page']['margin_right_cm'] = round(section.right_margin.cm, 1)

            # 查找正文段落格式
            for para in doc.paragraphs:
                if not self._check_heading(para, set())[0] and para.runs:
                    run = para.runs[0]
                    if run.font.name:
                        format_info['body_font']['name'] = run.font.name
                    if run.font.size:
                        format_info['body_font']['size_pt'] = run.font.size.pt
                    if para.paragraph_format.line_spacing:
                        format_info['line_spacing'] = para.paragraph_format.line_spacing
                    if para.paragraph_format.first_line_indent:
                        format_info['first_line_indent_cm'] = round(
                            para.paragraph_format.first_line_indent.cm, 2
                        )
                    break

            # 查找标题格式
            for para in doc.paragraphs:
                is_heading, _ = self._check_heading(para, set())
                if is_heading and para.runs:
                    run = para.runs[0]
                    if run.font.name:
                        format_info['heading_font']['name'] = run.font.name
                    if run.font.size:
                        format_info['heading_font']['size_pt'] = run.font.size.pt
                    break

        except Exception as e:
            print(f"   ⚠️ 格式提取有警告: {e}")

        return format_info

    # ============== 章节筛选 ==============

    def _resolve_target_sections(self, write_sections) -> Optional[List[str]]:
        """解析用户指定的目标章节"""
        if write_sections is None:
            return None

        if isinstance(write_sections, str):
            # 支持逗号、顿号、分号分隔
            separators = r'[,，、;；]'
            if re.search(separators, write_sections):
                return [s.strip() for s in re.split(separators, write_sections) if s.strip()]
            return [write_sections.strip()]

        if isinstance(write_sections, list):
            return [str(s).strip() for s in write_sections]

        return None

    def _determine_mode(self, write_sections, base_document) -> WritingMode:
        """确定写作模式"""
        if write_sections is None:
            return WritingMode.FULL
        elif base_document and os.path.exists(base_document):
            return WritingMode.INCREMENTAL
        elif isinstance(write_sections, list) and len(write_sections) > 1:
            return WritingMode.SECTIONS
        else:
            return WritingMode.SECTION

    def _filter_sections(
        self,
        all_sections: List[TemplateSection],
        target_sections: Optional[List[str]]
    ) -> List[TemplateSection]:
        """筛选要写作的章节"""
        if target_sections is None:
            # 全量模式：返回所有需要写作的章节
            return [s for s in all_sections if s.must_write]

        # 匹配指定章节
        matched = []
        for target in target_sections:
            found = False
            for section in all_sections:
                if section.matches(target):
                    if section not in matched:
                        matched.append(section)
                    found = True

            if not found:
                print(f"   ⚠️ 未找到匹配章节: '{target}'")
                # 尝试模糊匹配
                for section in all_sections:
                    if target.lower() in section.title.lower():
                        if section not in matched:
                            matched.append(section)
                            print(f"   💡 模糊匹配到: {section.full_path}")
                            break

        # 保持原有顺序
        matched.sort(key=lambda s: s.original_index)
        return matched

    # ============== 引用管理 ==============

    def _build_reference_pool(self, papers_data: List[Dict]):
        """从多篇论文中提取和合并参考文献"""
        # 避免重复提取
        if self.references_pool:
            return

        # 使用Claude提取引用
        prompt = self._build_reference_extraction_prompt(papers_data)

        try:
            response = self.claude.messages.create(
                model="claude-3-5-haiku-20241022",
                max_tokens=4000,
                temperature=0,
                messages=[{"role": "user", "content": prompt}]
            )

            refs_dict = self._parse_json_response(response.content[0].text)

            for ref_id, ref_data in refs_dict.items():
                self.references_pool[str(ref_id)] = Reference(
                    id=str(ref_id),
                    authors=ref_data.get('authors', ''),
                    title=ref_data.get('title', ''),
                    source=ref_data.get('source', ''),
                    year=str(ref_data.get('year', '')),
                    doi=ref_data.get('doi', ''),
                    text=ref_data.get('full_text', '')
                )

        except Exception as e:
            print(f"   ⚠️ AI提取引用失败，使用规则提取: {e}")
            self._fallback_reference_extraction(papers_data)

    def _build_reference_extraction_prompt(self, papers_data: List[Dict]) -> str:
        """构建引用提取的AI提示"""
        papers_text = []
        for i, paper in enumerate(papers_data):
            filename = paper.get('filename', f'paper_{i+1}')
            content = paper.get('content', '')
            # 主要查找参考文献部分
            ref_section = self._find_reference_section(content)
            papers_text.append(f"===论文{i+1}: {filename}===\n{ref_section[:5000]}")

        return f"""请从以下论文的参考文献部分提取所有引用文献。

要求：
1. 保持原论文中的引用编号（如[1]、[2]等）
2. 如果同一篇文献在多篇论文中被引用，以首次出现的编号为准
3. 提取完整的文献信息

输出JSON格式：
{{
    "1": {{
        "authors": "作者列表",
        "title": "文献标题",
        "source": "期刊/会议名称",
        "year": "发表年份",
        "doi": "DOI号（如有）",
        "full_text": "完整的引用文本"
    }}
}}

论文内容：
{chr(10).join(papers_text)}

请只输出JSON，不要有其他文字。"""

    def _find_reference_section(self, content: str) -> str:
        """在论文内容中查找参考文献部分"""
        # 常见的参考文献标题
        patterns = [
            r'(?:参考文献|References|Bibliography|REFERENCES)[\s\S]*',
            r'\[\d+\][\s\S]*',  # 以[1]开头的部分
        ]

        for pattern in patterns:
            match = re.search(pattern, content, re.IGNORECASE)
            if match:
                return match.group(0)

        # 如果没找到，返回最后30%的内容（参考文献通常在末尾）
        return content[-len(content)//3:]

    def _fallback_reference_extraction(self, papers_data: List[Dict]):
        """备用方案：用正则表达式提取引用"""
        # 匹配 [1] Author. Title[J]. Journal, Year.
        ref_pattern = r'\[(\d+)\]\s*(.+?)(?=\[\d+\]|\Z)'

        for paper in papers_data:
            content = paper.get('content', '')
            ref_section = self._find_reference_section(content)

            matches = re.findall(ref_pattern, ref_section, re.DOTALL)
            for ref_id, ref_text in matches:
                ref_text = ref_text.strip()[:300]
                if ref_id not in self.references_pool:
                    # 尝试提取作者和标题
                    parts = ref_text.split('.')
                    authors = parts[0] if len(parts) > 0 else ''
                    title = parts[1] if len(parts) > 1 else ''

                    self.references_pool[ref_id] = Reference(
                        id=ref_id,
                        text=ref_text,
                        authors=authors.strip(),
                        title=title.strip()
                    )

    def _filter_relevant_papers(self, papers_data: List[Dict], section: TemplateSection) -> List[Dict]:
        """筛选与章节相关的论文"""
        # 根据章节关键词筛选相关论文
        section_keywords = self._extract_section_keywords(section)

        if not section_keywords:
            return papers_data  # 返回所有论文

        relevant = []
        for paper in papers_data:
            content = paper.get('content', '').lower()
            # 计算相关性得分
            score = sum(1 for kw in section_keywords if kw.lower() in content)
            if score > 0:
                paper_copy = paper.copy()
                paper_copy['relevance_score'] = score
                relevant.append(paper_copy)

        # 按相关性排序
        relevant.sort(key=lambda p: p.get('relevance_score', 0), reverse=True)

        return relevant if relevant else papers_data

    def _extract_section_keywords(self, section: TemplateSection) -> List[str]:
        """从章节中提取关键词"""
        # 预定义的章节关键词映射
        keyword_map = {
            '研究现状': ['literature review', 'related work', 'state of the art', '综述', '进展'],
            '研究内容': ['method', 'approach', 'proposed', 'framework', '方法', '方案'],
            '技术路线': ['architecture', 'workflow', 'pipeline', 'implementation', '实施'],
            '关键技术': ['algorithm', 'technique', 'key technology', '算法', '核心技术'],
            '可行性分析': ['feasibility', 'validation', 'experiment', '实验', '验证'],
            '创新点': ['contribution', 'novelty', 'innovation', '创新', '贡献']
        }

        keywords = []
        for pattern, kws in keyword_map.items():
            if pattern in section.title:
                keywords.extend(kws)

        # 从标题和指示中提取额外关键词
        if section.instructions:
            # 提取中文关键词（2-4字）
            cn_keywords = re.findall(r'[一-鿿]{2,4}', section.instructions)
            keywords.extend(cn_keywords[:5])

        return list(set(keywords))

    # ============== 内容生成 ==============

    def _generate_section_content(
        self,
        section: TemplateSection,
        papers_data: List[Dict],
        format_info: Dict
    ) -> str:
        """为单个章节生成内容"""

        prompt = self._build_section_prompt(section, papers_data, format_info)

        response = self.claude.messages.create(
            model="claude-3-5-haiku-20241022",
            max_tokens=4000,
            temperature=0.3,
            messages=[{"role": "user", "content": prompt}]
        )

        return response.content[0].text

    def _build_section_prompt(
        self,
        section: TemplateSection,
        papers_data: List[Dict],
        format_info: Dict
    ) -> str:
        """构建章节写作的AI提示"""

        cfg = self.citation_config

        # 参考文献
        refs_list = sorted(
            self.references_pool.values(),
            key=lambda r: int(r.id) if r.id.isdigit() else float('inf')
        )
        refs_text = "\n".join([
            f"[{ref.id}] {ref.format_bibliography()[:200]}"
            for ref in refs_list
        ])

        # 论文内容
        papers_text = []
        for i, paper in enumerate(papers_data):
            filename = paper.get('filename', f'论文{i+1}')
            content = paper.get('content', '')
            # 限制长度，保留最相关部分
            max_len = min(len(content), 8000)
            papers_text.append(f"### 参考素材{i+1}: {filename}\n{content[:max_len]}")

        # 写作指示
        instructions = section.instructions if section.instructions else "按照学术标书规范撰写，要求内容详实、论证充分"

        # 构建引用规则（根据用户配置）
        citation_rules = self._build_citation_rules(cfg, refs_text)

        # 构建图片和表格规则
        figure_table_rules = self._build_figure_table_rules(self.figure_policy)

        # 篇幅信息
        fp = self.figure_policy
        cfg = self.citation_config
        figure_pages_est = cfg.figure_count * 0.3
        target_chars = int((cfg.target_page_count - figure_pages_est) * 800)

        prompt = f"""你是一位经验丰富的技术标书撰写专家。请撰写标书的「{section.title}」章节。

## 📋 写作要求
{instructions}

## 📏 篇幅要求
- 预期总篇幅：约{cfg.target_page_count}页
- 其中图片约占{figure_pages_est:.1f}页（{cfg.figure_count}张，每张按14cm×10cm常规尺寸计）
- 文字部分约需{target_chars}字（每页约800汉字）
- 请据此控制内容详细程度和段落密度

## 📐 格式规范
- 正文：{format_info['body_font']['name']}，{format_info['body_font']['size_pt']}pt
- 行距：{format_info['line_spacing']}倍
- 首行缩进：2字符
- 引用格式：上标方括号，如 [1]、[2,3]、[4-6]

{citation_rules}

{figure_table_rules}

## 📄 参考论文素材
{chr(10).join(papers_text)}

## ⚠️ 核心规则
1. **必须对论文内容进行改写**，不得直接复制原文
2. **保持学术严谨性**，每个观点都要有依据
3. **结构清晰**，合理分段，每段300-500字
4. **语言风格**：正式、客观、准确的技术中文

请直接输出章节正文内容（不要输出标题）：
"""

        return prompt

    def _build_figure_table_rules(self, fp: 'FigurePolicy') -> str:
        """根据图片策略构建图片和表格使用规则"""
        parts = []

        # 图片规则
        parts.append("## 🖼️ 图片使用规则")
        parts.append("- **除开篇概览性技术路线图外，不得自行绘制任何图片**")
        parts.append("- 如需插入图片，必须来自参考论文素材中的真实插图")
        parts.append("- 每张图片标注：文献名称 + 原文章节 + 原图编号")
        parts.append("- 图片编号使用标书自身体系（如图3.5-1），同时注明原论文出处")

        if fp.infer_from_caption:
            parts.append("- 当参考论文中的图片无法直接提取时，从论文的图题注和上下文推知图片大概内容")
            parts.append("- 根据推知的内容判断是否对标书有价值：")
            parts.append('  有价值 -> 标注插入点:【插入图片:文献名,第X章 图X-Y,题注:"..."】同时用文字概述核心信息')
            parts.append('  无价值 -> 跳过,直接用文字概述该图的相关结论')
            parts.append('- 插入点处的文字说明应使文档即使暂未插入图片也可独立阅读')

        # 表格规则
        if fp.prefer_text_over_table:
            parts.append('\n## 表格使用规则')
            parts.append('- **优先使用文字叙述替代表格**')
            parts.append('- 对比场景:以每个对象为一段展开叙述,不使用表格对比形式')
            parts.append('- 参数罗列:写成自然段落(如电化学储能功率成本800~1200元/kW,飞轮储能对应为1500~2500元/kW)')
            parts.append(f"- 仅当数据项超过{fp.max_table_items_for_text}项且每项有明确数值时才考虑使用表格")
            parts.append("- 若必须使用表格，仍需在表下方标注数据来源")

        return "\n".join(parts)

    def _build_citation_rules(self, cfg: 'CitationConfig', refs_text: str) -> str:
        """根据引用配置构建引用规则文本"""
        parts = []

        if cfg.include_in_text:
            parts.append("## 📚 可用参考文献（必须使用这些引用编号）")
            parts.append(refs_text)
            parts.append("\n**引用要求**：")
            parts.append("- 关键论述必须有文献支撑，引用分布合理")
            parts.append("- 引用格式：上标方括号，如 [1]、[2,3]、[4-6]")
            parts.append("- **必须使用上文列出的参考文献编号**进行引用，不得自创引用")
        else:
            parts.append("## ⚠️ 引用要求")
            parts.append("- **不要**在正文中标注引用编号（如[1]、[2]等）")
            parts.append("- 如需提及某研究，直接描述其作者或内容，不使用编号引用")
            parts.append("- 可参考论文素材中的文献信息，但不需要标注编号")

        if cfg.include_reference_list:
            parts.append(f"\n## 📋 参考文献列表")
            parts.append("- 请在输出的末尾，列出所有在正文中引用过的参考文献")
            parts.append("- 格式：编号. 作者. 标题[J]. 期刊, 年份.")

        return "\n".join(parts)

    # ============== 文档生成 ==============

    def _create_new_document(
        self,
        template_info: Dict,
        section_contents: Dict,
        output_path: str
    ):
        """创建新的标书文档"""
        doc = template_info['raw_document']
        format_info = template_info['format']

        # 遍历段落，填充内容
        current_content = None
        content_paragraphs = []

        for para in doc.paragraphs:
            # 检查是否为标题
            is_heading, _ = self._check_heading(para, set())

            if is_heading:
                # 先处理上一个章节的累积内容
                if current_content and content_paragraphs:
                    self._fill_paragraphs(content_paragraphs, current_content, format_info)
                    content_paragraphs = []
                    current_content = None

                # 查找是否有对应的生成内容
                clean_title = self._clean_title(para.text)
                for full_path, content in section_contents.items():
                    if clean_title in full_path:
                        current_content = content
                        break

            elif current_content:
                # 检查是否为占位符
                if self._is_placeholder(para.text) or not para.text.strip():
                    content_paragraphs.append(para)

        # 处理最后一个章节
        if current_content and content_paragraphs:
            self._fill_paragraphs(content_paragraphs, current_content, format_info)

        # 保存
        doc.save(output_path)

    def _fill_paragraphs(
        self,
        paragraphs: list,
        content: str,
        format_info: Dict
    ):
        """将生成的内容填充到占位段落中"""
        content_parts = self._split_content_into_paragraphs(content)

        for i, para in enumerate(paragraphs):
            if i < len(content_parts):
                # 清空原有内容
                for run in para.runs:
                    run.text = ''

                # 添加新内容
                part = content_parts[i].strip()
                if part:
                    # 处理引用标记
                    self._add_text_with_citations(para, part, format_info)
            else:
                # 多余段落保持为空
                for run in para.runs:
                    run.text = ''

    def _split_content_into_paragraphs(self, content: str) -> List[str]:
        """将内容分割为段落"""
        # 按双换行分割
        parts = re.split(r'\n\s*\n', content)
        # 清理空段落
        return [p.strip() for p in parts if p.strip()]

    def _add_text_with_citations(self, para, text: str, format_info: Dict):
        """添加带引用标记的文本"""
        # 分割引用标记 [1], [2,3], [4-6]
        pattern = r'(\[\d+(?:[-,]\d+)*\])'
        parts = re.split(pattern, text)

        font_name = format_info.get('body_font', {}).get('name', '宋体')
        font_size = format_info.get('body_font', {}).get('size_pt', 12)

        first_run = True
        for i, part in enumerate(parts):
            if not part:
                continue

            # 检查是否为引用标记
            if re.match(pattern, part):
                run = para.add_run(part)
                run.font.superscript = True
                run.font.size = Pt(font_size * 0.75)  # 引用稍小
            else:
                run = para.add_run(part)
                if first_run:
                    run.font.name = font_name
                    run.font.size = Pt(font_size)
                    first_run = False
                else:
                    run.font.name = font_name
                    run.font.size = Pt(font_size)

    def _modify_existing_document(
        self,
        base_document_path: str,
        template_info: Dict,
        section_contents: Dict,
        output_path: str
    ):
        """修改现有文档的指定章节"""
        doc = Document(base_document_path)
        format_info = template_info['format']

        # 查找并替换指定章节的内容
        for para in doc.paragraphs:
            is_heading, _ = self._check_heading(para, set())

            if is_heading:
                clean_title = self._clean_title(para.text)
                for full_path, content in section_contents.items():
                    if clean_title in full_path:
                        # 找到了要修改的章节
                        self._replace_section_content(
                            doc, para, content, format_info
                        )
                        break

        doc.save(output_path)

    def _replace_section_content(
        self,
        doc: Document,
        heading_para,
        content: str,
        format_info: Dict
    ):
        """替换指定章节的内容"""
        body = doc.element.body
        heading_element = heading_para._element

        # 收集要删除的元素（标题后到下一个标题前）
        elements_to_remove = []
        next_elem = heading_element.getnext()

        while next_elem is not None:
            # 检查是否是新标题
            pPr = next_elem.find(qn('w:pPr'))
            if pPr is not None:
                pStyle = pPr.find(qn('w:pStyle'))
                if pStyle is not None:
                    style_val = pStyle.get(qn('w:val'), '')
                    if 'heading' in style_val.lower() or '标题' in style_val:
                        break

            elements_to_remove.append(next_elem)
            next_elem = next_elem.getnext()

        # 删除旧内容
        for elem in elements_to_remove:
            body.remove(elem)

        # 插入新内容
        content_parts = self._split_content_into_paragraphs(content)
        insert_after = heading_element

        for part in content_parts:
            if part.strip():
                # 创建新段落
                new_p = doc.add_paragraph()
                self._add_text_with_citations(new_p, part, format_info)

                # 设置段落格式
                new_p.paragraph_format.first_line_indent = Cm(
                    format_info.get('first_line_indent_cm', 0.74)
                )
                new_p.paragraph_format.line_spacing = format_info.get('line_spacing', 1.5)

                # 移动到标题后面
                body.remove(new_p._element)
                insert_after.addnext(new_p._element)
                insert_after = new_p._element

        # 移除段落间的多余空行
        self._clean_empty_paragraphs(body, heading_element)

    def _clean_empty_paragraphs(self, body, heading_element):
        """清理标题后的空段落"""
        next_elem = heading_element.getnext()
        empty_elems = []

        while next_elem is not None:
            # 检查是否为空段落
            text = ''.join(node.text or '' for node in next_elem.iter()
                          if node.tag == qn('w:t'))
            if not text.strip():
                empty_elems.append(next_elem)
                next_elem = next_elem.getnext()
            else:
                break

        for elem in empty_elems:
            try:
                body.remove(elem)
            except:
                pass

    # ============== 工具方法 ==============

    # ============== 图表溯源 ==============

    @staticmethod
    def format_figure_source_annotation(
        figure_label: str,
        ref_ids: List[str],
        source_location: str = "",
        description: str = "图片来源"
    ) -> str:
        """
        格式化图表来源标注文本

        Args:
            figure_label: 图表编号，如"图3.5-1"
            ref_ids: 来源参考文献编号列表，如["1", "3"]
            source_location: 在原论文中的位置，如"第3章 图3-5"
            description: 描述类型，"图片来源"或"数据来源"

        Returns:
            格式化后的标注文本

        Examples:
            >>> format_figure_source_annotation("图3.5-1", ["1"], "第2.3节 图2-5", "数据来源")
            "（数据来源：参考文献[1]，第2.3节 图2-5）"

            >>> format_figure_source_annotation("表3.5-2", ["1", "3"], "", "数据来源")
            "（数据来源：综合参考文献[1]、[3]）"
        """
        if not ref_ids:
            return ""

        if len(ref_ids) == 1:
            ref_str = f"参考文献[{ref_ids[0]}]"
        else:
            ref_str = "综合" + "、".join(f"参考文献[{rid}]" for rid in ref_ids)

        if source_location:
            return f"（{description}：{ref_str}，{source_location}）"
        else:
            return f"（{description}：{ref_str}）"

    @staticmethod
    def format_figure_insertion_marker(
        paper_name: str,
        chapter: str,
        figure_number: str,
        caption: str = "",
        description: str = ""
    ) -> str:
        """
        当参考论文图片无法提取时，生成插入点标注

        Args:
            paper_name: 文献名称
            chapter: 原文章节，如"第3章"
            figure_number: 原图编号，如"图3-5"
            caption: 原图题注
            description: 对该图内容的文字概述（无法提取时尤其重要）

        Returns:
            插入点标注文本

        Example:
            >>> format_figure_insertion_marker(
            ...     "祝元恺，浙江大学博士学位论文，2024",
            ...     "第2.3节", "图2-5",
            ...     "储能系统全生命周期成本构成",
            ...     "该图展示了初始投资、运维、更换、损耗和回收五类成本占比"
            ... )
            【插入图片：祝元恺，浙江大学博士学位论文，2024，第2.3节 图2-5，题注："储能系统全生命周期成本构成"。内容概述：该图展示了初始投资、运维、更换、损耗和回收五类成本占比】
        """
        parts = [f"【插入图片：{paper_name}，{chapter} {figure_number}"]
        if caption:
            parts.append(f'，题注："{caption}"')
        parts.append("】")
        if description:
            parts.append(f"\n（内容概述：{description}）")
        return "".join(parts)

    @staticmethod
    def estimate_pages(char_count: int, figure_count: int = 0,
                       has_formulas: bool = False) -> dict:
        """
        估算页数

        Args:
            char_count: 总字符数
            figure_count: 图片数量
            has_formulas: 是否包含公式

        Returns:
            {'text_pages': 文字页数, 'figure_pages': 图片页数, 'total_pages': 总页数}

        计算规则：
        - 纯文字: 800字/页
        - 含公式: 700字/页
        - 图片: 每张0.3页（14cm×10cm常规尺寸，占A4版面约1/4~1/3）
        """
        chars_per_page = 700 if has_formulas else 800
        text_pages = char_count / chars_per_page
        figure_pages = figure_count * 0.3
        return {
            'text_pages': round(text_pages, 1),
            'figure_pages': round(figure_pages, 1),
            'total_pages': round(text_pages + figure_pages, 1)
        }

    def set_figure_count(self, count: int):
        """设置预计插入的图片数量，用于准确估算页数"""
        self.citation_config.figure_count = count

    def register_figure_source(
        self,
        figure_id: str,
        ref_ids: List[str],
        source_location: str = "",
        description: str = "图片来源"
    ):
        """为已注册的图表补充来源信息

        Args:
            figure_id: 图表ID（与注册时一致）
            ref_ids: 来源参考文献编号
            source_location: 原论文中的位置
            description: "图片来源"或"数据来源"
        """
        for fig in self.figures:
            if fig.id == figure_id:
                fig.source_ref_ids = ref_ids
                fig.source_location = source_location
                fig.source_description = description
                break
        else:
            # 图表尚未注册，创建一个
            self.figures.append(Figure(
                id=figure_id,
                source_ref_ids=ref_ids,
                source_location=source_location,
                source_description=description
            ))

    def _parse_json_response(self, text: str) -> Dict:
        """从AI响应中解析JSON"""
        # 尝试提取代码块中的JSON
        match = re.search(r'```(?:json)?\s*\n?(.*?)\n?```', text, re.DOTALL)
        if match:
            return json.loads(match.group(1))

        # 尝试直接找JSON对象
        match = re.search(r'\{.*\}', text, re.DOTALL)
        if match:
            return json.loads(match.group(0))

        return {}

    def get_statistics(self) -> Dict:
        """获取统计信息"""
        return {
            'total_sections': len(self.template_structure),
            'writable_sections': len([s for s in self.template_structure if s.must_write]),
            'total_references': len(self.references_pool),
            'total_equations': len(self.equations),
            'total_figures': len(self.figures),
            'sections': [
                {
                    'number': s.number,
                    'title': s.title,
                    'must_write': s.must_write,
                    'has_instructions': bool(s.instructions)
                }
                for s in self.template_structure
            ]
        }


# ============== 便捷函数 ==============

def create_test_template(output_path: str = "test_template.docx"):
    """创建测试用标书模板"""
    doc = Document()

    # 设置页面
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)

    # 标题
    title = doc.add_heading('技术方案', level=1)

    # 研究现状（需要写作）
    h2 = doc.add_heading('3.1 国内外研究现状【需要涵盖近5年文献，不少于2000字】', level=2)
    doc.add_paragraph('请在此处撰写研究现状')

    # 研究内容（需要写作）
    doc.add_heading('3.2 研究内容{{需包含技术创新点}}', level=2)
    doc.add_paragraph('待填写')

    # 技术路线（需要写作）
    doc.add_heading('3.4 技术路线', level=2)
    doc.add_paragraph('此处写入技术路线')

    # 已有内容章节
    doc.add_heading('3.3 项目背景', level=2)
    doc.add_paragraph('本项目源于...（这是已有的内容，不需要AI生成）')

    doc.save(output_path)
    print(f"测试模板已创建: {output_path}")
    return output_path


if __name__ == "__main__":
    # 测试代码
    print("Technical Proposal Writer Skill v2.0")
    print("=" * 40)

    # 创建测试模板
    create_test_template("test_template.docx")

    print("\n使用方法:")
    print("""
    from writer import TechnicalProposalWriter

    writer = TechnicalProposalWriter(claude_client)

    # 查看可写作章节
    writer.list_available_sections()

    # 写作标书（含引用控制）
    writer.write_proposal(
        template_path="template.docx",
        papers_data=papers_from_pdf_reader,
        output_path="output.docx",
        write_sections="3.1",          # 可选，不指定则全部写作
        include_citations=True,         # 正文中是否标注引用编号
        include_reference_list=True,    # 是否生成参考文献列表
        reference_source="papers"       # 引用编号来源
    )

    # 图表溯源标注
    writer.register_figure_source(
        figure_id="fig_3_5_1",
        ref_ids=["1", "3"],
        source_location="第3章 图3-5",
        description="数据来源"
    )
    annotation = writer.format_figure_source_annotation(
        "图3.5-1", ["1", "3"], "第2.3节 图2-5", "数据来源"
    )
    print(annotation)  # （数据来源：综合参考文献[1]、[3]，第2.3节 图2-5）
    """)
