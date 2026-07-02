from __future__ import annotations


# 提示词工程集中放在本文件，app.py 只负责传入材料类型、语言模式和目标长度。
# 这样后续优化摘要风格时，不需要在 Streamlit 页面逻辑里到处改字符串。

def prompt_material_label(academic: bool, material_style: str) -> str:
    """把内部材料类型转换为提示词中的自然语言标签。"""
    if material_style == "presentation":
        return "PPT演示材料"
    if material_style == "visual":
        return "图片或扫描件OCR材料"
    if academic:
        return "学术论文或研究报告"
    return "业务或通用文档"


def prompt_language_contract(summary_mode: str) -> str:
    """生成语言约束，避免双语摘要混排或英文落入中文框。"""
    if summary_mode == "英文摘要":
        return "只输出英文，不要夹杂中文解释。"
    if summary_mode == "双语摘要":
        return "先输出一段中文摘要，再另起一段以 English Summary 开头输出英文摘要；中英文内容不要混在同一段。"
    return "只输出中文，不要夹杂英文段落；必要英文术语可以保留。"


def prompt_evidence_contract(material_style: str) -> str:
    """生成证据约束，减少模型根据碎片化文本过度推断。"""
    # 所有摘要模式共用这组事实边界，防止模型在 OCR 噪声上过度发挥。
    common_rules = [
        "所有事实、变量、指标、结论和建议必须能在原文中找到依据。",
        "原文没有给出的样本量、年份、实验数值、效果判断、适用边界和政策建议不得补写。",
        "遇到 OCR 噪声、页眉页脚、目录、页码或重复标题时，应忽略其对摘要主线的干扰。",
        "不要使用“原文说明、原文指出、该文档提到、这段内容主要讨论”等转述腔。",
        "不要输出 Markdown、项目符号、表格、代码块或解释过程。",
    ]
    if material_style in {"presentation", "visual"}:
        common_rules.extend(
            [
                "PPT和图片常只有标题、短句和图中文字，应整理为可读说明，不要强行写成论文结论。",
                "只有原文明确出现风险、行动项或实验结果时，才写相关内容。",
                "公式、流程、变量和图示关系要转写成自然语言，避免保留箭头、加号、斜杠和括号堆叠。",
            ]
        )
    return "\n".join(f"- {rule}" for rule in common_rules)


def top_venue_abstract_contract(summary_mode: str) -> str:
    """顶刊/顶会摘要风格合同：模仿结构和语气，不复制范例事实。"""
    language = prompt_language_contract(summary_mode)
    return "\n".join(
        [
            f"- 语言约束：{language}",
            "- 写作目标：模仿顶刊/顶会摘要的论证结构、克制语气和信息密度，而不是普通读书笔记或材料概述。",
            "- 推荐结构：背景缺口或现实问题 -> 本文任务/研究问题 -> 方法或识别设计 -> 数据/实验设置 -> 主要发现或可确认结果 -> 边际贡献或应用含义。",
            "- 句法风格：少用口号式形容词，多用问题驱动、机制解释、证据承接和贡献边界；每句话承载一个清晰功能。",
            "- 贡献表达：贡献必须落在任务重构、特征/数据、方法机制、评估场景或理论/实践含义上，不能泛泛写“具有重要意义”。",
            "- 结果边界：原文没有实验数值或显著性结论时，只写研究设计和预期用途，不编造发现。",
            "- 禁止写法：不要写“本文主要围绕、原文指出、该文档提到、具有一定参考意义”等模板句。",
            "- 范例使用：下方范例只用于学习摘要组织方式、节奏和学术语气，不能复制范例事实、术语组合或结论。",
        ]
    )


def summary_output_spec(
    summary_mode: str,
    academic: bool,
    material_style: str = "document",
    cn_target: int | None = None,
    en_target: int | None = None,
) -> str:
    """根据用户选择生成摘要输出语言、长度和格式要求。"""
    cn_clause = f"目标长度：{cn_target}个汉字。" if cn_target else ""
    en_clause = f"Target length: {en_target} words." if en_target else ""
    # PPT/图片材料通常没有完整论文结构，摘要要求必须更保守。
    if material_style in {"presentation", "visual"}:
        material_name = "PPT演示材料" if material_style == "presentation" else "图片/视觉材料"
        if summary_mode == "英文摘要":
            return (
                f"Output an English summary for a {material_name}. Use four plain-text headings: "
                "Topic, Structure, Key Points, Reading Notes. Do not force conclusions, risks, or recommended actions "
                f"unless the source explicitly provides them. Total length: 140-220 words. {en_clause}"
            )
        if summary_mode == "双语摘要":
            return (
                f"先输出一段自然连贯的中文{material_name}摘要，随后另起一段以 English Summary 开头输出英文摘要。"
                f"不要使用小标题、列表、箭头、加号或模板化栏目；不要强行补写实验结论、风险或行动建议。{cn_clause} {en_clause}"
            )
        return (
            f"输出220-380个汉字的中文{material_name}摘要，使用材料主题、内容结构、关键要点、阅读提示四个纯文本小标题。"
            f"不要强行套用重要结论、风险或建议行动；原文没有的效果判断、实验数值和适用边界不要补充。{cn_clause}"
        )
    # 正文类材料再区分学术写作和通用业务摘要。
    if summary_mode == "英文摘要":
        if academic:
            return (
                "Output one top-journal/top-conference style research abstract, 180-260 words. "
                "Cover the motivation or gap, research question, method/design, data or evaluation setting, "
                f"main supported findings, contribution, and implications. Do not use headings, bullet points, Markdown, or explanations. {en_clause}"
            )
        return (
            "Output an English business-style summary with five plain-text headings: Overview, Key Information, "
            f"Main Conclusions, Risks or Caveats, Recommended Actions. Total length: 180-260 words. {en_clause}"
        )
    if summary_mode == "双语摘要":
        if academic:
            return (
                "先输出一段中文顶刊/顶会风格研究摘要，约300-500个汉字；随后另起一段以 English Summary 开头输出一段 English top-journal/top-conference style abstract, "
                f"about 160-240 words. Use plain text only, no headings except the English Summary marker, no Markdown, no bullet points. {cn_clause} {en_clause}"
            )
        return (
            f"先输出一段自然连贯的中文摘要，随后另起一段以 English Summary 开头输出英文摘要。"
            f"中文和英文都要覆盖材料主题、关键信息、可确认结论、风险或注意事项；不要使用核心概述、关键信息、建议行动等栏目标题。{cn_clause} {en_clause}"
        )
    if academic:
        return f"只输出一段中文顶刊/顶会风格研究摘要，300-500个汉字，不要标题、列表、Markdown或解释。{cn_clause}"
    return f"输出300-500个汉字的五段式中文业务摘要，使用核心概述、关键信息、重要结论、风险或注意事项、建议行动五个纯文本小标题。{cn_clause}"
