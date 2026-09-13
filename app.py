# -*- coding: utf-8 -*-
import streamlit as st
import pdfplumber
import re
import json
import time
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from dashscope import Generation
import os

AI_DISCLAIMER = (
    "本报告由人工智能辅助生成，仅用于辅助识别待核实线索，"
    "不构成审计意见，不构成对任何公司的评价。"
)

# ================= 1. 字体注册 =================
def register_chinese_font():
    candidates = [
        ('/System/Library/Fonts/PingFang.ttc', 0),
        ('/System/Library/Fonts/STHeiti Light.ttc', 0),
        ('/System/Library/Fonts/STHeiti Medium.ttc', 0),
        ('/System/Library/Fonts/Supplemental/Songti.ttc', 0),
        ('/System/Library/Fonts/Hiragino Sans GB.ttc', 0),
        ('/Library/Fonts/Arial Unicode.ttf', None),
        ('C:/Windows/Fonts/msyh.ttc', 0),
        ('C:/Windows/Fonts/simhei.ttf', None),
        ('C:/Windows/Fonts/simsun.ttc', 0),
        ('/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc', 0),
        ('/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc', 0),
        ('/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf', None),
    ]
    local_font = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fonts/SimHei.ttf')
    candidates.append((local_font, None))

    for path, subfont_idx in candidates:
        if os.path.exists(path):
            try:
                if subfont_idx is not None:
                    pdfmetrics.registerFont(TTFont('ChineseFont', path, subfontIndex=subfont_idx))
                else:
                    pdfmetrics.registerFont(TTFont('ChineseFont', path))
                return 'ChineseFont'
            except Exception:
                continue
    return None

FONT_NAME = register_chinese_font()

def replace_sensitive_words(text):
    if not isinstance(text, str):
        return text
    replacements = {
        "虚增": "待核实的错报风险线索", "造假": "待核实的错报风险线索",
        "舞弊": "待核实的错报风险线索", "财务造假": "待核实的错报风险线索",
        "应发表保留意见": "建议实施进一步审计程序", "非标意见": "建议实施进一步审计程序后由注册会计师判断",
        "持续经营危机": "可能与持续经营相关的事项", "高风险公司": "本样本识别出待核实线索",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text

def parse_json_result(text):
    if not text:
        return None
    for pattern in [r'\[.*?\]', r'\{.*?\}']:
        for m in re.finditer(pattern, text, re.DOTALL):
            try:
                return json.loads(m.group(0))
            except Exception:
                continue
    for pattern in [r'\[.*\]', r'\{.*\}']:
        m = re.search(pattern, text, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                continue
    return None

def call_llm(prompt, api_key, model_name="qwen-long", max_retries=3):
    for attempt in range(max_retries):
        try:
            response = Generation.call(
                api_key=api_key, model=model_name,
                messages=[
                    {"role": "system", "content": "你是严谨的审计AI助手，只输出JSON，不要任何多余文字。"},
                    {"role": "user", "content": prompt}
                ],
                result_format='message', timeout=120,
            )
            if response.status_code == 200:
                return response.output.choices[0].message.content
            if 'InvalidApiKey' in str(response.code) or 'AccessDenied' in str(response.code):
                return None
            time.sleep(3)
        except Exception:
            time.sleep(3)
    return None

def compute_indicators(pdf_text):
    pages = re.split(r'(?=【第\d+页】)', pdf_text)
    head_text = "\n".join(pages[:40])
    indicators = {"营业收入": None, "净利润": None, "经营活动现金流量净额": None, "应收账款": None}
    flags = []
    patterns = {
        "营业收入": r"营业收入[^\n\d\-]{0,20}?([\-\d][\d,\.]*)",
        "净利润": r"净利润[^\n\d\-]{0,20}?([\-\d][\d,\.]*)",
        "经营活动现金流量净额": r"经营活动产生的现金流量净额[^\n\d\-]{0,20}?([\-\d][\d,\.]*)",
        "应收账款": r"应收账款[^\n\d\-]{0,20}?([\-\d][\d,\.]*)",
    }
    for key, pattern in patterns.items():
        m = re.search(pattern, head_text)
        if m:
            try:
                indicators[key] = float(m.group(1).replace(",", ""))
            except ValueError:
                pass
    if indicators["营业收入"] and indicators["净利润"]:
        indicators["净利率"] = round(indicators["净利润"] / indicators["营业收入"], 4)
    if indicators["净利润"] and indicators["净利润"] != 0 and indicators["经营活动现金流量净额"] is not None:
        r = indicators["经营活动现金流量净额"] / indicators["净利润"]
        indicators["收现比"] = round(r, 4)
        if r < 0.3:
            flags.append(f"收现比 {r:.2%} 低于 30% 阈值，需到附注找解释")
    return indicators, flags

def split_annual_report(pdf_text):
    sections = {"财务报表主表": [], "会计政策附注": [], "关联方担保诉讼": [], "管理层讨论": [], "其他": []}
    pages = re.split(r'(?=【第\d+页】)', pdf_text)
    keywords = {
        "财务报表主表": r"合并资产负债表|合并利润表|合并现金流量表|母公司资产负债表|合并所有者权益变动表",
        "会计政策附注": r"重要会计政策|会计估计|会计政策变更|前期差错更正",
        "关联方担保诉讼": r"关联方|对外担保|重大诉讼|或有事项|承诺事项",
        "管理层讨论": r"管理层讨论与分析|经营情况讨论与分析|董事会报告",
    }
    for page in pages:
        if not page.strip():
            continue
        matched = False
        for section_name, pattern in keywords.items():
            if re.search(pattern, page):
                sections[section_name].append(page)
                matched = True
                break
        if not matched:
            sections["其他"].append(page)
    return {k: "\n".join(v) for k, v in sections.items()}

def condense_section(section_text, api_key, model_name, section_name):
    pages = [p for p in re.split(r'(?=【第\d+页】)', section_text) if p.strip()]
    if len(pages) <= 100:
        return section_text
    batch_size = 50
    summaries = []
    for i in range(0, len(pages), batch_size):
        batch = "\n".join(pages[i:i + batch_size])
        prompt = f"""下面是年报【{section_name}】章节的第 {i+1} 到 {min(i+batch_size, len(pages))} 页。
请对每一页生成一句话摘要，格式严格如下，不要任何多余文字：
第n页：xxx
不要合并页，不要遗漏任何一页。

{batch}"""
        s = call_llm(prompt, api_key, model_name)
        summaries.append(s if s else batch)
    return "\n".join(summaries)

PROMPTS = {
    "FA": """你是审计团队的【财务科目审计员】。
【输入】下面是 Python 已算好的指标表。你只做一件事：对每个超阈值的指标，去原文里找有没有合理解释。
【规则】
- 找到解释 → 引用原文 + 页码
- 找不到 → 写"未发现合理解释，建议实施进一步程序"
- 禁止自己算同比/比率，禁止下审计意见
【输出】严格输出 JSON 数组，不要任何多余文字：
[{{"指标名称":"","本期数":"","上期数":"","变动":"","是否超阈值":"","解释原文":"","页码":"","建议程序":""}}]

【Python计算的指标】: {indicators}
【异常Flags】: {flags}
【年报原文（含页码）】: {text}
""",
    "DA": """你是审计团队的【报表披露审计员】。
【任务】对照准则查披露完整性：
- 关联方看《企业会计准则第36号》
- 担保、诉讼看《企业会计准则第13号》
- 报表列报看《企业会计准则第30号》
【规则】每条输出固定格式：准则简称 + 章节 + 缺了什么 + 页码
【输出】严格输出 JSON 数组，不要任何多余文字：
[{{"准则":"","章节":"","缺失内容":"","原文证据":"","页码":"","建议程序":""}}]

【关联方担保诉讼原文】: {text1}
【会计政策附注原文】: {text2}
""",
    "SA": """你是审计团队的【专项审计员】。
【任务】只检查：MD&A 与报表是否打架；关联方、担保、诉讼是否完整。
【输出】严格输出 JSON 数组，不要任何多余文字：
[{{"事项类型":"","年报说法":"","报表数字":"","是否矛盾":"","页码":"","建议程序":""}}]

【管理层讨论原文】: {text1}
【关联方担保诉讼原文】: {text2}
""",
    "PM": """你是一个项目经理。请阅读下面三位专员的发现，用不超过150字总结这份年报的主要风险特征。只输出一段纯文本，不要任何JSON格式，不要任何Markdown标记。

【FA发现】: {fa}
【DA发现】: {da}
【SA发现】: {sa}
"""
}

def run_audit_analysis(pdf_text, api_key, model_name="qwen-long"):
    indicators, flags = compute_indicators(pdf_text)
    sections = split_annual_report(pdf_text)
    specialist_results = {}
    raw_results = {}

    try:
        fa_text = condense_section(sections.get("财务报表主表", "") + "\n" + sections.get("会计政策附注", ""), api_key, model_name, "财务报表与会计政策")
        prompt = PROMPTS["FA"].format(indicators=json.dumps(indicators, ensure_ascii=False), flags=flags, text=fa_text)
        raw_results["FA"] = call_llm(prompt, api_key, model_name)
        parsed = parse_json_result(raw_results["FA"])
        specialist_results["FA"] = parsed if isinstance(parsed, list) else []
    except Exception as e:
        st.warning(f"财务科目审计员失败：{e}")
        specialist_results["FA"] = []

    try:
        da_text1 = condense_section(sections.get("关联方担保诉讼", ""), api_key, model_name, "关联方担保诉讼")
        da_text2 = condense_section(sections.get("会计政策附注", ""), api_key, model_name, "会计政策附注")
        prompt = PROMPTS["DA"].format(text1=da_text1, text2=da_text2)
        raw_results["DA"] = call_llm(prompt, api_key, model_name)
        parsed = parse_json_result(raw_results["DA"])
        specialist_results["DA"] = parsed if isinstance(parsed, list) else []
    except Exception as e:
        st.warning(f"报表披露审计员失败：{e}")
        specialist_results["DA"] = []

    try:
        sa_text1 = condense_section(sections.get("管理层讨论", ""), api_key, model_name, "管理层讨论")
        sa_text2 = condense_section(sections.get("关联方担保诉讼", ""), api_key, model_name, "关联方担保诉讼")
        prompt = PROMPTS["SA"].format(text1=sa_text1, text2=sa_text2)
        raw_results["SA"] = call_llm(prompt, api_key, model_name)
        parsed = parse_json_result(raw_results["SA"])
        specialist_results["SA"] = parsed if isinstance(parsed, list) else []
    except Exception as e:
        st.warning(f"专项审计员失败：{e}")
        specialist_results["SA"] = []

    # ======== 纯 Python 兜底合并，绝对不报错 ========
    merged_clues = []
    for role in ["FA", "DA", "SA"]:
        for item in specialist_results.get(role, []):
            fact = item.get("解释原文") or item.get("缺失内容") or item.get("年报说法") or item.get("事实") or ""
            if "未发现合理解释" in fact:
                fact = "相关指标异常，但在报表附注中未找到合理解释。"
            evidence = item.get("原文证据", "")
            if not fact and not evidence:
                continue
            merged_clues.append({
                "编号": "A-" + str(len(merged_clues) + 1).zfill(2),
                "来源角色": role,
                "事项类型": item.get("指标名称") or item.get("准则") or item.get("事项类型", ""),
                "事实": fact or "无具体事实",
                "原文证据": evidence or "无",
                "页码": item.get("页码", "") or "无",
                "涉及认定": "完整性/准确性",
                "风险等级": "中",
                "是否交叉验证": "是" if len(specialist_results.get(role, [])) > 1 else "否",
                "建议程序": item.get("建议程序", "") or "建议实施进一步程序",
                "AI声明": "本报告由AI辅助生成，不构成审计意见",
            })

    # 让项目经理写一段纯文字评价即可
    pm_prompt = PROMPTS["PM"].format(
        fa=json.dumps(specialist_results.get("FA", []), ensure_ascii=False),
        da=json.dumps(specialist_results.get("DA", []), ensure_ascii=False),
        sa=json.dumps(specialist_results.get("SA", []), ensure_ascii=False)
    )
    raw_pm = call_llm(pm_prompt, api_key, model_name)
    overall_eval = raw_pm if raw_pm else "本样本经多维度分析，识别出若干待核实线索，供人工复核。"
    
    final_report = {
        "总体评价": overall_eval,
        "线索列表": merged_clues[:15]
    }
    raw_results["PM"] = raw_pm

    return final_report, raw_results, sections

def export_to_excel_json(final_json, raw_results, sections):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary_data = {"项目": ["总体评价", "AI声明", "生成时间"], "内容": [final_json.get("总体评价", "无"), AI_DISCLAIMER, datetime.now().strftime("%Y-%m-%d %H:%M")]}
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='最终报告', index=False)
        clues = final_json.get("线索列表", [])
        fixed_cols = ["编号", "来源角色", "事项类型", "事实", "原文证据", "页码", "涉及认定", "风险等级", "是否交叉验证", "建议程序"]
        rows = []
        for c in clues:
            row = {col: ("、".join(map(str, c.get(col, ""))) if isinstance(c.get(col, ""), list) else c.get(col, "")) for col in fixed_cols}
            rows.append(row)
        if rows:
            pd.DataFrame(rows, columns=fixed_cols).to_excel(writer, sheet_name='线索明细', index=False)
        else:
            pd.DataFrame({"提示": ["本次未识别出待核实线索"]}).to_excel(writer, sheet_name='线索明细', index=False)
        raw_data = [{"角色": k, "原始输出": v} for k, v in raw_results.items()]
        pd.DataFrame(raw_data).to_excel(writer, sheet_name='专员原始JSON', index=False)
        workbook = writer.book
        header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for sheet in workbook.worksheets:
            for cell in sheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center')
    output.seek(0)
    return output

# ================= 2. PDF 生成（最稳写法） =================
def export_to_pdf_json(final_json, raw_results):
    if FONT_NAME is None:
        return None, "未找到可用的中文字体，无法生成PDF。请将任意中文.ttf字体文件命名为 fonts/SimHei.ttf 放到程序同目录。"

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=2*cm, rightMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm)

    style_title = ParagraphStyle('Title', fontName=FONT_NAME, fontSize=18, leading=24, alignment=TA_CENTER, spaceAfter=20)
    style_heading = ParagraphStyle('Heading', fontName=FONT_NAME, fontSize=13, leading=18, spaceBefore=14, spaceAfter=8, textColor=colors.HexColor('#2F5496'))
    style_body = ParagraphStyle('Body', fontName=FONT_NAME, fontSize=10, leading=16, spaceBefore=3, spaceAfter=3)
    style_subtitle = ParagraphStyle('Subtitle', fontName=FONT_NAME, fontSize=9, leading=12, alignment=TA_CENTER, textColor=colors.grey, spaceAfter=20)

    elements = []
    elements.append(Paragraph('AI辅助审计分析报告', style_title))
    elements.append(Paragraph(AI_DISCLAIMER, style_subtitle))
    elements.append(Paragraph(f'生成时间：{datetime.now().strftime("%Y-%m-%d %H:%M")}', style_subtitle))
    elements.append(Spacer(1, 10))
    
    elements.append(Paragraph('【总体评价】', style_heading))
    elements.append(Paragraph(str(final_json.get("总体评价", "无")), style_body))

    elements.append(Paragraph('【待核实线索明细】', style_heading))
    clues = final_json.get("线索列表", [])
    if not clues:
        elements.append(Paragraph("本次未识别出待核实线索。", style_body))
    
    for clue in clues:
        header = f"编号: {clue.get('编号','')} | 风险等级: {clue.get('风险等级','')} | 涉及认定: {clue.get('涉及认定','')}"
        elements.append(Paragraph(header, style_heading))
        for label, key in [("事实", "事实"), ("原文证据", "原文证据"), ("建议程序", "建议程序"), ("页码", "页码")]:
            val = replace_sensitive_words(str(clue.get(key, '')))
            if val:
                safe_val = val.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                elements.append(Paragraph(f"{label}: {safe_val}", style_body))
        elements.append(Spacer(1, 10))

    elements.append(PageBreak())
    elements.append(Paragraph('附录：专员原始输出', style_heading))
    for role, text in raw_results.items():
        elements.append(Spacer(1, 8))
        elements.append(Paragraph(f'【{role}】', style_heading))
        for line in str(text).split('\n'):
            line = line.strip()
            if line:
                safe_line = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                elements.append(Paragraph(safe_line, style_body))

    try:
        doc.build(elements)
        return buffer.getvalue(), None
    except Exception as e:
        return None, f"PDF生成失败：{str(e)}"

# ================= 3. Streamlit UI =================
st.set_page_config(page_title="AI审计辅助系统", layout="wide")
st.title("📊 AI 审计辅助系统（年报待核实线索识别）")
st.caption(AI_DISCLAIMER)

with st.sidebar:
    st.header("配置")
    api_key = st.text_input("DashScope API Key", type="password", help="阿里云百炼平台获取")
    st.caption("模型锁定：qwen-long（长文本必需）")
    st.divider()
    st.caption("架构：Python算指标 → 分块 → FA/DA/SA/PM 四角色 → JSON → Excel/PDF")

uploaded_file = st.file_uploader("上传年报 PDF", type=["pdf"])

if uploaded_file and api_key:
    if st.button("🚀 启动审计分析", type="primary"):
        with st.spinner("正在解析 PDF..."):
            pdf_text = ""
            try:
                with pdfplumber.open(uploaded_file) as pdf:
                    for i, page in enumerate(pdf.pages):
                        t = page.extract_text()
                        if t:
                            pdf_text += f"【第{i+1}页】\n{t}\n\n"
            except Exception as e:
                st.error(f"PDF解析失败：{e}")
                st.stop()

        if len(pdf_text.strip()) < 100:
            st.error("未提取到有效文字，可能是扫描件，请先 OCR 后重试。")
            st.stop()

        t0 = time.time()
        with st.spinner("大模型正在交叉核对（约 1-3 分钟）..."):
            final_report, raw_results, sections = run_audit_analysis(pdf_text, api_key)
        elapsed = round(time.time() - t0, 1)

        st.session_state['final_report'] = final_report
        st.session_state['raw_results'] = raw_results
        st.session_state['sections'] = sections
        st.session_state['pdf_text'] = pdf_text
        st.session_state['elapsed'] = elapsed

if st.session_state.get('final_report'):
    report = st.session_state['final_report']
    sections = st.session_state.get('sections', {})
    pdf_text = st.session_state.get('pdf_text', '')
    elapsed = st.session_state.get('elapsed', 0)

    total_pages = pdf_text.count('【第')
    sec_pages = {k: v.count('【第') for k, v in sections.items()}
    st.caption(
        f"年报共 {total_pages} 页，"
        f"财务报表主表 {sec_pages.get('财务报表主表', 0)} 页、"
        f"会计政策附注 {sec_pages.get('会计政策附注', 0)} 页、"
        f"关联方担保诉讼 {sec_pages.get('关联方担保诉讼', 0)} 页、"
        f"管理层讨论 {sec_pages.get('管理层讨论', 0)} 页、"
        f"其他 {sec_pages.get('其他', 0)} 页，全部命中页已纳入分析，未做截断。"
        f"总耗时 {elapsed} 秒。"
    )

    st.subheader("📝 总体评价")
    st.info(report.get("总体评价", "无"))

    st.subheader("🔍 待核实线索列表")
    clues = report.get("线索列表", [])
    if not clues:
        st.warning("本次未识别出待核实线索。可检查 PDF 是否为扫描件，或更换年报重试。")
    for clue in clues:
        title = f"[{clue.get('风险等级','')}] {clue.get('编号','')} - {clue.get('事项类型','')}"
        with st.expander(title):
            st.write(f"**事实：** {clue.get('事实','')}")
            st.write(f"**原文证据：** {clue.get('原文证据','')}")
            st.write(f"**涉及认定：** {clue.get('涉及认定','')}")
            st.write(f"**是否交叉验证：** {clue.get('是否交叉验证','')}")
            st.write(f"**建议程序：** {clue.get('建议程序','')}")
            st.write(f"**页码：** {clue.get('页码','')}")

    st.divider()
    st.subheader("📥 下载报告")
    col1, col2 = st.columns(2)
    with col1:
        try:
            excel_data = export_to_excel_json(report, st.session_state.get('raw_results', {}), st.session_state.get('sections', {}))
            st.download_button("下载 Excel 报告", excel_data, f"AI审计报告_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Excel 生成失败：{e}")
    with col2:
        pdf_data, pdf_error = export_to_pdf_json(report, st.session_state.get('raw_results', {}))
        if pdf_error:
            st.warning(pdf_error)
        else:
            st.download_button("下载 PDF 报告", pdf_data, f"AI审计报告_{datetime.now().strftime('%Y%m%d')}.pdf", mime="application/pdf")
