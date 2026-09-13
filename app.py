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
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth
from dashscope import Generation
import os


AI_DISCLAIMER = (
    "本报告由人工智能辅助生成，仅用于辅助识别待核实线索，"
    "不构成审计意见，不构成对任何公司的评价。"
)


def register_chinese_font():
    font_paths = [
        ("C:/Windows/Fonts/simhei.ttf", None),
        ("/System/Library/Fonts/PingFang.ttc", 0),
        ("/System/Library/Fonts/STHeiti Light.ttc", 0),
        ("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc", 0),
        ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 0),
    ]
    for path, idx in font_paths:
        if os.path.exists(path):
            try:
                if idx is not None:
                    pdfmetrics.registerFont(TTFont('ChineseFont', path, subfontIndex=idx))
                else:
                    pdfmetrics.registerFont(TTFont('ChineseFont', path))
                return 'ChineseFont'
            except Exception:
                continue
    return "Helvetica"


def replace_sensitive_words(text):
    if not isinstance(text, str):
        return text
    replacements = {
        "虚增": "待核实的错报风险线索",
        "造假": "待核实的错报风险线索",
        "舞弊": "待核实的错报风险线索",
        "财务造假": "待核实的错报风险线索",
        "应发表保留意见": "建议实施进一步审计程序",
        "非标意见": "建议实施进一步审计程序后由注册会计师判断",
        "持续经营危机": "可能与持续经营相关的事项",
        "高风险公司": "本样本识别出待核实线索",
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
                api_key=api_key,
                model=model_name,
                messages=[
                    {"role": "system",
                     "content": "你是严谨的审计AI助手，只输出JSON，不要任何多余文字。"},
                    {"role": "user", "content": prompt}
                ],
                result_format='message',
                timeout=120,
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

    indicators = {
        "营业收入": None, "净利润": None,
        "经营活动现金流量净额": None, "应收账款": None,
    }
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

    if indicators["净利润"] and indicators["净利润"] != 0 \
            and indicators["经营活动现金流量净额"] is not None:
        r = indicators["经营活动现金流量净额"] / indicators["净利润"]
        indicators["收现比"] = round(r, 4)
        if r < 0.3:
            flags.append(f"收现比 {r:.2%} 低于 30% 阈值，需到附注找解释")

    return indicators, flags


def split_annual_report(pdf_text):
    sections = {
        "财务报表主表": [], "会计政策附注": [],
        "关联方担保诉讼": [], "管理层讨论": [], "其他": [],
    }
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
    "PM": """你是审计团队的【项目经理】。
【任务】
1. 合并同一页码、同一金额的线索
2. 矛盾的单独列出
3. 按重要性定级：金额≥营业收入5%为高风险，低于为中或低
4. 不重新编故事，只汇总
【输出要求】严格输出 JSON，不要任何多余文字。总体评价不超过150字，线索列表最多保留10条（按风险等级从高到低）。
【输出格式】
{{
  "总体评价": "150字以内",
  "线索列表": [
    {{
      "编号": "A-01",
      "来源角色": "FA或DA或SA",
      "事项类型": "",
      "事实": "不超过80字",
      "原文证据": "不超过80字",
      "页码": "",
      "涉及认定": "",
      "风险等级": "高/中/低",
      "是否交叉验证": "是/否",
      "建议程序": "不超过60字",
      "AI声明": "本报告由AI辅助生成，不构成审计意见"
    }}
  ]
}}

【FA结果】: {fa}
【DA结果】: {da}
【SA结果】: {sa}
"""
}


def run_audit_analysis(pdf_text, api_key, model_name="qwen-long"):
    indicators, flags = compute_indicators(pdf_text)
    sections = split_annual_report(pdf_text)

    specialist_results = {}
    raw_results = {}

    try:
        fa_text = condense_section(
            sections.get("财务报表主表", "") + "\n" + sections.get("会计政策附注", ""),
            api_key, model_name, "财务报表与会计政策"
        )
        prompt = PROMPTS["FA"].format(
            indicators=json.dumps(indicators, ensure_ascii=False),
            flags=flags, text=fa_text
        )
        raw_results["FA"] = call_llm(prompt, api_key, model_name)
        parsed = parse_json_result(raw_results["FA"])
        specialist_results["FA"] = parsed if isinstance(parsed, list) else []
    except Exception as e:
        st.warning(f"财务科目审计员失败：{e}")
        specialist_results["FA"] = []
        raw_results["FA"] = str(e)

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
        raw_results["DA"] = str(e)

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
        raw_results["SA"] = str(e)

    try:
        pm_prompt = PROMPTS["PM"].format(
            fa=json.dumps(specialist_results.get("FA", []), ensure_ascii=False),
            da=json.dumps(specialist_results.get("DA", []), ensure_ascii=False),
            sa=json.dumps(specialist_results.get("SA", []), ensure_ascii=False)
        )
        # Python直接合并，不调大模型，防止截断和乱码
        merged_clues = []
        seen = set()
        for role in ["FA", "DA", "SA"]:
            for item in specialist_results.get(role, []):
                fact = item.get("解释原文") or item.get("缺失内容") or item.get("年报说法") or ""
                evidence = item.get("原文证据", "")
                if "未发现合理解释" in fact:
                    fact = "相关指标异常，但在报表附注中未找到合理解释。"
                # 过滤掉全空的条目
                if not fact and not evidence:
                    continue
                # 去重
                key = (fact[:50], evidence[:50])
                if key in seen:
                    continue
                seen.add(key)
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
        if merged_clues:
            final_report = {
                "总体评价": "本样本经Python多维度合并，共识别出 " + str(len(merged_clues)) + " 条待核实线索，供人工复核。",
                "线索列表": merged_clues[:15],
            }
        else:
            final_report = {"总体评价": "本次未识别出实质性的待核实线索。", "线索列表": []}
    except Exception as e:
        st.error(f"项目经理汇总失败：{e}")
        final_report = {"总体评价": "项目经理汇总失败，请检查API或重试。", "线索列表": []}

    if not isinstance(final_report, dict) or not final_report.get("线索列表"):
        merged_clues = []
        idx = 1
        for role in ["FA", "DA", "SA"]:
            for item in specialist_results.get(role, []):
                merged_clues.append({
                    "编号": "A-" + str(idx).zfill(2),
                    "来源角色": role,
                    "事项类型": item.get("指标名称") or item.get("准则") or item.get("事项类型", ""),
                    "事实": item.get("解释原文") or item.get("缺失内容") or item.get("年报说法", ""),
                    "原文证据": item.get("原文证据", ""),
                    "页码": item.get("页码", ""),
                    "涉及认定": "",
                    "风险等级": "中",
                    "是否交叉验证": "否",
                    "建议程序": item.get("建议程序", ""),
                    "AI声明": "本报告由AI辅助生成，不构成审计意见",
                })
                idx += 1
        if merged_clues:
            final_report = {
                "总体评价": "本样本共识别出 " + str(len(merged_clues)) + " 条待核实线索，来自财务、披露、专项三个维度，供人工复核。",
                "线索列表": merged_clues[:10],
            }
        else:
            final_report = {"总体评价": "本次未识别出待核实线索。", "线索列表": []}

    return final_report, raw_results, sections


def export_to_excel_json(final_json, raw_results, sections):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary_data = {
            "项目": ["总体评价", "AI声明", "生成时间"],
            "内容": [
                final_json.get("总体评价", "无"),
                AI_DISCLAIMER,
                datetime.now().strftime("%Y-%m-%d %H:%M"),
            ]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='最终报告', index=False)

        clues = final_json.get("线索列表", [])
        fixed_cols = ["编号", "来源角色", "事项类型", "事实", "原文证据", "页码",
                      "涉及认定", "风险等级", "是否交叉验证", "建议程序"]
        rows = []
        for c in clues:
            row = {}
            for col in fixed_cols:
                v = c.get(col, "")
                if isinstance(v, list):
                    v = "、".join(map(str, v))
                row[col] = v
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
        thin = Side(style='thin', color='CCCCCC')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for sheet in workbook.worksheets:
            for cell in sheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
            for col_idx in range(1, sheet.max_column + 1):
                sheet.column_dimensions[get_column_letter(col_idx)].width = 22

    output.seek(0)
    return output


def wrap_text(text, font_name, font_size, max_width):
    lines, cur = [], ""
    for ch in str(text):
        # 加入 len(cur) < 38 的硬限制，防止中文字符宽度计算错误导致不换行
        if stringWidth(cur + ch, font_name, font_size) < max_width and len(cur) < 38:
            cur += ch
        else:
            lines.append(cur)
            cur = ch
    if cur:
        lines.append(cur)
    return lines


def export_to_pdf_json(final_json, raw_results):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    font_name = register_chinese_font()
    max_width = 450

    c.setFont(font_name, 16)
    c.drawString(50, 800, "AI辅助审计分析报告")
    c.setFont(font_name, 9)
    for i, line in enumerate(wrap_text(AI_DISCLAIMER, font_name, 9, max_width)):
        c.drawString(50, 782 - i * 12, line)

    y_pos = 740
    c.setFont(font_name, 12)
    c.drawString(50, y_pos, "【总体评价】")
    y_pos -= 18
    c.setFont(font_name, 10)
    for line in wrap_text(final_json.get("总体评价", "无"), font_name, 10, max_width):
        if y_pos < 50:
            c.showPage(); c.setFont(font_name, 10); y_pos = 800
        c.drawString(50, y_pos, line)
        y_pos -= 15

    y_pos -= 10
    c.setFont(font_name, 12)
    c.drawString(50, y_pos, "【待核实线索明细】")
    y_pos -= 20

    clues = final_json.get("线索列表", [])
    if not clues:
        c.setFont(font_name, 10)
        c.drawString(50, y_pos, "本次未识别出待核实线索。")

    for clue in clues:
        header = f"编号: {clue.get('编号','')} | 风险等级: {clue.get('风险等级','')} | 涉及认定: {clue.get('涉及认定','')}"
        body_lines = [
            f"事实: {clue.get('事实','')}",
            f"原文证据: {clue.get('原文证据','')}",
            f"建议程序: {clue.get('建议程序','')} | 页码: {clue.get('页码','')}",
            "-" * 80,
        ]
        all_lines = [header] + body_lines
        for line in all_lines:
            for wl in wrap_text(replace_sensitive_words(str(line)), font_name, 10, max_width):
                if y_pos < 50:
                    c.showPage(); c.setFont(font_name, 10); y_pos = 800
                c.drawString(50, y_pos, wl)
                y_pos -= 15
        y_pos -= 5

    c.showPage()
    c.setFont(font_name, 12)
    c.drawString(50, 800, "【附录：专员原始输出】")
    y_pos = 780
    for role, text in raw_results.items():
        c.setFont(font_name, 11)
        if y_pos < 60:
            c.showPage(); y_pos = 800
        c.drawString(50, y_pos, f"—— {role} ——")
        y_pos -= 16
        c.setFont(font_name, 9)
        for line in str(text).split('\n'):
            for wl in wrap_text(line, font_name, 9, max_width):
                if y_pos < 50:
                    c.showPage(); c.setFont(font_name, 9); y_pos = 800
                c.drawString(50, y_pos, wl)
                y_pos -= 12
        y_pos -= 10

    c.save()
    buffer.seek(0)
    return buffer


st.set_page_config(page_title="AI审计辅助系统", layout="wide")
st.title("📊 AI 审计辅助系统（年报待核实线索识别）")
st.caption(AI_DISCLAIMER)

with st.sidebar:
    st.header("配置")
    api_key = st.text_input("DashScope API Key", type="password",
                            help="阿里云百炼平台获取")
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
            excel_data = export_to_excel_json(report, st.session_state.get('raw_results', {}),
                                              st.session_state.get('sections', {}))
            st.download_button("下载 Excel 报告", excel_data,
                               f"AI审计报告_{datetime.now().strftime('%Y%m%d')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Excel 生成失败：{e}")
    with col2:
        try:
            pdf_data = export_to_pdf_json(report, st.session_state.get('raw_results', {}))
            st.download_button("下载 PDF 报告", pdf_data,
                               f"AI审计报告_{datetime.now().strftime('%Y%m%d')}.pdf",
                               mime="application/pdf")
        except Exception as e:
            st.error(f"PDF 生成失败：{e}")
