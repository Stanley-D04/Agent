#!/usr/bin/env python3
"""
用户反馈分析报告生成器
用法：
    python3 gen_report.py <excel_path> [output_path]

Excel 必须包含以下列（列名可通过 COLUMN_MAP 配置）：
    - 时间（提交时间）
    - 品类
    - 评价（非常不满意 / 一般 / 满意）
    - 问卷环节
    - 描述（用户反馈文字）
    - 标签（多标签用 | 分隔）
    - 图片（可选，多图用 | 分隔，URL）
    - 预估价（可选）
    - 质检价（可选）
    - 评分（可选，数字）

输出：单个自包含 HTML 文件，含完整交互功能。
"""

import sys
import json
import re
import math
from collections import Counter, defaultdict
from pathlib import Path

# ── 依赖检查 ──────────────────────────────────────────────────────────────
try:
    import openpyxl
except ImportError:
    print("正在安装依赖 openpyxl ...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "-q"])
    import openpyxl

try:
    import pandas as pd
except ImportError:
    print("正在安装依赖 pandas ...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas", "openpyxl", "-q"])
    import pandas as pd


# ── 列名映射（根据实际 Excel 表头修改） ──────────────────────────────────────
# key: 脚本内部字段名，value: Excel 列名（支持多个候选名，取第一个匹配的）
COLUMN_MAP = {
    "time":        ["提交时间", "时间", "日期", "date", "time"],
    "cat":         ["品类", "category", "产品品类"],
    "eval_type":   ["评价", "满意度", "用户评价", "evaluation"],
    "survey":      ["问卷环节", "环节", "survey", "问卷阶段"],
    "desc":        ["描述", "用户描述", "反馈内容", "description", "content", "用户反馈"],
    "tags":        ["标签", "问题标签", "tag", "tags"],
    "img":         ["图片", "图片url", "image", "img", "图片链接"],
    "est_price":   ["预估价", "estimated_price", "预估价格"],
    "check_price": ["质检价", "check_price", "质检价格", "实际价格"],
    "score":       ["评分", "score", "星级"],
}

# ── 标签 → 分类映射（可扩展） ────────────────────────────────────────────────
TAG_CLASS = {
    "缺少品牌":  "A",
    "缺少型号":  "A",
    "缺少商品":  "A",
    "缺少品类":  "A",
    "缺少实拍图":"A",
    "价格不满意":"B",
    "质检结果存疑":"C",
    "催促验机/物流":"D",
    "服务质量":  "D",
    "联系不上工程师":"D",
    "密码账号相关":"E",
    "不清楚如何操作":"E",
}

CLS_META = {
    "A": {"label": "品类/品牌/型号缺失",  "badge": "高频 · 核心痛点", "badge_cls": "high",
          "card_cls": "A", "desc": "用户在选机型环节找不到对应品牌、型号或品类，无法完成回收下单流程，是最直接的转化损失点。"},
    "B": {"label": "回收价格不满意",       "badge": "高频 · 强烈不满",  "badge_cls": "mid",
          "card_cls": "B", "desc": "用户认为回收定价偏低，或预估价与质检后实际报价落差过大，情绪强烈，差评率高。"},
    "C": {"label": "质检结果存疑",         "badge": "中频 · 信任危机",  "badge_cls": "low",
          "card_cls": "C", "desc": "用户对质检结论不认可，包括假货判定争议、损伤定级偏差、与下单型号不符等问题，严重影响用户信任。"},
    "D": {"label": "物流/服务问题",        "badge": "服务体验",         "badge_cls": "info",
          "card_cls": "D", "desc": "验机等待时间长、物流状态未更新、无法联系客服或工程师，影响用户整体服务体验。"},
    "E": {"label": "操作/功能问题",        "badge": "操作体验",         "badge_cls": "low",
          "card_cls": "C", "desc": "用户对操作流程不熟悉，遇到密码账号、功能缺失等问题导致流程受阻，无法顺利完成卖出。"},
    "MISC": {"label": "其他反馈",          "badge": "其他",             "badge_cls": "info",
          "card_cls": "D", "desc": "未归入特定分类的用户反馈。"},
}

EVAL_COLORS = {"非常不满意": "eval-bad", "一般": "eval-mid", "满意": "eval-good"}
SURVEY_COLORS = {"选机型": "survey-sel", "质检": "survey-check", "下单": "survey-order",
                 "服务": "survey-service", "估价": "survey-price"}

INSIGHT_TEMPLATES = {
    "A": {"level": "p1", "priority": "P0 · 最高优先级",
          "title": "品类/型号覆盖不足是第一道流失关卡",
          "body": "大量用户在选机型环节因找不到对应品牌/型号直接放弃，建议集中补录高频缺失SKU，并开放用户上报通道形成持续更新机制。"},
    "B": {"level": "p1", "priority": "P0 · 最高优先级",
          "title": "预估价与质检价落差是差评核心来源",
          "body": "用户普遍将预估价视为承诺价，质检后骤降触发强烈反弹。需在预估页加注「仅供参考」，并在降价幅度较大时主动说明原因。"},
    "C": {"level": "p2", "priority": "P1 · 重要",
          "title": "质检结论缺乏透明证据，信任损伤风险高",
          "body": "用户对假货判定和损伤定级提出质疑，建议在质检报告中增加图文证据，并为争议设立快速申诉通道。"},
    "D": {"level": "p2", "priority": "P1 · 重要",
          "title": "验机时效与服务响应是体验短板",
          "body": "多名用户反映等待过久或联系不上客服。建议设定验机时效SLA，超时主动推送进度通知，并优化客服入口。"},
    "E": {"level": "p3", "priority": "P2 · 优化",
          "title": "操作流程门槛偏高，用户容易卡壳",
          "body": "部分用户对操作步骤感到困惑或遇到账号问题。建议完善引导说明和帮助文档，降低操作门槛。"},
}

REC_TEMPLATES = {
    "A": {"title": "补录缺失品牌/型号",    "body": "梳理高频反馈中缺失的品牌和型号，优先录入系统；同时建立用户上报通道，持续扩大覆盖范围。"},
    "B": {"title": "优化预估价预期管理",    "body": "预估页明确标注「仅供参考」；质检后降价超过20%时触发系统通知，并附降价原因说明，减少用户预期落差。"},
    "C": {"title": "提升质检透明度与申诉机制", "body": "质检报告中增加图文依据；对假货判定设立申诉通道；支持用户上传证明材料参与复核，重建信任。"},
    "D": {"title": "设定服务时效SLA",       "body": "明确验机时效（如3-5工作日），超时自动推送进度；完善客服联系入口，保障响应效率。"},
    "E": {"title": "完善操作引导与帮助",    "body": "针对高频卡壳步骤增加操作说明和帮助文档；账号密码问题提供自助解决方案，减少流程阻力。"},
}

FILL_COLORS = ["fill-red", "fill-orange", "fill-blue", "fill-teal", "fill-gray",
               "fill-gray", "fill-gray", "fill-gray"]


# ── 数据读取 ─────────────────────────────────────────────────────────────────
def find_column(df_cols, candidates):
    for c in candidates:
        for col in df_cols:
            if str(col).strip().lower() == c.lower():
                return col
    return None


def load_excel(excel_path: str) -> list[dict]:
    """读取 Excel，返回标准化行列表"""
    df = pd.read_excel(excel_path, dtype=str)
    df = df.fillna("")

    col_mapping = {}
    for field, candidates in COLUMN_MAP.items():
        matched = find_column(df.columns.tolist(), candidates)
        col_mapping[field] = matched

    rows = []
    for _, raw in df.iterrows():
        row = {}
        for field, col in col_mapping.items():
            row[field] = str(raw[col]).strip() if col and col in raw else ""
        # 清洗
        row["tags"] = "|".join(t.strip() for t in row["tags"].replace("，", "|").replace(",", "|").split("|") if t.strip())
        rows.append(row)
    return rows


# ── HTML 生成工具 ─────────────────────────────────────────────────────────────
def esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def make_img_html(img_raw):
    imgs = [u.strip() for u in str(img_raw or "").split("|") if u.strip() and u.startswith("http")]
    if not imgs:
        return '<span class="no-img">无图</span>'
    return "".join(
        f'<a href="{u}" target="_blank" class="thumb-wrap">'
        f'<img class="thumb" src="{u}" onerror="this.parentElement.style.display=\'none\'"/>'
        f'</a>'
        for u in imgs
    )


def score_stars(score):
    try:
        return "⭐" * int(float(score))
    except:
        return ""


def make_row_html(row):
    img_html = make_img_html(row.get("img", ""))
    desc = esc(row.get("desc", ""))
    tags_raw = row.get("tags", "")
    tag_pills = "".join(
        f'<span class="tag-pill">{esc(t.strip())}</span>'
        for t in tags_raw.split("|") if t.strip()
    )
    survey = row.get("survey", "")
    survey_cls = SURVEY_COLORS.get(survey, "survey-other")
    eval_type = row.get("eval_type", "")
    eval_cls = EVAL_COLORS.get(eval_type, "eval-mid")
    stars = score_stars(row.get("score", ""))
    time_str = esc(row.get("time", ""))
    price_parts = []
    if row.get("est_price"):
        price_parts.append(f'预估 ¥{row["est_price"]}')
    if row.get("check_price"):
        price_parts.append(f'质检 ¥{row["check_price"]}')
    price_html = f'<div class="price-info">{" → ".join(price_parts)}</div>' if price_parts else ""

    return (
        f'<tr class="detail-row">\n'
        f'  <td class="td-img"><div class="img-grid">{img_html}</div></td>\n'
        f'  <td class="td-desc"><div class="desc-text">{desc}</div>{price_html}</td>\n'
        f'  <td class="td-tags">{tag_pills}</td>\n'
        f'  <td class="td-survey"><span class="survey-badge {survey_cls}">{esc(survey)}</span></td>\n'
        f'  <td class="td-eval"><span class="eval-badge {eval_cls}">{esc(eval_type)}</span>'
        f'<div class="score-num">{stars}</div></td>\n'
        f'  <td class="td-time">{time_str}</td>\n'
        f'</tr>'
    )


def build_panel_data(rows):
    """按标签分组，多标签行重复归入每个分类"""
    panel_rows = defaultdict(list)
    for row in rows:
        tags = [t.strip() for t in row.get("tags", "").split("|") if t.strip()]
        classes = set(TAG_CLASS[t] for t in tags if t in TAG_CLASS)
        if not classes:
            classes = {"MISC"}
        for cls in classes:
            panel_rows[cls].append(row)
    return dict(panel_rows)


def make_panel_html(uid, cls, rows):
    count = len(rows)
    label = CLS_META.get(cls, {}).get("label", cls)
    all_tags = sorted({t.strip() for r in rows for t in r.get("tags","").split("|") if t.strip()})
    surveys = sorted({r.get("survey","") for r in rows if r.get("survey","")})
    evals = [e for e in ["非常不满意", "一般", "满意"] if any(r.get("eval_type","") == e for r in rows)]

    tag_opts = "".join(f'<option value="{esc(t)}">{esc(t)}</option>' for t in all_tags)
    survey_opts = "".join(f'<option value="{esc(s)}">{esc(s)}</option>' for s in surveys)
    eval_opts = "".join(f'<option value="{esc(e)}">{esc(e)}</option>' for e in evals)
    rows_html = "\n".join(make_row_html(r) for r in rows)

    return f'''<div class="detail-panel" id="{uid}" style="display:none;">
          <div class="detail-header">
            <span>📋 {esc(label)}明细（共 <strong>{count}</strong> 条）</span>
            <div class="detail-filters">
              <select class="filter-sel filter-tag" onchange="filterDetail('{uid}', 'tag', this.value)" title="按标签筛选">
                <option value="">全部标签</option>{tag_opts}
              </select>
              <select class="filter-sel" onchange="filterDetail('{uid}', 'eval', this.value)">
                <option value="">全部评价</option>{eval_opts}
              </select>
              <select class="filter-sel" onchange="filterDetail('{uid}', 'survey', this.value)">
                <option value="">全部环节</option>{survey_opts}
              </select>
            </div>
          </div>
          <div class="table-wrap">
            <table class="detail-table">
              <thead><tr>
                <th style="width:72px">图片</th>
                <th>问题描述</th>
                <th style="width:130px">反馈标签</th>
                <th style="width:75px">环节</th>
                <th style="width:90px">评价</th>
                <th style="width:120px">提交时间</th>
              </tr></thead>
              <tbody id="tbody_{uid}">
{rows_html}
              </tbody>
            </table>
          </div>
        </div>'''


def make_cat_section_html(cat_name, rows, all_panels):
    """生成某品类的完整 section HTML（数据总览+分类汇总+标签统计+满意度+洞察+建议）"""
    total = len(rows)

    # 满意度统计
    eval_cnt = Counter(r.get("eval_type", "") for r in rows)
    bad = eval_cnt.get("非常不满意", 0)
    mid = eval_cnt.get("一般", 0)
    good = eval_cnt.get("满意", 0)
    bad_rate = round(bad / total * 100, 1) if total else 0

    # 环节分布
    survey_cnt = Counter(r.get("survey", "") for r in rows if r.get("survey", ""))
    survey_items = sorted(survey_cnt.items(), key=lambda x: -x[1])
    max_s = survey_items[0][1] if survey_items else 1
    fill_cls_list = ["fill-s0", "fill-s1", "fill-s2", "fill-s3", "fill-s4"]
    survey_bars = ""
    for i, (s, cnt) in enumerate(survey_items[:8]):
        pct = max(1, round(cnt / max_s * 100))
        fc = fill_cls_list[i] if i < len(fill_cls_list) else "fill-gray"
        survey_bars += f'<div class="bar-row"><div class="bar-label">{esc(s)}</div><div class="bar-wrap"><div class="bar-fill {fc}" style="width:{pct}%"></div></div><div class="bar-num">{cnt}条</div></div>\n'

    # 标签统计 top8
    tag_cnt = Counter()
    for r in rows:
        for t in r.get("tags","").split("|"):
            t = t.strip()
            if t:
                tag_cnt[t] += 1
    top_tags = tag_cnt.most_common(8)
    max_t = top_tags[0][1] if top_tags else 1
    tag_bars = ""
    for i, (t, cnt) in enumerate(top_tags):
        pct = max(1, round(cnt / max_t * 100))
        fc = FILL_COLORS[i] if i < len(FILL_COLORS) else "fill-gray"
        tag_bars += f'<div class="bar-row"><div class="bar-label">{esc(t)}</div><div class="bar-wrap"><div class="bar-fill {fc}" style="width:{pct}%"></div></div><div class="bar-num">{cnt}</div></div>\n'

    # 甜甜圈 SVG
    circ = 2 * math.pi * 42
    def arc(val, offset):
        arc_len = round(circ * val / total, 1) if total else 0
        gap = round(circ - arc_len, 1)
        return f'stroke-dasharray="{arc_len} {gap}" stroke-dashoffset="{offset}"'
    bad_arc = arc(bad, 0)
    mid_arc = arc(mid, -round(circ * bad / total, 1) if total else 0)
    good_offset = -round(circ * (bad + mid) / total, 1) if total else 0
    good_arc = arc(good, good_offset)
    donut_svg = f'''<svg class="donut" width="110" height="110" viewBox="0 0 110 110">
  <circle cx="55" cy="55" r="42" fill="none" stroke="#f0f2f5" stroke-width="13"/>
  <circle cx="55" cy="55" r="42" fill="none" stroke="#e74c3c" stroke-width="13" {bad_arc} transform="rotate(-90 55 55)"/>
  <circle cx="55" cy="55" r="42" fill="none" stroke="#f39c12" stroke-width="13" {mid_arc} transform="rotate(-90 55 55)"/>
  <circle cx="55" cy="55" r="42" fill="none" stroke="#27ae60" stroke-width="13" {good_arc} transform="rotate(-90 55 55)"/>
  <text x="55" y="50" text-anchor="middle" font-size="17" font-weight="800" fill="#e74c3c">{bad_rate}%</text>
  <text x="55" y="64" text-anchor="middle" font-size="10" fill="#7f8c8d">差评率</text>
</svg>'''

    bad_pct = f"{bad/total*100:.1f}%" if total else "0%"
    mid_pct = f"{mid/total*100:.1f}%" if total else "0%"
    good_pct = f"{good/total*100:.1f}%" if total else "0%"

    # 问题分类卡片
    panel_data = build_panel_data(rows)
    # 按条数降序排列
    cls_order = sorted(panel_data.keys(), key=lambda c: (-len(panel_data[c]), c))
    cat_cards_html = ""
    for cls in cls_order:
        meta = CLS_META.get(cls, CLS_META["MISC"])
        cls_rows = panel_data[cls]
        uid = f"cls_{cat_name}_{cls}"
        cls_count = len(cls_rows)
        pct = round(cls_count / total * 100) if total else 0

        # sub-items
        sub_tag_cnt = Counter()
        for r in cls_rows:
            for t in r.get("tags","").split("|"):
                t = t.strip()
                if t:
                    sub_tag_cnt[t] += 1

        dot_color_map = {"A": "dot-red", "B": "dot-orange", "C": "dot-blue", "D": "dot-green", "E": "dot-blue", "MISC": "dot-gray"}
        num_color_map = {"A": "color:var(--red)", "B": "color:var(--orange)", "C": "color:#2980b9", "D": "color:var(--green)", "E": "color:#2980b9", "MISC": "color:#888"}
        dot_cls = dot_color_map.get(cls, "dot-gray")
        num_style = num_color_map.get(cls, "")

        sub_items = ""
        for tag, tcnt in sub_tag_cnt.most_common():
            sub_items += (
                f'<div class="sub-item" onclick="openModalByTag(\'{uid}\', \'{esc(tag)}\')" title="点击查看「{esc(tag)}」全部明细">'
                f'<div class="dot {dot_cls}"></div>'
                f'<span class="tag-name">{esc(tag)}</span>'
                f'<span class="tag-num" style="{num_style}">{tcnt}条</span>'
                f'<span class="tag-link-hint">›</span>'
                f'</div>\n'
            )

        cat_cards_html += f'''      <div class="cat-card {meta['card_cls']}">
        <div class="cat-header" onclick="openModal('{uid}')" style="cursor:pointer;user-select:none;">
          <span class="cat-badge {meta['badge_cls']}">{esc(meta['badge'])}</span>
          <span class="cat-title">{esc(meta['label'])}</span>
          <span class="cat-count">{cls_count}条</span>
          <span class="view-detail-btn">查看明细 ›</span>
        </div>
        <div class="cat-desc">{esc(meta['desc'])} 占本品类总反馈的 <strong>{pct}%</strong>。</div>
        <div class="sub-list">{sub_items}</div>
      </div>\n\n'''

    # 洞察 & 建议（根据存在的分类动态生成）
    active_cls = [c for c in ["A","B","C","D","E"] if c in panel_data]
    insights_html = ""
    for cls in active_cls[:4]:
        tmpl = INSIGHT_TEMPLATES.get(cls)
        if tmpl:
            insights_html += f'<div class="insight-card {tmpl["level"]}"><div class="tag">{tmpl["priority"]}</div><h4>{tmpl["title"]}</h4><p>{tmpl["body"]}</p></div>'

    recs_html = ""
    for i, cls in enumerate(active_cls[:4], 1):
        tmpl = REC_TEMPLATES.get(cls)
        if tmpl:
            recs_html += f'<div class="recommend-item"><div class="rec-num">{i}</div><div class="rec-text"><h5>{tmpl["title"]}</h5><p>{tmpl["body"]}</p></div></div>'

    return f'''
  <div class="section">
    <div class="section-title"><span class="icon" style="background:#e74c3c">📊</span>数据总览 · {esc(cat_name)}</div>
    <div class="stats-grid">
      <div class="stat-card red"><div class="num">{bad}</div><div class="label">非常不满意</div></div>
      <div class="stat-card orange"><div class="num">{mid}</div><div class="label">一般（中评）</div></div>
      <div class="stat-card green"><div class="num">{good}</div><div class="label">满意</div></div>
      <div class="stat-card blue"><div class="num">{bad_rate}%</div><div class="label">差评率</div></div>
    </div>
    <div style="margin-top:22px">
      <div style="font-size:14px;font-weight:700;margin-bottom:10px;color:var(--sub)">按问卷环节分布</div>
      <div class="bar-section">{survey_bars}</div>
    </div>
  </div>
  <div class="section">
    <div class="section-title"><span class="icon" style="background:#e67e22">🗂️</span>用户问题分类汇总 <span style="font-size:12px;font-weight:400;color:var(--sub);margin-left:8px">点击分类标题可展开明细</span></div>
    <div class="category-grid">
{cat_cards_html}    </div>
  </div>
  <div class="two-col">
    <div class="section">
      <div class="section-title"><span class="icon" style="background:#2d6a9f">📌</span>反馈标签统计（Top 8）</div>
      <div class="bar-section">{tag_bars}</div>
    </div>
    <div class="section">
      <div class="section-title"><span class="icon" style="background:#8e44ad">🎯</span>用户满意度分布</div>
      <div class="eval-section">
        <div>{donut_svg}</div>
        <div class="eval-legend">
          <div class="legend-item"><div class="legend-dot" style="background:#e74c3c"></div><span style="flex:1">非常不满意</span><span class="legend-val">{bad}</span><span class="legend-pct">&nbsp;({bad_pct})</span></div>
          <div class="legend-item"><div class="legend-dot" style="background:#f39c12"></div><span style="flex:1">一般</span><span class="legend-val">{mid}</span><span class="legend-pct">&nbsp;({mid_pct})</span></div>
          <div class="legend-item"><div class="legend-dot" style="background:#27ae60"></div><span style="flex:1">满意</span><span class="legend-val">{good}</span><span class="legend-pct">&nbsp;({good_pct})</span></div>
        </div>
      </div>
    </div>
  </div>
  <div class="section">
    <div class="section-title"><span class="icon" style="background:#c0392b">🔍</span>核心问题洞察</div>
    <div class="insight-grid">{insights_html}</div>
  </div>
  <div class="section">
    <div class="section-title"><span class="icon" style="background:#27ae60">🚀</span>优化建议</div>
    <div class="recommend-list">{recs_html}</div>
  </div>
  <div style="text-align:center;color:var(--sub);font-size:12px;padding:12px 0 28px;">
    数据范围：{esc(cat_name)}品类反馈共 {total} 条
  </div>'''


# ── HTML 模板（CSS + JS） ─────────────────────────────────────────────────────
HTML_CSS = """<style>
:root{
  --primary:#1a3a5c;--accent:#e05c2e;--accent2:#f5a623;
  --green:#27ae60;--red:#e74c3c;--orange:#e67e22;
  --gray-bg:#f7f8fa;--border:#e2e6ea;--text:#2c3e50;--sub:#7f8c8d;
}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:-apple-system,'PingFang SC','Microsoft YaHei',sans-serif;background:#eef1f6;color:var(--text);}
.header{background:linear-gradient(135deg,#1a3a5c 0%,#2d6a9f 100%);color:#fff;padding:36px 48px 28px;}
.header h1{font-size:24px;font-weight:700;letter-spacing:.5px;}
.header p{margin-top:6px;font-size:13px;opacity:.75;}
.header .meta{display:flex;gap:32px;margin-top:18px;flex-wrap:wrap;}
.meta-item{display:flex;flex-direction:column;}
.meta-item .num{font-size:32px;font-weight:800;}
.meta-item .label{font-size:12px;opacity:.7;margin-top:2px;}
.selector-bar{background:#fff;border-bottom:1px solid var(--border);padding:14px 32px;display:flex;align-items:center;gap:14px;flex-wrap:wrap;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.07);}
.selector-label{font-size:13px;font-weight:700;color:var(--primary);white-space:nowrap;}
.cat-select{height:36px;padding:0 12px;border:1.5px solid var(--border);border-radius:8px;font-size:13px;color:var(--text);background:#fff;cursor:pointer;min-width:180px;outline:none;}
.cat-select:focus{border-color:#2980b9;}
.total-badge{font-size:12px;color:var(--sub);background:var(--gray-bg);padding:4px 10px;border-radius:12px;}
.container{max-width:1200px;margin:0 auto;padding:28px 24px 40px;}
.section{background:#fff;border-radius:12px;padding:28px 32px;margin-bottom:24px;box-shadow:0 2px 8px rgba(0,0,0,.06);}
.section-title{font-size:17px;font-weight:700;color:var(--primary);margin-bottom:20px;padding-bottom:12px;border-bottom:2px solid var(--border);display:flex;align-items:center;gap:10px;}
.icon{width:28px;height:28px;border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:14px;color:#fff;flex-shrink:0;}
.stats-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;}
.stat-card{background:var(--gray-bg);border-radius:10px;padding:20px;text-align:center;border:1px solid var(--border);}
.stat-card .num{font-size:32px;font-weight:800;}
.stat-card .label{font-size:13px;color:var(--sub);margin-top:6px;}
.stat-card.red .num{color:var(--red);}
.stat-card.orange .num{color:var(--orange);}
.stat-card.green .num{color:var(--green);}
.stat-card.blue .num{color:#2d6a9f;}
.category-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px;}
.cat-card{border-radius:10px;padding:22px 24px;border-left:5px solid;transition:box-shadow .2s;}
.cat-card:hover{box-shadow:0 4px 16px rgba(0,0,0,.1);}
.cat-card.A{background:#fff5f5;border-color:var(--red);}
.cat-card.B{background:#fff9f0;border-color:var(--orange);}
.cat-card.C{background:#f0f8ff;border-color:#3498db;}
.cat-card.D{background:#f5fff5;border-color:var(--green);}
.cat-card .cat-header{display:flex;align-items:center;gap:10px;margin-bottom:12px;padding:4px 0;border-radius:6px;transition:background .15s;}
.cat-card .cat-header:hover{background:rgba(0,0,0,.03);}
.cat-badge{font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;color:#fff;flex-shrink:0;}
.cat-badge.high{background:var(--red);}
.cat-badge.mid{background:var(--orange);}
.cat-badge.low{background:#3498db;}
.cat-badge.info{background:var(--green);}
.cat-title{font-size:15px;font-weight:700;}
.cat-count{margin-left:auto;font-size:22px;font-weight:800;}
.cat-card.A .cat-count{color:var(--red);}
.cat-card.B .cat-count{color:var(--orange);}
.cat-card.C .cat-count{color:#3498db;}
.cat-card.D .cat-count{color:var(--green);}
.cat-desc{font-size:13px;color:var(--sub);line-height:1.7;}
.sub-list{margin-top:12px;display:flex;flex-direction:column;gap:6px;}
.sub-item{display:flex;align-items:center;gap:8px;font-size:13px;cursor:pointer;border-radius:6px;padding:3px 5px;margin:-3px -5px;transition:background .15s;}
.sub-item:hover{background:rgba(0,0,0,.05);}
.sub-item .tag-link-hint{font-size:11px;color:#aaa;margin-left:2px;opacity:0;transition:opacity .15s;}
.sub-item:hover .tag-link-hint{opacity:1;}
.sub-item .dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;}
.dot-red{background:var(--red);}
.dot-orange{background:var(--orange);}
.dot-blue{background:#3498db;}
.dot-green{background:var(--green);}
.dot-gray{background:#aaa;}
.sub-item .tag-name{flex:1;}
.sub-item .tag-num{font-weight:700;font-size:14px;}
.detail-panel{margin-top:16px;border-top:1px dashed var(--border);padding-top:14px;}
.detail-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;font-size:13px;font-weight:700;color:var(--primary);flex-wrap:wrap;gap:8px;}
.detail-filters{display:flex;gap:8px;flex-wrap:wrap;}
.filter-sel{height:28px;padding:0 8px;border:1px solid var(--border);border-radius:6px;font-size:12px;outline:none;background:#fff;cursor:pointer;}
.filter-sel:focus{border-color:#2980b9;}
.table-wrap{overflow-x:auto;border-radius:8px;border:1px solid var(--border);}
.detail-table{width:100%;border-collapse:collapse;font-size:13px;}
.detail-table th{background:#f0f4f8;color:var(--primary);padding:9px 12px;text-align:left;font-weight:600;white-space:nowrap;font-size:12px;}
.detail-table td{padding:10px 12px;border-bottom:1px solid #f0f2f5;vertical-align:top;}
.detail-table tr:last-child td{border-bottom:none;}
.detail-table tr:hover td{background:#fafbfc;}
.detail-row.hidden{display:none;}
.td-img{min-width:68px;max-width:220px;}
.td-img .img-grid{display:flex;flex-wrap:wrap;gap:4px;}
.thumb-wrap{display:block;width:60px;height:60px;border-radius:6px;overflow:hidden;border:1px solid var(--border);background:#f5f5f5;flex-shrink:0;}
.thumb{width:100%;height:100%;object-fit:cover;transition:opacity .2s;}
.thumb:hover{opacity:.85;}
.no-img{font-size:11px;color:#bbb;}
.desc-text{font-size:13px;line-height:1.6;color:var(--text);margin-bottom:4px;}
.price-info{font-size:11px;color:var(--orange);font-weight:600;background:#fff9f0;padding:2px 6px;border-radius:4px;}
.tag-pill{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;background:#e8f4fd;color:#2980b9;margin:2px 2px 2px 0;}
.survey-badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;}
.survey-sel{background:#fde8e8;color:#c0392b;}
.survey-check{background:#fef3e2;color:#d35400;}
.survey-order{background:#e8f4fd;color:#1a5276;}
.survey-service{background:#e6f9ee;color:#1e8449;}
.survey-price{background:#f3e8fd;color:#7d3c98;}
.survey-other{background:#f0f0f0;color:#888;}
.eval-badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;}
.eval-bad{background:#fde8e8;color:var(--red);}
.eval-mid{background:#fef3e2;color:var(--orange);}
.eval-good{background:#e6f9ee;color:var(--green);}
.score-num{font-size:11px;color:var(--sub);margin-top:2px;}
.td-time{font-size:12px;color:var(--sub);white-space:nowrap;}
.bar-section{margin-top:4px;}
.bar-row{display:flex;align-items:center;gap:12px;margin-bottom:10px;}
.bar-label{width:120px;font-size:13px;text-align:right;color:var(--text);flex-shrink:0;}
.bar-wrap{flex:1;background:var(--gray-bg);border-radius:20px;height:22px;overflow:hidden;}
.bar-fill{height:100%;border-radius:20px;display:flex;align-items:center;padding-left:10px;font-size:12px;color:#fff;font-weight:600;transition:width .5s;min-width:4px;}
.bar-num{width:36px;font-size:13px;font-weight:700;text-align:right;}
.fill-red{background:linear-gradient(90deg,#e74c3c,#ff6b5b);}
.fill-orange{background:linear-gradient(90deg,#e67e22,#f39c12);}
.fill-blue{background:linear-gradient(90deg,#2980b9,#3498db);}
.fill-teal{background:linear-gradient(90deg,#16a085,#1abc9c);}
.fill-gray{background:linear-gradient(90deg,#7f8c8d,#95a5a6);}
.fill-s0{background:linear-gradient(90deg,#c0392b,#e74c3c);}
.fill-s1{background:linear-gradient(90deg,#d35400,#e67e22);}
.fill-s2{background:linear-gradient(90deg,#1a5276,#2980b9);}
.fill-s3{background:linear-gradient(90deg,#117a65,#16a085);}
.fill-s4{background:linear-gradient(90deg,#6c3483,#8e44ad);}
.insight-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;}
.insight-card{background:var(--gray-bg);border-radius:10px;padding:18px 20px;border:1px solid var(--border);}
.insight-card .tag{font-size:11px;font-weight:700;padding:2px 8px;border-radius:8px;display:inline-block;margin-bottom:8px;}
.insight-card.p1 .tag{background:#fde8e8;color:var(--red);}
.insight-card.p2 .tag{background:#fef3e2;color:var(--orange);}
.insight-card.p3 .tag{background:#e8f4fd;color:#2980b9;}
.insight-card h4{font-size:14px;font-weight:700;margin-bottom:6px;}
.insight-card p{font-size:13px;color:var(--sub);line-height:1.7;}
.recommend-list{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:4px;}
.recommend-item{display:flex;gap:12px;align-items:flex-start;background:var(--gray-bg);border-radius:8px;padding:14px 16px;}
.rec-num{width:24px;height:24px;border-radius:50%;background:var(--primary);color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:1px;}
.rec-text h5{font-size:13px;font-weight:700;margin-bottom:4px;}
.rec-text p{font-size:12px;color:var(--sub);line-height:1.6;}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:20px;}
svg.donut{overflow:visible;}
.eval-section{display:flex;gap:28px;align-items:center;flex-wrap:wrap;}
.eval-legend{display:flex;flex-direction:column;gap:8px;}
.legend-item{display:flex;align-items:center;gap:8px;font-size:13px;}
.legend-dot{width:12px;height:12px;border-radius:3px;flex-shrink:0;}
.legend-val{font-weight:700;}
.legend-pct{color:var(--sub);font-size:12px;}
@media(max-width:800px){
  .category-grid,.insight-grid,.recommend-list,.two-col{grid-template-columns:1fr;}
  .stats-grid{grid-template-columns:repeat(2,1fr);}
  .header{padding:24px 20px;}
  .container{padding:16px 12px;}
  .selector-bar{padding:12px 16px;}
}
.modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:1000;display:none;align-items:center;justify-content:center;padding:20px;}
.modal-box{background:#fff;border-radius:16px;width:min(96vw,1100px);max-height:88vh;display:flex;flex-direction:column;box-shadow:0 24px 64px rgba(0,0,0,.3);overflow:hidden;}
.modal-head{display:flex;align-items:center;gap:12px;padding:18px 24px 14px;border-bottom:1px solid var(--border);background:#fff;flex-shrink:0;}
.modal-head h3{font-size:16px;font-weight:700;color:var(--primary);flex:1;margin:0;}
.modal-close{background:none;border:none;font-size:20px;cursor:pointer;color:var(--sub);padding:4px 8px;border-radius:6px;line-height:1;}
.modal-close:hover{background:var(--gray-bg);color:var(--text);}
.modal-body{overflow-y:auto;flex:1;}
#modalPanels .detail-panel{display:none;padding:0;}
#modalPanels .detail-header{display:flex;align-items:center;justify-content:space-between;padding:14px 24px 10px;border-bottom:1px solid var(--border);gap:12px;flex-wrap:wrap;}
#modalPanels .detail-header > span{font-size:14px;font-weight:600;color:var(--primary);}
#modalPanels .detail-filters{display:flex;gap:8px;flex-wrap:wrap;}
#modalPanels .table-wrap{overflow-x:auto;}
.view-detail-btn{margin-left:auto;font-size:12px;color:var(--accent);font-weight:600;padding:3px 10px;border:1.5px solid var(--accent);border-radius:20px;white-space:nowrap;}
.cat-header:hover .view-detail-btn{background:var(--accent);color:#fff;}
</style>"""

HTML_JS = """<script>
let _currentUid = null;

function _updatePanelCount(uid) {
  const tbody = document.getElementById('tbody_' + uid);
  const panel = document.getElementById(uid);
  if (!tbody || !panel) return;
  const total = tbody.querySelectorAll('tr.detail-row').length;
  const visible = tbody.querySelectorAll('tr.detail-row:not(.hidden)').length;
  const suffix = visible < total ? '，已筛选' : '';
  const headerSpan = panel.querySelector('.detail-header > span');
  if (!headerSpan) return;
  headerSpan.innerHTML = headerSpan.innerHTML.replace(
    /（共 <strong>\\d+<\\/strong> 条[^）]*）/,
    '（共 <strong>' + visible + '</strong> 条' + suffix + '）'
  );
  const modalTitle = document.getElementById('modalTitle');
  if (modalTitle) {
    if (/（共 \\d+ 条[^）]*）/.test(modalTitle.textContent)) {
      modalTitle.textContent = modalTitle.textContent.replace(/（共 \\d+ 条[^）]*）/, '（共 ' + visible + ' 条' + suffix + '）');
    } else {
      modalTitle.textContent = modalTitle.textContent.trim() + '（共 ' + visible + ' 条' + suffix + '）';
    }
  }
}

function openModal(uid) {
  document.querySelectorAll('#modalPanels .detail-panel').forEach(function(el) {
    el.style.display = 'none';
  });
  const panel = document.getElementById(uid);
  if (panel) {
    panel.style.display = 'block';
    _currentUid = uid;
    const tbody = document.getElementById('tbody_' + uid);
    if (tbody) {
      tbody.dataset.tagFilter = '';
      tbody.querySelectorAll('tr.detail-row').forEach(function(tr) { tr.classList.remove('hidden'); });
    }
    panel.querySelectorAll('.filter-sel').forEach(function(sel) { sel.value = ''; });
    const headerSpan = panel.querySelector('.detail-header > span');
    if (headerSpan && tbody) {
      const total = tbody.querySelectorAll('tr.detail-row').length;
      headerSpan.innerHTML = headerSpan.innerHTML.replace(/（共 <strong>\\d+<\\/strong> 条[^）]*）/, '（共 <strong>' + total + '</strong> 条）');
    }
    const titleEl = panel.querySelector('.detail-header > span');
    document.getElementById('modalTitle').textContent = titleEl ? titleEl.textContent.replace(/\\s*✕.*$/, '') : '反馈明细';
  }
  document.getElementById('modalOverlay').style.display = 'flex';
  document.body.style.overflow = 'hidden';
}

function closeModal() {
  document.getElementById('modalOverlay').style.display = 'none';
  document.body.style.overflow = '';
  _currentUid = null;
}

function handleOverlayClick(e) {
  if (e.target === document.getElementById('modalOverlay')) closeModal();
}

function filterDetail(uid, field, val) {
  const tbody = document.getElementById('tbody_' + uid);
  if (!tbody) return;
  const rows = tbody.querySelectorAll('tr.detail-row');
  rows.forEach(function(tr) {
    const evalText   = tr.querySelector('.eval-badge')   ? tr.querySelector('.eval-badge').textContent   : '';
    const surveyText = tr.querySelector('.survey-badge') ? tr.querySelector('.survey-badge').textContent : '';
    let show = true;
    if (field === 'eval'   && val && evalText   !== val) show = false;
    if (field === 'survey' && val && surveyText !== val) show = false;
    if (show) {
      const otherField = field === 'eval' ? 'survey' : 'eval';
      const otherSel = tr.closest('.detail-panel').querySelector('.filter-sel[onchange*="\\'' + otherField + '\\'"]');
      if (otherSel && otherSel.value) {
        const otherText = otherField === 'eval' ? evalText : surveyText;
        if (otherText !== otherSel.value) show = false;
      }
    }
    const tagFilter = tbody.dataset.tagFilter || '';
    const tagSelEl = tbody.closest('.detail-panel') ? tbody.closest('.detail-panel').querySelector('.filter-tag') : null;
    const tagSelVal = tagSelEl ? tagSelEl.value : '';
    const activeTag = tagFilter || tagSelVal;
    if (activeTag) {
      const pills = tr.querySelectorAll('.tag-pill');
      let hasTag = false;
      pills.forEach(function(p) { if (p.textContent.trim() === activeTag) hasTag = true; });
      if (!hasTag) show = false;
    }
    tr.classList.toggle('hidden', !show);
  });
  _updatePanelCount(uid);
}

function openModalByTag(uid, tagName) {
  openModal(uid);
  const tbody = document.getElementById('tbody_' + uid);
  if (!tbody) return;
  tbody.dataset.tagFilter = tagName;
  const panel = document.getElementById(uid);
  if (panel) {
    panel.querySelectorAll('.filter-sel').forEach(function(sel) { sel.value = ''; });
    const tagSel = panel.querySelector('.filter-tag');
    if (tagSel) { tagSel.value = tagName; }
  }
  const rows = tbody.querySelectorAll('tr.detail-row');
  rows.forEach(function(tr) {
    const pills = tr.querySelectorAll('.tag-pill');
    let hasTag = false;
    pills.forEach(function(p) { if (p.textContent.trim() === tagName) hasTag = true; });
    tr.classList.toggle('hidden', !hasTag);
  });
  const visibleCount = tbody.querySelectorAll('tr.detail-row:not(.hidden)').length;
  const modalTitle = document.getElementById('modalTitle');
  if (modalTitle) modalTitle.textContent = '「' + tagName + '」反馈明细（共 ' + visibleCount + ' 条）';
  if (panel) {
    const headerSpan = panel.querySelector('.detail-header > span');
    if (headerSpan) {
      headerSpan.innerHTML = '📋 「<strong>' + tagName + '」</strong>明细（共 <strong>' + visibleCount + '</strong> 条）'
        + ' <span style="font-size:11px;color:#888;font-weight:400;margin-left:8px;cursor:pointer;" onclick="clearTagFilter(\\'' + uid + '\\')" title="清除标签过滤">✕ 查看全部</span>';
    }
  }
}

function clearTagFilter(uid) {
  const tbody = document.getElementById('tbody_' + uid);
  if (!tbody) return;
  tbody.dataset.tagFilter = '';
  const rows = tbody.querySelectorAll('tr.detail-row');
  rows.forEach(function(tr) { tr.classList.remove('hidden'); });
  const panelEl = document.getElementById(uid);
  if (panelEl) {
    const tagSel = panelEl.querySelector('.filter-tag');
    if (tagSel) tagSel.value = '';
  }
  const panel = document.getElementById(uid);
  if (panel) {
    const headerSpan = panel.querySelector('.detail-header > span');
    if (headerSpan) {
      const total = rows.length;
      headerSpan.innerHTML = headerSpan.innerHTML
        .replace(/📋 「.*?」<\\/strong>明细/, '📋 反馈明细')
        .replace(/（共 <strong>\\d+<\\/strong> 条[^）]*）/, '（共 <strong>' + total + '</strong> 条）')
        .replace(/<span[^>]*>✕ 查看全部<\\/span>/, '');
    }
    panel.querySelectorAll('.filter-sel').forEach(function(sel) { sel.value = ''; });
  }
  const modalTitle = document.getElementById('modalTitle');
  if (modalTitle && panel) {
    const titleEl = panel.querySelector('.detail-header span');
    if (titleEl) modalTitle.textContent = titleEl.textContent.replace(/^📋\\s*/, '').replace(/（共.*?条）.*/, '').trim() || '反馈明细';
  }
}

document.addEventListener('keydown', function(e) { if (e.key === 'Escape') closeModal(); });
switchCat(CAT_LIST[0]);
</script>"""


# ── 主函数 ────────────────────────────────────────────────────────────────────
def generate(excel_path: str, output_path: str = None):
    print(f"📖 读取 Excel：{excel_path}")
    rows = load_excel(excel_path)
    print(f"   共 {len(rows)} 行数据")

    # 按品类分组
    cat_data = defaultdict(list)
    for row in rows:
        cat = row.get("cat", "未知品类").strip() or "未知品类"
        cat_data[cat].append(row)

    cats = list(cat_data.keys())
    total_rows = len(rows)
    print(f"   品类：{cats}")

    # 生成每个品类的 panel + cat HTML
    cat_html_dict = {}
    all_panels_html = []
    for cat_name, cat_rows in cat_data.items():
        panel_data = build_panel_data(cat_rows)
        for cls, cls_rows in panel_data.items():
            uid = f"cls_{cat_name}_{cls}"
            all_panels_html.append(make_panel_html(uid, cls, cls_rows))
        cat_html_dict[cat_name] = make_cat_section_html(cat_name, cat_rows, panel_data)

    # 品类下拉选项
    cat_options = "\n".join(
        f'    <option value="{esc(c)}">{esc(c)}（{len(cat_data[c])}条）</option>'
        for c in cats
    )
    cat_list_js = json.dumps(cats, ensure_ascii=False)
    cat_html_js = json.dumps(cat_html_dict, ensure_ascii=False)
    # 防止 </script> 注入
    cat_html_js = cat_html_js.replace("</script>", "<\\/script>")

    panels_html = "\n".join(all_panels_html)

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>用户反馈分析报告</title>
{HTML_CSS}
</head>
<body>

<div class="header">
  <h1>📊 用户反馈分析报告</h1>
  <p>数据来源：{esc(excel_path.split("/")[-1])} · 选择品类查看深度分析 · 点击问题分类展开每条反馈明细</p>
  <div class="meta">
    <div class="meta-item"><span class="num">{total_rows:,}</span><span class="label">反馈总条数</span></div>
    <div class="meta-item"><span class="num">{len(cats)}</span><span class="label">已分析品类数</span></div>
    <div class="meta-item"><span class="num">深度</span><span class="label">可展开逐条明细</span></div>
  </div>
</div>

<div class="selector-bar">
  <span class="selector-label">选择品类：</span>
  <select class="cat-select" id="catSelect" onchange="switchCat(this.value)">
{cat_options}
  </select>
  <span class="total-badge" id="totalBadge"></span>
</div>

<div class="container" id="mainContent"></div>

<script>
const CAT_HTML = {cat_html_js}
const CAT_LIST = {cat_list_js}

function switchCat(name) {{
  document.getElementById('catSelect').value = name;
  document.getElementById('mainContent').innerHTML = CAT_HTML[name] || '<p>暂无数据</p>';
  const badge = document.getElementById('totalBadge');
  // 从 CAT_HTML 里取总条数
  const m = CAT_HTML[name] ? CAT_HTML[name].match(/共 (\\d+) 条/) : null;
  badge.textContent = m ? name + ' · ' + m[1] + ' 条' : name;
}}
</script>

<!-- 模态浮框 -->
<div class="modal-overlay" id="modalOverlay" onclick="handleOverlayClick(event)">
  <div class="modal-box">
    <div class="modal-head">
      <h3 id="modalTitle">反馈明细</h3>
      <button class="modal-close" onclick="closeModal()" title="关闭（ESC）">✕</button>
    </div>
    <div class="modal-body">
      <div id="modalPanels">
{panels_html}
      </div>
    </div>
  </div>
</div>

{HTML_JS}
</body>
</html>"""

    if not output_path:
        excel_stem = Path(excel_path).stem
        output_path = str(Path(excel_path).parent / f"{excel_stem}_分析报告.html")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅ 报告已生成：{output_path}")
    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python3 gen_report.py <excel_path> [output_path]")
        sys.exit(1)
    excel_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    generate(excel_path, output_path)
