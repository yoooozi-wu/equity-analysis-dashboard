# -*- coding: utf-8 -*-
"""
build_dashboard.py — 股权/投资分析看板构建模板（equity-analysis-dashboard skill）
=====================================================================
用法：
    python build_dashboard.py
前置：
    1) 把 assets/echarts.min.js 复制到本脚本同目录（输出目录）。
    2) 修改下方 CONFIG / EXCEL 指向目标公司的股权梳理 Excel。
    3) 按 TODO 标注调整各 sheet 的提取逻辑（不同公司表头/列序不同）。

产物：<公司>投资分析看板.html（自包含、离线、内嵌一份 echarts）。
依赖：pandas / numpy / openpyxl（已装在隔离 python 环境）。

说明：本模板以"港股 18A 生物医药·少数股东/PE 视角"为范例，含 6 张核心图 +
少数股东分析卡 + 方案逻辑拆解（逻辑链图+工具对比+关键判断）。按公司数据改
提取段即可，图表 option 结构可直接复用。详细规则见 references/workflow.md
与 references/logic-chain.md。
"""
import json
import re
from pathlib import Path

import pandas as pd
import numpy as np

# ============================== CONFIG（按公司改） ==============================
OUT = Path(__file__).resolve().parent
EXCEL = r"C:\path\to\equity_data.xlsx"  # TODO: 改成本地股权梳理 Excel 路径
GEN_DATE = "2026-09-10"
COMPANY = "示例公司"          # TODO: 目标公司名
TICKER = "XXXX.HK"            # TODO: 证券代码
CURRENT = 3.60          # 现价（港元）；TODO: 取真实值
UNLOCK = "2026-08-14"   # 解禁日（禁售承诺最后一日）；TODO: 按业务确认
TOTAL_SHARES = 4.57e8  # 总股本（用于市值）；TODO
# ================================================================================

TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>__COMPANY__ 投资分析看板</title>
<style>
* { box-sizing: border-box; }
:root { --bg:#f4f6f8; --card:#fff; --text:#1f2933; --muted:#6b7280; --line:#e5e7eb;
  --good:#0f9d58; --warn:#b45309; --bad:#c62828; --accent:#1a6fc4; }
body { margin:0; background:var(--bg); color:var(--text);
  font-family:'Microsoft YaHei','PingFang SC','Hiragino Sans GB','Noto Sans CJK SC',sans-serif;
  font-size:14px; line-height:1.6; }
.wrap { max-width:1180px; margin:0 auto; padding:0 20px 60px; }
header.masthead { padding:34px 0 20px; }
.masthead h1 { margin:0 0 6px; font-size:26px; letter-spacing:.5px; }
.masthead p { margin:0; color:var(--muted); font-size:13px; }
nav { position:sticky; top:0; z-index:20; background:rgba(255,255,255,.94); backdrop-filter:blur(8px);
  border-bottom:1px solid var(--line); margin:0 -20px 24px; padding:10px 20px; display:flex; flex-wrap:wrap; gap:6px; }
.nav-item { text-decoration:none; color:var(--muted); font-size:12.5px; padding:5px 11px; border-radius:999px;
  border:1px solid var(--line); transition:all .15s; white-space:nowrap; }
.nav-item:hover { color:var(--accent); border-color:var(--accent); background:#f0f7ff; }
.nav-num { font-weight:700; color:var(--accent); margin-right:5px; }
.kpi-grid { display:grid; gap:14px; margin-bottom:30px; grid-template-columns:repeat(auto-fit,minmax(178px,1fr)); }
.kpi { background:var(--card); border:1px solid var(--line); border-radius:10px; padding:14px 16px; }
.kpi-label { font-size:12px; color:var(--muted); }
.kpi-value { font-size:23px; font-weight:700; margin:2px 0 3px; letter-spacing:-.3px; }
.kpi-unit { font-size:12px; font-weight:400; color:var(--muted); margin-left:3px; }
.kpi-delta { font-size:12px; font-weight:600; }
.kpi-delta.good { color:var(--good); } .kpi-delta.warn { color:var(--warn); } .kpi-delta.bad { color:var(--bad); }
.kpi-note { font-size:11.5px; color:var(--muted); margin-top:2px; }
.card { background:var(--card); border:1px solid var(--line); border-radius:12px; padding:20px 22px 18px; margin-bottom:22px; }
.card-head { display:flex; align-items:flex-start; gap:12px; margin-bottom:14px; }
.card-num { flex:none; width:30px; height:30px; border-radius:8px; background:var(--accent); color:#fff;
  font-size:13px; font-weight:700; display:flex; align-items:center; justify-content:center; }
.card-title { margin:0; font-size:17px; font-weight:700; line-height:1.4; }
.card-sub { font-size:12px; color:var(--muted); margin-top:2px; }
.btn { margin-left:auto; flex:none; font-size:12px; padding:5px 13px; cursor:pointer; background:#fff;
  color:var(--muted); border:1px solid var(--line); border-radius:6px; }
.btn:hover { color:var(--accent); border-color:var(--accent); }
.chart { width:100%; height:440px; }
.note { margin-top:14px; padding:12px 15px; background:#f7f9fb; border-left:3px solid var(--accent);
  border-radius:0 6px 6px 0; font-size:13px; color:#374151; line-height:1.75; }
.note-tag { display:inline-block; font-size:11px; font-weight:700; color:var(--accent); background:#e8f1fb;
  border-radius:4px; padding:1px 7px; margin-right:8px; vertical-align:1px; }
.grid2 { display:grid; grid-template-columns:1fr 1fr; gap:14px; }
.acard { border:1px solid var(--line); border-radius:10px; padding:14px 16px; background:#fcfdff; }
.acard h3 { margin:0 0 8px; font-size:15px; color:var(--accent); }
.acard p { margin:0; font-size:13px; color:#374151; line-height:1.7; }
.acard.warn { border-color:#f3c0c0; background:#fdf5f5; } .acard.warn h3 { color:var(--bad); }
.muted { color:var(--muted); font-size:12px; }
footer { color:var(--muted); font-size:12px; text-align:center; padding-top:10px; }
.cmp-tbl { width:100%; border-collapse:collapse; margin-top:12px; font-size:12px; }
.cmp-tbl th, .cmp-tbl td { border:1px solid var(--line); padding:6px 8px; text-align:center; }
.cmp-tbl th { background:#eef3f8; color:#334155; font-weight:600; }
.cmp-tbl td:first-child { text-align:left; font-weight:600; color:#334155; }
.cmp-tbl tr.me td { background:#fdf0f0; color:var(--bad); font-weight:700; }
@media (max-width:720px){ .wrap{padding:0 12px 40px;} nav{margin:0 -12px 20px; padding:9px 12px;}
  .chart{height:360px;} .card{padding:16px;} .btn{display:none;} .grid2{grid-template-columns:1fr;} }
</style>
</head>
<body>
<div class="wrap">
  <header class="masthead">
    <h1>__COMPANY__ (__TICKER__) 投资分析看板 · 少数股东视角</h1>
    <p>数据源：__EXCEL_NAME__ · 框架：股票财务分析 + AI尽调·财务分析专家 + 方案逻辑拆解 · 制图日期 __GEN_DATE__</p>
  </header>
  <nav>__NAV__</nav>
  <div class="kpi-grid">__KPI__</div>
__SECTIONS__
  <footer>由 Smart Charts / ECharts 5.4.4 生成 · 仅供分析参考，非投资建议</footer>
</div>
<script>__ECHARTS__</script>
<script>
var charts = {};
function register(id, domId, option) {
  delete option.title; option.backgroundColor = 'transparent';
  var dom = document.getElementById(domId);
  var chart = echarts.init(dom); chart.setOption(option); charts[id] = chart;
  new ResizeObserver(function () { chart.resize(); }).observe(dom);
  window.addEventListener('resize', function () { chart.resize(); });
}
__INIT__
Array.prototype.forEach.call(document.querySelectorAll('.btn[data-save]'), function (btn) {
  btn.addEventListener('click', function () {
    var chart = charts[btn.getAttribute('data-save')]; if (!chart) return;
    var url = chart.getDataURL({ type:'png', pixelRatio:2, backgroundColor:'#ffffff' });
    var a = document.createElement('a'); a.href = url;
    a.download = btn.closest('.card').querySelector('.card-title').textContent.trim() + '.png';
    a.click();
  });
});
</script>
</body>
</html>
"""

# ---------------- 1. 上市后股价序列（TODO: 按实际 sheet/列调整） ----------------
d = pd.read_excel(EXCEL, sheet_name="上市后交易情况", header=None)  # TODO: 改 sheet 名
price_dates, price_close, price_vol = [], [], []
for i in range(1, len(d)):
    dt = d.iloc[i, 0]
    if pd.notna(dt) and isinstance(dt, str) and dt[0].isdigit():
        price_dates.append(str(dt)[:10]); price_close.append(round(float(d.iloc[i, 3]), 3))
        price_vol.append(round(float(d.iloc[i, 5]) / 1e4, 1))   # 万股
first_close, last_close = price_close[0], price_close[-1]
drop_pct = (last_close - first_close) / first_close * 100
avg_vol = np.mean(price_vol)
mkt_cap = round(last_close * TOTAL_SHARES / 1e8, 1)

# ---------------- 1.5 解禁窗口 & 可比公司（TODO: 改 sheet 名/列） ----------------
ui = price_dates.index(UNLOCK)
pre_idx = list(range(max(0, ui - 8), ui)); post_idx = list(range(ui + 1, min(len(price_dates), ui + 9)))
pre_dates = [price_dates[i] for i in pre_idx]; pre_close = [price_close[i] for i in pre_idx]
post_dates = [price_dates[i] for i in post_idx]; post_close = [price_close[i] for i in post_idx]
unlock_close = price_close[ui]
win_dates = pre_dates + [UNLOCK] + post_dates
win_close = pre_close + [unlock_close] + post_close
win_chg = [round((win_close[k] / win_close[k - 1] - 1) * 100, 2) for k in range(1, len(win_close))]
pre_cum = round((pre_close[-1] / pre_close[0] - 1) * 100, 2)
chg_unlock = round((unlock_close / pre_close[-1] - 1) * 100, 2)
post_cum = round((post_close[-1] / unlock_close - 1) * 100, 2) if post_close else None

peer_sheets = {"可比公司A":"可比公司1", "可比公司B":"可比公司2",
               "可比公司C":"可比公司3", "可比公司D":"可比公司4"}  # TODO: 改可比公司名与 sheet 名
cmp_rows = []
for nm, sh in peer_sheets.items():
    dd = pd.read_excel(EXCEL, sheet_name=sh, header=None)
    s = [(str(dd.iloc[i, 0])[:10], float(dd.iloc[i, 1])) for i in range(1, len(dd))
         if pd.notna(dd.iloc[i, 0]) and isinstance(dd.iloc[i, 0], str) and dd.iloc[i, 0][0].isdigit()]
    if not s: continue
    cmp_rows.append({"name": nm, "list_d": s[0][0], "unlock_d": "?", "first": s[0][1], "last": s[-1][1],
                     "chg": round((s[-1][1] / s[0][1] - 1) * 100, 1)})
self_chg = round((last_close / first_close - 1) * 100, 1)
cmp_rows = [{"name": COMPANY, "list_d": price_dates[0], "unlock_d": UNLOCK,
             "first": first_close, "last": last_close, "chg": self_chg}] + cmp_rows

# ---------------- 2. 融资估值轨迹（TODO: 轮次/成本按文档录入） ----------------
round_labels = ["成立", "老股-1", "Pre-A", "A轮", "老股-2", "B轮", "B+轮", "老股-3"]  # TODO: 改真实轮次
round_cost = [1.00, 2.00, 2.761, 8.287, 9.459, 9.886, 11.065, 6.920]  # TODO: 改真实成本
peak_idx = round_cost.index(max(round_cost))

# ---------------- 3. 股东结构 & 成本（TODO: 按股东表测算） ----------------
holder_struct = [("控股股东及一致行动人", 33.2, "#1a6fc4"), ("机构股东A", 9.62, "#E69F00"),
                ("机构股东B", 5.81, "#009E73"), ("机构股东C", 5.55, "#D55E00"),
                ("其他机构及公众股东", 45.82, "#999999")]  # TODO: 按股东表测算
holder_cost = [("控股股东(加权)", 1.00), ("机构股东B", 2.64), ("机构股东C", 5.33), ("老股", 6.92),
               ("机构股东A(后期)", 8.29), ("B+轮", 11.07)]  # TODO: 按股东表测算

KPI = [
    ("最新收盘价", "%.2f" % CURRENT, "港元", "较首日 %.1f%%" % drop_pct, "bad", "首日 %.2f" % first_close),
    ("当前市值", "≈%.0f" % mkt_cap, "亿港元", "较 IPO 大幅缩水", "bad", "总股本约 %.2f 亿股" % (TOTAL_SHARES/1e8)),
    ("日均成交量", "%.0f" % avg_vol, "万股", "流动性极弱", "warn", "占流通盘比例很低"),
    ("解禁日", UNLOCK, "", "已解禁可减持", "warn", "禁售承诺最后一日"),
    ("控股股东持股", "33.2", "%", "控制权集中", "good", "控股股东+ESOP+一致行动"),
    ("累计融资", "16.77", "亿元", "IPO 前多轮", "warn", "每股成本 1→11 元"),
]

# ---------------- 图表 option 构建 ----------------
def build_option_price():
    return {"tooltip":{"trigger":"axis"},
        "legend":{"data":["收盘价","成交量"],"top":"bottom","type":"scroll"},
        "grid":{"left":"3%","right":"3%","bottom":"14%","top":"8%","containLabel":True},
        "xAxis":{"type":"category","data":price_dates,"axisLabel":{"interval":"auto"}},
        "yAxis":[{"type":"value","name":"收盘价(港元)","scale":True},
                 {"type":"value","name":"成交量(万股)","splitLine":{"show":False}}],
        "series":[{"name":"收盘价","type":"line","smooth":True,"showSymbol":False,"data":price_close,
            "markLine":{"symbol":"none","lineStyle":{"color":"#c62828","type":"dashed"},
                        "label":{"formatter":"解禁日 "+UNLOCK,"color":"#c62828"},"data":[{"xAxis":UNLOCK}]}},
            {"name":"成交量","type":"bar","yAxisIndex":1,"data":price_vol,
             "itemStyle":{"color":"rgba(26,111,196,0.25)"}}]}

def build_option_round():
    return {"tooltip":{"trigger":"axis"},
        "grid":{"left":"3%","right":"6%","bottom":"12%","top":"12%","containLabel":True},
        "xAxis":{"type":"category","data":round_labels,"axisLabel":{"interval":0,"rotate":20}},
        "yAxis":{"type":"value","name":"每股成本(元)"},
        "series":[{"name":"每股成本","type":"line","symbolSize":9,"data":round_cost,
            "markPoint":{"data":[{"coord":[peak_idx,round_cost[peak_idx]],"value":"B+轮峰值","itemStyle":{"color":"#c62828"}},
                                {"coord":[len(round_cost)-1,round_cost[-1]],"value":"老股折价37%","itemStyle":{"color":"#b45309"}}]}}]}

def build_option_struct():
    names=[h[0] for h in holder_struct][::-1]; vals=[h[1] for h in holder_struct][::-1]; cols=[h[2] for h in holder_struct][::-1]
    return {"tooltip":{"trigger":"axis","axisPointer":{"type":"shadow"}},
        "grid":{"left":"3%","right":"8%","bottom":"6%","top":"6%","containLabel":True},
        "xAxis":{"type":"value","name":"持股占比(%)"},"yAxis":{"type":"category","data":names},
        "series":[{"type":"bar","data":[{"value":v,"itemStyle":{"color":c}} for v,c in zip(vals,cols)]}]}

def build_option_cost():
    names=[h[0] for h in holder_cost][::-1]; vals=[h[1] for h in holder_cost][::-1]
    return {"tooltip":{"trigger":"axis","axisPointer":{"type":"shadow"}},
        "grid":{"left":"3%","right":"8%","bottom":"6%","top":"6%","containLabel":True},
        "xAxis":{"type":"value","name":"持股成本(元/股)"},"yAxis":{"type":"category","data":names},
        "series":[{"type":"bar","data":vals,"markLine":{"symbol":"none","lineStyle":{"color":"#c62828","width":2},
            "label":{"formatter":"现价 %.2f"%CURRENT,"color":"#c62828"},"data":[{"xAxis":CURRENT}]}}]}

def build_option_unlock():
    return {"tooltip":{"trigger":"axis","axisPointer":{"type":"cross"}},
        "legend":{"data":["收盘价","当日涨跌幅"],"bottom":"1%","type":"scroll"},
        "grid":{"left":"3%","right":"4%","bottom":"16%","top":"8%","containLabel":True},
        "xAxis":{"type":"category","data":win_dates,"axisLabel":{"interval":0,"rotate":42,"fontSize":10}},
        "yAxis":[{"type":"value","name":"收盘价(港元)","scale":True},
                 {"type":"value","name":"涨跌幅(%)","axisLabel":{"formatter":"{value}%"},"splitLine":{"show":False}}],
        "series":[{"name":"收盘价","type":"line","smooth":True,"symbolSize":6,"data":win_close,
            "markLine":{"symbol":"none","lineStyle":{"color":"#c62828","type":"dashed"},
                        "label":{"formatter":"解禁日 "+UNLOCK,"color":"#c62828","position":"insideEndTop"},"data":[{"xAxis":UNLOCK}]}},
            {"name":"当日涨跌幅","type":"bar","yAxisIndex":1,"data":[None]+win_chg,"itemStyle":{"color":"#b45309"},
             "markArea":{"silent":True,"itemStyle":{"color":"rgba(26,111,196,.07)"},
                        "data":[[{"xAxis":pre_dates[0]},{"xAxis":pre_dates[-1]}]]}}]}

def build_option_cmp():
    names=[r["name"] for r in cmp_rows]; chgs=[r["chg"] for r in cmp_rows]
    cols=["#c62828" if n==COMPANY else "#1a6fc4" for n in names]
    return {"tooltip":{"trigger":"axis","axisPointer":{"type":"shadow"},
            "formatter":"function(p){return p[0].name+'<br/>较首发: '+p[0].value+'%';}"},
        "grid":{"left":"3%","right":"10%","bottom":"6%","top":"8%","containLabel":True},
        "xAxis":{"type":"value","name":"较首发收盘价涨跌幅(%)","axisLabel":{"formatter":"{value}%"}},
        "yAxis":{"type":"category","data":names[::-1]},
        "series":[{"type":"bar","data":[{"value":v,"itemStyle":{"color":c}} for v,c in zip(chgs,cols)][::-1],
            "label":{"show":True,"position":"right","formatter":"{c}%","fontSize":11}}]}

charts = [
    ("c1","上市后交易情况（股价+成交量）","日期 · 港元/万股",
     "股价从首日 %.2f 港元一路跌至 %.2f 港元，区间跌幅约 %.0f%%；%s 解禁后创 %.2f 港元新低。日均成交仅 %.0f 万股，流动性极弱。"
     % (first_close,last_close,drop_pct,UNLOCK,min(price_close),avg_vol), build_option_price()),
    ("c2","历史融资估值轨迹（每股成本）","历次融资 · 元/股",
     "每股成本从成立时 1.00 元升至 B+轮峰值；末期老股转让价较 B+轮折价约 37%，是 down-round 信号。后期进场者成本普遍高于现价，已深度套牢。", build_option_round()),
    ("c3","股东持股结构","持股总数占比",
     "控股股东及一致行动人合计约 33.2%，控制权集中；前三大机构股东合计约 21%；其余为公众及其他机构。", build_option_struct()),
    ("c4","主要股东持股成本 vs 现价","元/股 · 现价 %.2f 港元"%CURRENT,
     "控股股东成本仅 1.00 元附近，即便跌至现价仍大幅浮盈；B+轮/机构股东/老股等成本均高于现价，处于套牢。现价线以下减持动力弱，以上面临退出亏损。", build_option_cost()),
    ("c5","解禁日前 8 天改变（比例）","前8日+解禁日+后8日 · 港元/%",
     "解禁前 8 日（%s 至 %s）仅累计 %+.2f%%；**解禁当日（%s）收 %.2f、单日 %+.2f%%**，随后 8 日再累计 %+.2f%%——悬崖在解禁之后。少数股东退出须警惕解禁后集中供给与流动性踩踏。"
     % (pre_dates[0],pre_dates[-1],pre_cum,UNLOCK,unlock_close,chg_unlock,post_cum), build_option_unlock()),
    ("c6","可比公司对比（较首发涨跌幅）","港股 18A 同业 · %",
     "以各公司上市后首发收盘价计，%s累计 %+.1f%% 为同业最差；同业中已商业化者一枝独秀，其余多数下跌 35%%–55%%。" % (COMPANY, self_chg), build_option_cmp()),
]
cmp_table_html = ('<table class="cmp-tbl"><thead><tr><th>公司</th><th>上市日</th><th>解禁日</th><th>首发收</th><th>最新收</th><th>较首发</th></tr></thead><tbody>'
    + "".join('<tr class="%s"><td>%s</td><td>%s</td><td>%s</td><td>%.2f</td><td>%.2f</td><td>%+.1f%%</td></tr>'
             % ("me" if r["name"]==COMPANY else "", r["name"], r["list_d"], r["unlock_d"], r["first"], r["last"], r["chg"]) for r in cmp_rows)
    + "</tbody></table>")

# ---------------- 方案逻辑拆解（见 references/logic-chain.md） ----------------
TOOLTIP_FN = "function(p){ if(p.data && p.data.detail){ return p.data.detail.replace(/\\n/g,'<br/>'); } return p.name; }"
logic_cats = ["解禁前提","股东冲动","抛压推论","行动原则","减持阶段","退出目标"]
logic_cat_color = ["#3b82f6","#f59e0b","#ef4444","#10b981","#14b8a6","#8b5cf6"]
logic_nodes = [
    {"id":"p1","name":"天量解禁","cat":0,"x":60,"y":130,"detail":"%s 禁售承诺最后一日\n控股股东+一致行动人 33.2%%\n解禁前股价已跌至新低"%UNLOCK},
    {"id":"p2","name":"股东减持\n冲动分化","cat":1,"x":60,"y":380,"detail":"控股股东成本极低→深度浮盈·减持动力弱\n后期机构成本高于现价→套牢但有退出/止损需求"},
    {"id":"i1","name":"抛压大\n流动性弱","cat":2,"x":280,"y":255,"detail":"日均成交仅 %.0f 万股（极弱）\n解禁当日 %+.2f%%\n后 8 日累计 %+.2f%%\n悬崖在解禁之后"%(avg_vol,chg_unlock,post_cum)},
    {"id":"a1","name":"抢窗口\n控冲击","cat":3,"x":480,"y":255,"detail":"核心原则：趁抛压全面兑现前退出\n用组合工具压低价格冲击\n避免与先期大股东减持挤兑"},
    {"id":"s1","name":"阶段一\n解禁初期\n大宗竞价","cat":4,"x":680,"y":130,"detail":"先落袋一部分\n大宗一次性出·集中竞价走量\n趁窗口未塌前退出（折价 8–9 折）"},
    {"id":"s2","name":"阶段二\n大额大宗","cat":4,"x":680,"y":380,"detail":"待流动性放大、股价企稳\n做更大额、折让更小的大宗\n实现收益最大化"},
    {"id":"g1","name":"退出目标\n收益最大","cat":5,"x":880,"y":255,"detail":"在价格冲击与时间/窗口风险间权衡\n是否保留部分持仓拿管线上行"},
]
logic_links = [("p1","i1"),("p2","i1"),("i1","a1"),("a1","s1"),("a1","s2"),("s1","g1"),("s2","g1")]
def build_option_logic():
    gn=[{"id":n["id"],"name":n["name"],"category":n["cat"],"x":n["x"],"y":n["y"],
         "symbolSize":100 if n["cat"] in (3,5) else 90,
         "itemStyle":{"color":logic_cat_color[n["cat"]],"borderColor":"#fff","borderWidth":2},
         "label":{"show":True,"color":"#fff","fontSize":11,"fontWeight":"bold","lineHeight":14,"position":"inside","overflow":"break","width":78},
         "detail":n["detail"]} for n in logic_nodes]
    gl=[{"source":s,"target":t,"lineStyle":{"color":"#64748b","width":2.5,"curveness":0.08}} for s,t in logic_links]
    return {"backgroundColor":"transparent","tooltip":{"trigger":"item","formatter":"__TOOLTIP_FN__",
        "backgroundColor":"#1e293b","borderColor":"#334155","textStyle":{"color":"#e2e8f0","fontSize":12}},
        "legend":{"show":False},"series":[{"type":"graph","layout":"none","roam":False,"draggable":False,
        "edgeSymbol":["none","arrow"],"edgeSymbolSize":11,"label":{"show":True},"lineStyle":{"opacity":0.9},
        "categories":[{"name":c,"itemStyle":{"color":logic_cat_color[i]}} for i,c in enumerate(logic_cats)],
        "data":gn,"links":gl}]}
logic_table_rows = [
    ("减持额度","单日≤当日成交 20%","无明确限制(建议 15-20 倍日均)","无明确限制(建议 15-20 倍日均)","无明确限制"),
    ("折价情况","相对市价有一定折价","相对市价有一定折价","无折让(≥当日VWAP，<收盘)","市价无折让"),
    ("定价机制","买卖双方协商","券商簿记定价","当日 VWAP","随市价卖出"),
    ("信息保密","披露券商/规模/价","非公开簿记则不披露","不显示盘口","持续挂卖=抛压可见"),
    ("核心优势","流程简·一次性·确定性高","单笔规模大·综合价高·合规严","零折让·符合内部合规","费率低·操作简洁"),
    ("主要缺点","单笔规模受限","需律师/协议·需求不足风险","价格与时间无法提前确定","持续抛压·周期长"),
    ("方案采用","✓ 阶段一(集中竞价)","— 备选未主推","— 备选未主推","— 未主推"),
]
logic_table_html = ('<table class="cmp-tbl"><thead><tr><th>维度</th><th>大宗交易(非自动对盘)</th><th>旧股配售</th><th>VWAP 减持</th><th>二级市场直接减持</th></tr></thead><tbody>'
    + "".join("<tr>"+"".join("<td%s>%s</td>"%(" style='font-weight:700;color:#334155;text-align:left;background:#eef3f8'" if j==0 else "",c) for j,c in enumerate(r))+"</tr>" for r in logic_table_rows)
    + "</tbody></table>")
logic_section_html = """
    <section class="card" id="logic">
      <div class="card-head"><div class="card-num">⚙</div><div>
        <h2 class="card-title">方案逻辑拆解 · 解禁与减持决策框架</h2>
        <div class="card-sub">套用「方案逻辑拆解」逻辑链范式 · 适配 __COMPANY__ 少数股东视角</div></div></div>
      <div class="note"><span class="note-tag">逻辑链</span>左→右：解禁前提 → 股东减持冲动分化 → 抛压与流动性推论 → 行动原则（抢窗口/控冲击）→ 分阶段减持 → 退出目标。前提数据取自上方 c1/c3/c4/c5。</div>
      <div class="chart" id="dom_logic" style="height:500px;"></div>
      <div class="grid2" style="margin-top:14px;">
        <div style="overflow-x:auto;">__LOGIC_TABLE__</div>
        <div style="background:linear-gradient(135deg,#fef3c7,#fde68a);border:1px solid #fcd34d;border-radius:12px;padding:16px 18px;">
          <div style="font-size:14px;font-weight:800;color:#92400e;margin-bottom:6px;">关键判断（少数股东）</div>
          <ul style="margin:0 0 0 18px;font-size:13px;color:#78350f;line-height:1.75;">
            <li><b>框架成立</b>：天量解禁→抛压→抢窗口→组合工具，经验上站得住（同类案例解禁首日可跌 10%–30%，__COMPANY__后 8 日 __POSTCUM__%）。</li>
            <li><b>量化缺口</b>：卖多少、什么价、阶段一占比、价格底线，文件一句未给，须少数股东自己定。</li>
            <li><b>真要拍板的</b>：分阶段比例、止盈/止损点、税务外汇（人民币成本→港币退出）、是否保留部分持仓拿管线上行、弱流动性下 8–9 折预期。</li>
          </ul>
        </div>
      </div>
    </section>
""".replace("__LOGIC_TABLE__", logic_table_html).replace("__COMPANY__", COMPANY).replace("__POSTCUM__", "%.2f" % post_cum)

# ---------------- KPI / 导航 / 图表卡 HTML ----------------
kpi_html = "\n".join('      <div class="kpi"><div class="kpi-label">%s</div><div class="kpi-value">%s<span class="kpi-unit">%s</span></div><div class="kpi-delta %s">%s</div><div class="kpi-note">%s</div></div>' % (l,v,u,cls,dl,n) for l,v,u,dl,cls,n in KPI)
nav_html = "\n".join('        <a class="nav-item" href="#%s"><span class="nav-num">%02d</span>%s</a>' % (cid,i+1,title.split("，")[0]) for i,(cid,title,sub,ann,opt) in enumerate(charts))
nav_html += '\n        <a class="nav-item" href="#analysis"><span class="nav-num">★</span>少数股东要点</a>'
nav_html += '\n        <a class="nav-item" href="#logic"><span class="nav-num">⚙</span>方案逻辑拆解</a>'
sections=[]
for i,(cid,title,sub,ann,opt) in enumerate(charts):
    sections.append('    <section class="card" id="%s"><div class="card-head"><div class="card-num">%02d</div><div><h2 class="card-title">%s</h2><div class="card-sub">%s</div></div><button class="btn" data-save="%s">保存图片</button></div><div class="chart" id="dom_%s"></div><div class="note"><span class="note-tag">图表说明</span>%s%s</div></section>' % (cid,i+1,title,sub,cid,cid,ann, cmp_table_html if cid=="c6" else ""))
sections_html = "\n".join(sections)

# 少数股东分析卡（8 项框架要点，按需填充）
analysis_html = """
    <section class="card analysis" id="analysis">
      <div class="card-head"><div class="card-num">★</div><div><h2 class="card-title">少数股东投资要点</h2>
      <div class="card-sub">框架：股票财务分析（盈利/偿债/现金流/运营四维诊断）· AI尽调·财务分析专家（同业对标）· 方案逻辑拆解</div></div></div>
      <div class="grid2">
        <div class="acard"><h3>① 控制权与一致行动人</h3><p>TODO: 控股股东+一致行动协议+ESOP 合计占比；核查董事会席位、类别股东保护。</p></div>
        <div class="acard"><h3>② 估值判断</h3><p>TODO: 后期进场者浮亏比例；是否存在 down-round 信号。</p></div>
        <div class="acard"><h3>③ 解禁抛压</h3><p>TODO: 解禁日、大股东有无额外禁售；解禁后供给放量对股价压制。</p></div>
        <div class="acard"><h3>④ 流动性风险</h3><p>TODO: 日均成交、极端清仓周期；大额退出只能依赖大宗/协议转让及折价预期。</p></div>
        <div class="acard"><h3>⑤ 财务诊断（四维代理指标）</h3><p>未盈利公司用股价/融资/估值/流动性代理；显式标注"非传统财务比率，需审计财报补全"。</p></div>
        <div class="acard"><h3>⑥ 投资保护条款清单</h3><p>优先清算权、回购/对赌、反稀释、重大事项一票否决、信息权、随售/拖带、股息优先权。</p></div>
        <div class="acard"><h3>⑦ 可比公司与同业对标</h3><p>TODO: 可比公司较首发涨跌幅；补最新市值/营收/管线进度做横向锚定。</p></div>
        <div class="acard warn"><h3>⑧ 数据缺口与尽调清单</h3><p>三表与审计意见、现金与有息负债、研发管线、客户集中度、实控人及关联交易、优先权与对赌、是否港股通。</p></div>
      </div>
    </section>
"""

init_charts = "\n".join("  register('%s','dom_%s',%s);" % (cid,cid,json.dumps(opt,ensure_ascii=False)) for cid,_,_,_,opt in charts)
logic_js = json.dumps(build_option_logic(), ensure_ascii=False).replace('"__TOOLTIP_FN__"', TOOLTIP_FN)
init_js = init_charts + "\n  register('logic','dom_logic',%s);" % logic_js

page = (TEMPLATE
    .replace("__COMPANY__", COMPANY).replace("__TICKER__", TICKER)
    .replace("__EXCEL_NAME__", Path(EXCEL).name).replace("__GEN_DATE__", GEN_DATE)
    .replace("__KPI__", kpi_html).replace("__NAV__", nav_html)
    .replace("__SECTIONS__", sections_html + analysis_html + logic_section_html)
    .replace("__INIT__", init_js)
    .replace("__ECHARTS__", (OUT/"echarts.min.js").read_text(encoding="utf-8")))

target = OUT / ("%s投资分析看板.html" % COMPANY)
target.write_text(page, encoding="utf-8")
print("生成:", target, "大小: %.1f MB" % (target.stat().st_size/1024/1024))
print("股价:", first_close, "->", last_close, "跌幅 %.1f%%" % drop_pct, "市值约", mkt_cap, "亿")
