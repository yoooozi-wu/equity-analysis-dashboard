# 「方案逻辑拆解」框架规格

把一份"减持 / 退出 / 解禁方案"或"提案逻辑"拆成一张**逻辑链图 + 工具对比表 + 关键判断**的看板块。适配少数股东 / PE 退出决策视角，可套用于任意港股 / A 股解禁减持场景。

---

## 1. 逻辑链图（ECharts graph）

### 1.1 六类节点（category）

| 类 | 含义 | 颜色 |
|----|------|------|
| 解禁前提 | 天量解禁、禁售到期 | `#3b82f6` 蓝 |
| 股东冲动 | 各股东减持动机分化（浮盈 vs 套牢） | `#f59e0b` 橙 |
| 抛压推论 | 抛压集中 + 流动性弱 → 股价下行 | `#ef4444` 红 |
| 行动原则 | 抢窗口、控价格冲击 | `#10b981` 绿 |
| 减持阶段 | 阶段一（解禁初期集中竞价+大宗）/ 阶段二（企稳后大额大宗） | `#14b8a6` 青 |
| 退出目标 | 退出收益最大化 + 风险可控 | `#8b5cf6` 紫 |

### 1.2 通用链路（左→右）

```
[解禁前提]──┐
            ├─→[抛压推论]─→[行动原则]─┬─→[阶段一]─┐
[股东冲动]──┘                        └─→[阶段二]─┴─→[退出目标]
```

每条 link：`{source, target, lineStyle:{color:'#64748b', width:2.5, curveness:0.08}}`。

### 1.3 节点 / 坐标设计（适配约 910px 内宽）

- 符号尺寸：行动原则与目标用 `100`，其余 `90`。
- label：`{show:true, color:'#fff', fontSize:11, fontWeight:'bold', lineHeight:14, position:'inside', overflow:'break', width:78}`。
- itemStyle 加白边：`{color: 类色, borderColor:'#fff', borderWidth:2}`。
- 坐标（x 跨度 60–880，y 130/255/380，避免裁切）：
  - 前提类 x=60；推论/原则/目标 x=280/480/880（居中 y=255）；阶段类 x=680（y=130/380）。
- 每个节点带 `detail` 字段（多行 `\n`），作为 tooltip 详情；节点名则压成 2–3 个短行。

### 1.4 tooltip formatter 函数注入（关键）

`json.dumps` 无法输出真正的 JS 函数，必须在 json 后用字符串替换：

```python
TOOLTIP_FN = "function(p){ if(p.data && p.data.detail){ return p.data.detail.replace(/\\n/g,'<br/>'); } return p.name; }"
opt = {... "tooltip": {"trigger":"item", "formatter":"__TOOLTIP_FN__", ...}}
js = json.dumps(opt, ensure_ascii=False).replace('"__TOOLTIP_FN__"', TOOLTIP_FN)
```

> 注意 `TOOLTIP_FN` 在 Python 源里 `\\n` 表示正则里的 `\n`（换行匹配）；`detail` 字符串用真实 `\n` 换行。

### 1.5 graph option 模板

```python
graph_option = {
  "backgroundColor": "transparent",
  "tooltip": {"trigger":"item","formatter":"__TOOLTIP_FN__",
              "backgroundColor":"#1e293b","borderColor":"#334155",
              "textStyle":{"color":"#e2e8f0","fontSize":12}},
  "legend": {"show": False},
  "series": [{
    "type":"graph","layout":"none","roam":False,"draggable":False,
    "edgeSymbol":["none","arrow"],"edgeSymbolSize":11,"label":{"show":True},
    "lineStyle":{"opacity":0.9},
    "categories":[{"name":c,"itemStyle":{"color":logic_cat_color[i]}} for i,c in enumerate(cats)],
    "data": gn, "links": gl,
  }],
}
```

注册时与看板其他图一致：`register('logic','dom_logic', <js>)`（js 已是替换后的 JSON 字符串，不要再 json.dumps）。

---

## 2. 港股减持工具对比表（4 维度 × 7 行）

直接复用（适用于任意港股减持场景）。`<table class="cmp-tbl">` 风格，第 0 列维度名加粗左对齐。

| 维度 | 大宗交易(非自动对盘) | 旧股配售 | VWAP 减持 | 二级市场直接减持 |
|------|------|------|------|------|
| 减持额度 | 单日≤当日成交 20% | 无明确限制(建议 15-20 倍日均) | 无明确限制(建议 15-20 倍日均) | 无明确限制 |
| 折价情况 | 相对市价有一定折价 | 相对市价有一定折价 | 无折让(≥当日VWAP，<收盘) | 市价无折让 |
| 定价机制 | 买卖双方协商 | 券商簿记定价 | 当日 VWAP | 随市价卖出 |
| 信息保密 | 披露券商/规模/价 | 非公开簿记则不披露 | 不显示盘口 | 持续挂卖=抛压可见 |
| 核心优势 | 流程简·一次性·确定性高 | 单笔规模大·综合价高·合规严 | 零折让·符合内部合规 | 费率低·操作简洁 |
| 主要缺点 | 单笔规模受限 | 需律师/协议·需求不足风险 | 价格与时间无法提前确定 | 持续抛压·周期长 |
| 方案采用 | ✓ 阶段一(集中竞价) | — 备选未主推 | — 备选未主推 | — 未主推 |

---

## 3. 关键判断黄条（少数股东视角模板）

三句话，结构固定，内容按公司填：

1. **框架成立**：天量解禁→抛压→抢窗口→组合工具，经验上站得住（可引同类解禁案例的量级作参照，如解禁首日 10%–30% 级别的下跌）。
2. **量化缺口**：卖多少、什么价、阶段一占比、价格底线，提案/文件一句未给，须少数股东自己定。
3. **真要拍板的**：分阶段比例、止盈/止损点、税务外汇（人民币成本→港币退出）、是否保留部分持仓拿管线上行、弱流动性下 8–9 折预期。

> 提醒：若来源是券商"揽活提案"而非投资人"退出决策书"，务必在判断里点明——框架可借鉴，但量化与拍板权在投资人自己手里。

---

## 4. 装配位置

放在看板 `__SECTIONS__` 末尾（原 6 图 + 少数股东分析之后）。结构：

```html
<section class="card" id="logic">
  <div class="card-head"><div class="card-num">⚙</div><div>
    <h2 class="card-title">方案逻辑拆解 · 解禁与减持决策框架</h2>
    <div class="card-sub">套用「方案逻辑拆解」逻辑链范式 · 适配 XX 少数股东视角</div></div></div>
  <div class="note"><span class="note-tag">逻辑链</span>左→右：解禁前提 → 股东冲动分化 → 抛压推论 → 行动原则 → 分阶段减持 → 退出目标。前提数据取自上方 c1/c3/c4/c5。</div>
  <div class="chart" id="dom_logic" style="height:500px;"></div>
  <div class="grid2" style="margin-top:14px;">
    <div style="overflow-x:auto;">__LOGIC_TABLE__</div>
    <div style="...黄条渐变背景...">关键判断三句话</div>
  </div>
</section>
```

导航与页脚各加一项（⚙ 方案逻辑拆解）。
