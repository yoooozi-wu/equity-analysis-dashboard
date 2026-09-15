# 工作流详解：Excel → 自包含投资分析看板

本文件是 `equity-analysis-dashboard` skill 的详细操作手册。配合 `scripts/build_dashboard.py`（可运行模板）与 `references/logic-chain.md`（逻辑链框架）使用。

---

## 0. 运行环境

- Python 3（需 pandas / numpy / openpyxl）。建议用虚拟环境，切勿全局 `pip install`。
- Node.js（跑 `validate.js` 校验）。
- 输出目录建一个 `out/`，把 `assets/echarts.min.js` 复制进去，构建脚本读取它内嵌进 HTML。

---

## 1. 探查 Excel

```python
import pandas as pd
xl = pd.ExcelFile(EXCEL)
print(xl.sheet_names)
for sh in xl.sheet_names:
    d = pd.read_excel(EXCEL, sheet_name=sh, header=None)
    # 打印前几行，识别表头/日期列位置
    print(sh, d.shape)
```

常见 sheet 形态（以港股 18A 股权梳理为例）：
- `上市后交易情况`：第 0 行标题、第 1 行表头；日期在 col0（`2025-08-15` 这类字符串），收盘价 col3，成交量 col5（单位股，绘图时 ÷1e4 转万股）。
- `历史融资情况`：轮次 + 每股成本，常需手工录入 `round_cost` 列表（PDF/文档里更全）。
- `股东-股数&成本`：股东名 + 持股数 + 成本；先算合计占比。
- `可比公司N`：子表 col1=收盘价、col2=成交量；**列序与主表不同**，务必先核对。
- `解禁日/港股通`：可能为空，按业务确认。

识别日期行的通用写法：
```python
for i in range(1, len(d)):
    v = d.iloc[i, 0]
    if pd.notna(v) and isinstance(v, str) and v[0].isdigit():
        dt = str(v)[:10]          # 仅取 YYYY-MM-DD
        c  = float(d.iloc[i, 3])  # 收盘价
        vol= float(d.iloc[i, 5])  # 成交量
```

---

## 2. 六类核心图表 option 模板

所有 option 用 `json.dumps(opt, ensure_ascii=False)` 注入。短中文名是硬要求（见 Known pitfalls）。

### c1 上市后交易（双轴）
- `yAxis[0]` 收盘价（name 含单位"港元"），`yAxis[1]` 成交量（`splitLine.show:false`）。
- series：收盘价 `line`（`showSymbol:false`）+ 成交量 `bar`（`yAxisIndex:1`）。
- 用 `markLine` 标解禁日：`data:[{"xAxis": UNLOCK}]`，红色虚线。

### c2 融资估值轨迹
- 单 `line`，x=轮次标签（interval:0, rotate:20），y=每股成本。
- `markPoint` 标 B+轮峰值、老股折价（down-round 信号）。

### c3 股东结构（横向条形）
- `xAxis` 数值（持股占比%），`yAxis` 类目（股东名，倒序使最大在顶）。
- series：`bar`，每根用股东专属色（`itemStyle.color` 数组）。

### c4 持股成本 vs 现价
- 横向 `bar`，x=成本(元/股)，`markLine` 画现价参考线（`data:[{"xAxis": CURRENT}]`，红色）。
- 图注点明：现价线以上=套牢、以下=浮盈。

### c5 解禁窗口（前8+当日+后8）
- `xAxis` 17 个交易日（axisLabel interval:0, rotate:42, fontSize:10）。
- series：收盘价 `line` + 当日涨跌幅 `bar`（`yAxisIndex:1`，首点 `None`）。
- `markLine` 解禁日；`markArea` 浅色高亮"前 8 日"区间。
- 关键结论模板：「解禁前 8 日仅累计 X%；**解禁当日收 Y、单日 Z%**；随后 8 日再累计 W%——悬崖在解禁之后」。

### c6 可比公司（较首发涨跌幅）
- 横向 `bar`，本公司红色（`#c62828`）、同业蓝（`#1a6fc4`）。
- 配套 `<table class="cmp-tbl">`：公司 / 上市日 / 解禁日 / 首发收 / 最新收 / 较首发。

---

## 3. 装配单文件 HTML 看板

- HTML 模板里 CSS 用了 `%` 单位（`width:100%` 等），所以**不能用 `%` 占位符格式化**；改用 `__TOKEN__` 链式 `.replace()`：
  - `__GEN_DATE__` / `__KPI__` / `__NAV__` / `__SECTIONS__` / `__INIT__` / `__ECHARTS__`。
- `__ECHARTS__` = `assets/echarts.min.js` 全文（仅一份，所有图共用）。
- 图表注册用统一函数：
  ```js
  function register(id, domId, option) {
    delete option.title; option.backgroundColor = 'transparent';
    var chart = echarts.init(document.getElementById(domId));
    chart.setOption(option);
    new ResizeObserver(function(){ chart.resize(); }).observe(dom);
    window.addEventListener('resize', function(){ chart.resize(); });
  }
  ```
- 每图卡带「保存图片」按钮（监听 `.btn[data-save]`，`chart.getDataURL({type:'png',pixelRatio:2})` 触发下载）。
- 少数股东分析卡（8 项框架要点）：控制权与一致行动人、估值判断、解禁抛压、流动性风险、财务诊断（代理指标）、投资保护条款清单、可比公司比对、数据缺口与尽调清单。

---

## 4. 融入「方案逻辑拆解」框架

参考 `references/logic-chain.md`。在 `__SECTIONS__` 末尾追加一块 `id="logic"`：
- 逻辑链图（ECharts `graph`，layout:'none'）+ 港股减持工具对比表（大宗/旧股配售/VWAP/二级市场，4 维度 7 行）+ 关键判断黄条（少数股东视角）。
- 工具对比表行定义见 logic-chain.md；可直接复用。
- 导航与页脚同步加一项。

---

## 5. 校验（必做）

见 `scripts/validate.js`。它从生成的 HTML 抽取含 `function register` 的 init 脚本，mock `echarts.init`（返回带 `setOption`/`getDataURL`/`resize` 的桩）、`document`/`window`/`ResizeObserver`，`vm.runInContext` 跑一遍，并断言：
- 语法通过（`new vm.Script`）。
- `setOption` 调用数 = 图表数（6 原图 + 逻辑链图 = 7）。
- 逻辑链 `series[0].type === 'graph'`，`data.length === 7`，`links.length === 7`。
- `tooltip.formatter` 是 `function`（非字符串）——验证函数注入成功。

通过后删掉临时 `_validate.js`。

---

## 6. Known pitfalls（展开）

1. **长中文名重叠**：系列名 / 轴名过长会在图例、坐标轴互相堆叠截断。→ 短名覆盖（毛利率→毛利；经营/投资/筹资；研发/销售/管理费率；阶段一/阶段二）。≥3 系列的图把 legend 放**左侧垂直**（`orient:'vertical'`，给 `grid.left` 让出 18%），≤2 系列放底部居中 + `type:'scroll'`。左侧竖写 `graphic` 轴名要与左图例错开，避免重叠。
2. **子表列序差异**：可比公司子表 col1=收盘价、col2=成交量；主表 col3=收盘、col5=成交量。误读列会拿成交额当收盘价，涨跌幅全错。→ 提取前打印列名/前几行确认。
3. **graph tooltip 函数注入**：`json.dumps` 会把 Python 里的 `function(){...}` 字符串当普通字符串输出（带引号），ECharts 不会执行。→ option 里 formatter 写成占位符 `"__TOOLTIP_FN__"`，`json.dumps` 后 `.replace('"__TOOLTIP_FN__"', TOOLTIP_FN)`，其中 `TOOLTIP_FN = "function(p){...}"`（JS 字符串里 `/\n/g` 用 `\\n` 表示换行匹配）。
4. **graph 固定坐标不缩放**：`layout:'none'` 的 x/y 是像素，不自动适配容器。节点中心 + 符号半径（如 50）必须 ≤ 卡片内宽（约 910px），否则右侧被裁。→ 整体 x 跨度收到 60–880，符号 90/100，label `position:'inside'`、`fontSize:11`、`overflow:'break'`、`width:78`。若文字仍看不清：放大符号、压缩换行、加深连线色（`#64748b`）。
5. **解禁日探测**：常是「禁售承诺最后一日」而非首个日期。→ 按文档业务含义确认，勿机械取首个 date-like 值。
6. **18A 无三表**：用股价/融资/估值/流动性作四维诊断代理指标，图注与卡片显式标注「非传统财务比率，需以审计财报补全」。
7. ** `%` 占位符冲突**：模板 CSS 用 `%`，故装配用 `__TOKEN__` 而非 `%s` / `format()`。
8. **中文文件名 bash 通配失效**：`rm *_hash.html` 在中文名下 glob 不到；改用 Python `os.listdir` + 精确后缀删除。

---

## 7. 交付清单

- `xxx投资分析看板.html`（自包含，约 1MB，离线可用）。
- `build_xxx.py`（构建脚本，内联提取，作为以后重跑/改动的入口）。
- `out/echarts.min.js`（内嵌库副本）。
- 可选：`charts.json`（若用 smart-charts 单图起手）。
