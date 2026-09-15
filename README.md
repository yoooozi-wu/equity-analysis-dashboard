# equity-analysis-dashboard

把「股权梳理 / 财务 / 交易 / 可比公司」类 Excel，转换成**单文件、离线、自包含**的 HTML 投资分析看板（内嵌一份 ECharts，无需联网）。看板以**少数股东 / PE 视角**组织，并内置可复用的**「方案逻辑拆解」框架**（逻辑链图 + 港股减持工具对比 + 关键判断），服务解禁 / 减持 / 退出决策。

> 源自一套真实的私募少数股东分析工作流（港股 18A 生物医药场景）。仓库内示例数据已脱敏为占位符。

## 适用场景

- 你有一个股权 / 财务梳理 Excel，要做「投资分析看板 / 数据可视化 / 少数股东视角分析 / 解禁减持逻辑拆解 / 方案逻辑拆解」。
- 关心退出、解禁抛压、流动性、控制权、可比公司。

## 目录结构

```
equity-analysis-dashboard/
├── SKILL.md                  # 触发条件 + 工作流概要 + 资源索引
├── README.md                 # 本文件（GitHub 说明）
├── references/
│   ├── workflow.md           # 完整流程、pandas 提取模式、6 类图表 option、合并技巧、已知坑
│   └── logic-chain.md        # 方案逻辑拆解框架：6 类节点、graph option 模板、函数注入、坐标适配
├── scripts/
│   ├── build_dashboard.py    # 可运行模板（改 CONFIG/提取段即可用于新公司）
│   ├── validate.js           # Node 校验桩：语法 + 全部 setOption + 逻辑链完整性
│   └── fetch_echarts.py      # 下载 ECharts 到 assets/（首次使用跑一次）
└── assets/
    └── echarts.min.js        # ECharts 5.4.4 离线库（约 1MB，未入库，由 fetch_echarts.py 下载）
```

## 快速开始

```bash
# 0) 首次：下载 ECharts 到 assets/（约 1MB，未入库）
python scripts/fetch_echarts.py

# 1) 准备：把 echarts.min.js 复制到输出目录（构建脚本会内嵌它）
cp assets/echarts.min.js ./out/

# 2) 改 scripts/build_dashboard.py 顶部的 CONFIG / EXCEL 路径与各 sheet 提取段
#    （不同公司表头/列序不同，见 references/workflow.md 的"探查 Excel"）

# 3) 运行（用隔离的 python 环境）
python scripts/build_dashboard.py        # -> ./out/<公司>投资分析看板.html

# 4) 校验（Node）
node scripts/validate.js ./out/<公司>投资分析看板.html
```

## 看板包含

- **6 张核心图**：上市后交易（股价+成交量双轴）、融资估值轨迹、股东结构、持股成本 vs 现价、解禁窗口（前 8 日+当日+后 8 日）、可比公司（较首发涨跌幅）。
- **少数股东分析卡**（8 项框架要点）：控制权与一致行动人、估值判断、解禁抛压、流动性风险、财务诊断（代理指标）、投资保护条款清单、可比对标、数据缺口与尽调清单。
- **方案逻辑拆解**：逻辑链图（ECharts graph，7 节点 / 6 类）+ 港股减持工具对比表（大宗/旧股配售/VWAP/二级市场）+ 关键判断黄条。

## 关键纪律

- 所有数值由脚本从原始 Excel **确定性提取**，缺失资料不推测、不编造。
- 图注必须与图上实际绘制的系列一致。
- 港股 18A 未盈利、常无三表 → 用"股价 / 融资 / 估值 / 流动性"作四维诊断代理指标，并显式标注"非传统财务比率"。
- 长中文名会挤爆图例/坐标轴 → 用短显示名；graph 固定坐标不缩放 → 整体跨度 + 符号半径 ≤ 卡片内宽（约 910px）。

## 校验点

- `node --check` 通过；`validate.js` 跑通全部 `setOption`，逻辑链节点 = 7、连线 = 7，tooltip formatter 为真实函数。

## License

MIT（示例代码与文档，可用于任意内部 / 商业分析，不构成投资建议）。
