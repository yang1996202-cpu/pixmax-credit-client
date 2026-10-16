# PIXMAX Credit Client - 积分消耗分析（客户/老板视角）

一个 WorkBuddy / Claude Code Skill，把客户导出的 PIXMAX 积分流水 CSV，一键变成**可以直接发给客户老板看**的分析 Excel。

> 给自己（PIXMAX 侧）看经营/催单分析？用姊妹仓库 [pixmax-credit-business](https://github.com/yang1996202-cpu/pixmax-credit-business)。

## 它能做什么

**输入：** PIXMAX 后台导出的积分流水 CSV（企业版 25 列 / 团队版 15 列，自动识别）

**输出：** 9 张表的 Excel，首页「老板重点」结论先行

| 表 | 内容 |
|---|---|
| 老板重点 | 一句话结论 + 三件事（还能烧几天 / 谁花最多 / 能省多少）+ 异常清单 + 原因拆解 + 建议动作 |
| 总览 | KPI + 按周小表 + 按天/按周趋势图 |
| 按天汇总 / 按周汇总 | 日期/星期/积分/次数/金额/占比 |
| 按人汇总 | 排名/消耗/次数/金额/主力模型/占比 |
| 人×模型 | 谁在用什么模型（调用次数 + 花费金额双矩阵，热力图） |
| 按模型汇总 | 类别/单次积分/配置档位数 |
| 充值对账 | 充值明细 + 消耗进度 + 实充是否用尽 |
| 原始流水 | 全量明细，可筛选 |

## 使用场景

- 客户问「积分怎么花这么快」→ 出这张表回答他
- 续费沟通前的用量复盘材料
- 给客户提「怎么省积分」的建议（含降级模型的理论节省估算）

## 用法

```bash
python3 scripts/analyze.py <csv路径> --view client [--out <输出目录>]
```

依赖：Python 3.10+，`openpyxl`（`pip install openpyxl`）。

## 关键口径（表内均已标注）

- **调用次数** = 生成任务数（一次任务计 1 次）
- **单次积分** = 消耗积分 ÷ 调用次数（积分/次，非金额）
- **高配调用** = 单次 ≥ 400 分；**极端** = 单次 ≥ 800 分
- **降级省钱** = PIXDANCE_2_5 降级到 FAST 的理论节省，FAST 单次均价优先取客户实测值，无数据时用 90 分/次经验值
- 金额为平台标价折算（约 0.0266 元/积分），非客户实付

## 格式兼容

| 格式 | 特征 | 处理 |
|---|---|---|
| 企业版 | 25 列，含 `AI模型`/`子账号ID` | 按子账号汇总，模型相关表全出 |
| 团队版 | 15 列，含 `账号姓名` | 按操作人汇总，自动跳过模型表 |
| 企业版·未开子账号 | `子账号ID` 全空 | 归主账号，显示「主账号（未开子账号）」，不再出现「（未知）」 |

## 自测

```bash
python3 scripts/analyze.py examples/sample-enterprise.csv --view client --out /tmp/demo
python3 scripts/analyze.py examples/sample-team.csv --view client --out /tmp/demo
```

样例已脱敏（`examples/`），可放心使用。

## License

MIT
