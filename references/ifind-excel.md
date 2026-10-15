# Excel iFinD：接口、口径与复现

先检查 Excel 的已连接 COM 加载项，再用实际证券的一条行情和一条财务公式验证。Windows 64位 Excel 可能通过 `ThsXlsSlug.ExcelAddin` 加载；本机路径与版本需当次发现，不把32位 XLL强行装到64位 Excel。禁止更改宏安全、读取或写出账户凭据；使用已有授权连接。

## 查询清单与脚本

`query_ifind.ps1 -QueryPath 查询JSON路径 -OutputDir 项目内路径 -TimeoutSeconds 180`

查询JSON为数组，每项包含 `Id, Code, Period, Field, Unit, AsOf, Window, Formula`。例如（日期按任务实际基准替换）：

```json
[
  {"Id":"Q001","Code":"603799.SH","Period":"2026-10-12","Field":"收盘价","Unit":"元/股","AsOf":"2026-10-12","Window":"","Formula":"=thsiFinD(\"ths_close_price_stock\",\"603799.SH\",\"2026-10-12\",100,\"\")"}
]
```

输出包括加载项状态、实际含公式且有数值缓存的XLSX，以及逐项JSON快照。状态 `numeric` 表示已返回数字，仍需财务与单位核验；`zero_review` 表示零值需判断业务含义；error/timeout/missing/text_review均不当作正常值。脚本超时后停止重试并关闭自建工作簿，保留能取得的证据，不改其他用户文件。

## 已实测公式模式

下列是本环境曾返回有效数据的函数模式，执行时替换证券、日期、年度并复核当次厂商定义。参数次序不同，不可统一拼接。

| 指标 | 模式 | 原始单位 |
|---|---|---|
| 收盘价 | `=thsiFinD("ths_close_price_stock","代码","交易日",100,"")` | 元/股 |
| 总股本 | `=thsiFinD("ths_aghj_stock","代码","交易日")` | 股 |
| 总市值 | `=thsiFinD("ths_zsz_stock","代码","交易日")` | 元 |
| 合并利润表 | `=thsiFinD("ths_lrb_stock","代码",项目号,"报告期",100)` | EPS元/股，金额元 |
| 研发费用 | `=thsiFinD("ths_rad_cost_sum_stock","代码","报告期",100)` | 元 |
| 实际加权ROE | `=thsiFinD("ths_roe_wgt_publish_stock","代码","报告期")` | % |
| 收入一致预期 | `=thsiFinD("ths_fore_mob_income_mean_consensus_stock","代码",年度,"截至日")` | 元 |
| 归母净利润一致预期 | `=thsiFinD("ths_fore_np_mean_consensus_stock","代码",年度,"截至日")` | 元 |
| EPS一致预期 | `=thsiFinD("ths_fore_eps_mean_consensus_stock","代码",年度,"截至日")` | 元/股 |
| EPS机构数 | `=thsiFinD("ths_eps_fore_org_num_consensus_stock","代码",年度,"截至日")` | 家 |
| 历史时点预测PE | `=thsiFinD("ths_fore_pe_his_stock","代码","交易日",年度)` | 倍 |
| 预测PEG | `=thsiFinD("ths_fore_peg_stock","代码",年度,"交易日")` | 倍 |

利润表项目号：2051营业收入、2047营业成本、2012归母净利润、2044销售费用、2010管理费用、2002财务费用、2015基本EPS、2043稀释EPS。先与最新披露复核，不仅凭项目号猜含义。上述一致预期均值字段在本机厂商定义为距截至日固定六个月内预测报告均值，不能随意追加180参数。

## 口径核验

1. 金额原值元；展示亿元除以1e8、百万元除以1e6；EPS、股数与百分比单独处理。
2. 同日价格×总股本应等于市值。PE应与同日市值/对应预测归母利润互核，差异先检查股本、样本和定义。
3. 基本EPS用加权平均股本，不强行改成期末股本反推。机构EPS均值与利润均值/当前股本可存在样本差异，保留原值。
4. 收入预测字段标题曾使用营业收入，定义/算法又出现主营业务收入。执行时查询厂商说明或匹配个券数据；未核齐则A端营业收入、E端平台预测收入分列标注，跨A/E收入增速不作为核心论据。
5. ROE/ROA核对百分比尺度、净利润归属、权益/资产平均方法，无法核齐时正式表留“—”。不用未来期末股本假设冒充机构预测股本。
6. 机构计数异常0不能当作零家覆盖，也不能用EPS计数替代收入/利润计数。取数时间不能作为机构更新日期。
7. 年度、半年与同期分别处理；半年直接×2只允许明确标记的情景假设，不当成真实预测。
8. 原始公式/缓存与静态数据底稿两份保留；同一字段出现异常先限定一次合理刷新，超时记录原因和手动步骤。

## 手动补取

在已有授权的Excel iFinD功能区用指标搜索/函数查询选择证券，输入报告期/年度与统计截至日，按任务所需货币单位导出。记录字段完整名称、参数、样本窗口、原值及取数时间。等待“抓取中”变为数值后另存原始公式，再复制值形成静态快照；用公告实际值、股价×股本、市值/利润和分项合计交叉复核。失败时提供该项目完整查询清单，不能只写“自行登录数据库”。

厂商帮助：[Excel插件常见问题](https://ftcm.10jqka.com.cn/thsft/iFindService/docCenter?ID=EXCEL20151028)、[官方接口示例](https://quantapi.10jqka.com.cn/gwstatic/static/ds_web/quantapi-web/example.html)。本机函数查询/安装包定义为字段参数依据；使用非Excel接口须另行验证，不把本技能公式直接当SDK调用。
