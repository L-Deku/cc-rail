# RecoLearning 学习库工具

学习库设计见 `docs/superpowers/specs/2026-07-18-工程量学习库-design.md`。
库在 192.168.2.213 的 SQL Server 上,库名 RecoLearning。

## 表结构

当前共 12 张表：

- 原始流水：`BindingLog`
- 推荐核心：`QuantityAlias`、`QuotaBox`、`QuotaBoxTarget`、`SignatureBoxMap`
- 数量公式：`QuantityFormulaRule`、`QuantityFormulaOperand`
- 条目与模板参考：`SignatureEntryMap`、`EntryQuota`、`ChapterEntry`、`EngineeringTemplate`、`SheetTemplateRow`

推荐预览直接读取推荐核心、数量公式和条目映射；`BindingLog` 是审计及全量重建来源。
`SheetTemplateRow` 是三期“一表一模板”的预留原料，目前只由全量重算脚本写入，插件尚未读取，保留不删；同一组件跨条目时按目标条目生成多行，不能再用组件组的首个条目代表整组。
`EngineeringTemplate` 仅归集前两位为数字且其余字符只含数字或横杠的分类条目码。普通定额和 `SF` 分别按自身目标条目归集，纯 `SF` 框也会生成设备购置费条目行；`ZLF`、`LF`、材料和 `SH` 不独立扩展专业范围。一个 `box_id` 因此可以同时出现在同专业的安装工程费和设备购置费范围中。

## 脚本

| 脚本 | 用途 |
|---|---|
| Initialize-RecoLearning.ps1 | 建库建表 |
| Import-JsonlLibraries.ps1 | 默认只重载章节树/条目定额库；旧映射流水必须显式一次性迁移 |
| Import-ExcelLinks.ps1 | 收割 ExcelLinks\*.xml 绑定明细(回连项目库解析条目号) |
| Rebuild-Aggregates.ps1 | 在同一事务内由流水全量重算聚合表；`-DryRun` 全程演练后回滚，但同样持有 `BindingLog` 独占锁 |
| Get-LearningStats.ps1 | 体检报告 |

## 标准流程

新机器初始化 / 定期整理的安全序列如下。`-DryRun` 不是普通只读检查：它与正式执行一样获取
`BindingLog` 的 `TABLOCKX` 独占锁，在事务内执行 `TRUNCATE` 和批量写入，最后才回滚；两者都必须放在冻结绑定写入的维护窗口内。

    pwsh -File Initialize-RecoLearning.ps1
    pwsh -File Import-JsonlLibraries.ps1
    pwsh -File Import-ExcelLinks.ps1
    pwsh -File Rebuild-Aggregates.ps1 -DryRun
    # 核对演练统计后，维护窗口内再正式执行：
    pwsh -File Rebuild-Aggregates.ps1
    pwsh -File Get-LearningStats.ps1

执行前后都应记录 `SELECT ISNULL(MAX(id),0) FROM dbo.BindingLog`。只有水位一致时才比较演练与正式重算的聚合行数；期间若有新流水，两次行数不保证相同。

旧 `mapping-boxes.jsonl` / `learning.jsonl` 仅在确认尚未双写时迁移一次：

    pwsh -File Import-JsonlLibraries.ps1 -ImportBindingHistory -SourceId <本次唯一来源>

脚本会拒绝重复使用同一 `SourceId`，避免学习计数翻倍。

## 插件双写

插件绑定/铺量成功后同时尝试本机 `mapping-boxes.jsonl` 和中央 SQL，入口为
`tools\RecoExpandPanel\LearningDbFeature.cs`。SQL 内的 `BindingLog`、推荐核心、
数量公式和 `SignatureEntryMap` 在同一事务提交，因此下一次推荐预览即可读取；
并发死锁、唯一键或短暂网络错误会整笔重试一次，失败不会永久停用后续学习，
也不会影响本机绑定和 JSONL 备份。

同量纲标准单位（如 `kg→t`、`m³→10m³`）始终按当前 Excel 单位和当前版本定额单位实时换算。
跨量纲关系保存为 `V0/V1...` 参数公式，例如 `V0*0.2` 或
`V0*V1*V2*V2*3.14`；推荐时只在当前表同章节、邻近行中按精确名称找到唯一参数后计算，
缺参数、重名或单位不符时不自动勾选。
