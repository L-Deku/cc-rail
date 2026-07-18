# RecoLearning 学习库工具

学习库设计见 `docs/superpowers/specs/2026-07-18-工程量学习库-design.md`。
库在 192.168.2.213 的 SQL Server 上,库名 RecoLearning。

## 脚本(全部幂等,可重复执行)

| 脚本 | 用途 |
|---|---|
| Initialize-RecoLearning.ps1 | 建库建表 |
| Import-JsonlLibraries.ps1 | 导入 RecoQuotaData 四个 jsonl(章节树/条目定额库全量重载;映射框/扶正流水追加) |
| Import-ExcelLinks.ps1 | 收割 ExcelLinks\*.xml 绑定明细(回连项目库解析条目号) |
| Rebuild-Aggregates.ps1 | 由流水全量重算聚合表 —— "定期整理"入口 |
| Get-LearningStats.ps1 | 体检报告 |

## 标准流程

新机器初始化 / 定期整理都执行同一序列:

    pwsh -File Initialize-RecoLearning.ps1
    pwsh -File Import-JsonlLibraries.ps1
    pwsh -File Import-ExcelLinks.ps1
    pwsh -File Rebuild-Aggregates.ps1
    pwsh -File Get-LearningStats.ps1

## 插件双写

插件绑定/铺量成功后向 BindingLog 追加流水(source 前缀 `plugin:`),
入口 `tools\RecoExpandPanel\LearningDbFeature.cs`。写库失败只记插件日志并在
本进程内停用,不影响绑定;之后由 Rebuild-Aggregates 把流水折算进聚合表。
