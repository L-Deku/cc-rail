# Migrate2020EstimateTo2024

把 `RecoData2020` 中的 2020 概算/估算定额迁移到 `RecoData2024`，并将消耗中的人材机电算代号替换为 2024 资源；2024 缺失的材料/机械按补充资源预分配编号。

默认流程是先预检，不直接写库。

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Users\谢刚\Desktop\自动预算\tools\Migrate2020EstimateTo2024\build.ps1"
& "C:\Users\谢刚\Desktop\自动预算\tools\Migrate2020EstimateTo2024\bin\Migrate2020EstimateTo2024.exe" Precheck
& "C:\Users\谢刚\Desktop\自动预算\tools\Migrate2020EstimateTo2024\bin\Migrate2020EstimateTo2024.exe" Verify
```

`Apply` 需要显式确认文件，防止误写数据库：

```powershell
Set-Content -LiteralPath "C:\Users\谢刚\Desktop\自动预算\tools\Migrate2020EstimateTo2024\reports\apply.confirm.txt" -Encoding UTF8 -Value "APPLY RecoData2020 estimate-to-2024 migration"
& "C:\Users\谢刚\Desktop\自动预算\tools\Migrate2020EstimateTo2024\bin\Migrate2020EstimateTo2024.exe" Apply --confirm "C:\Users\谢刚\Desktop\自动预算\tools\Migrate2020EstimateTo2024\reports\apply.confirm.txt"
```

输出文件位于 `reports/`：

- `precheck-summary.txt`
- `resource-map.csv`
- `missing-resources.csv`
- `name-review.csv`
- `unit-differences.csv`
- `conflicts.csv`
- `rollback.sql`

工具只写当前工作区文件；只有 `Apply` 会写 `RecoData2024` 数据库。

## 注意事项

- 预检和验证需要访问 `192.168.2.13,1433`；在受限沙箱中可能出现 SQL “拒绝访问”，需用已授权的非沙箱命令运行。
- 工具必须按 x86 编译和运行，因为需要反射加载 2024 主程序中的 `RecoNet.Security`。
- 2020 消耗串解密后按“前 9 位电算代号 + 后续消耗量”解析，不能按固定小数位切分。
- 代码 `10`、`11` 在接触网概算中作为特殊人工代码保留，不按材料或机械补充。
- 资源匹配顺序为：原始名称完全一致、规范化名称一致、唯一高置信相似匹配；低置信候选不自动替换，只写入 `missing-resources.csv` 的候选列供人工审核。
- `resource-map.csv` 包含每个 2020 电算代号的替换前名称、替换后名称、单位和匹配分数。
- `name-review.csv` 只列出规范化或相似名称匹配的行，适合人工重点审核。
- `missing-resources.csv` 只放没有可靠 2024 对应项、准备按补充材料/补充机械写入的资源；其中 `best_candidate_*` 是找到但未自动采用的最佳候选。
- 机械候选会额外检查动力类型、设备类别、单筒/双筒、慢速/快速等关键词；候选相似度低于 60 时仍写入兜底 `best_candidate_code`，但 `reason` 会明确标注为低置信、仅供人工审核。
- 当前版本会对 `missing-resources.csv` 做兜底候选搜索，尽量保证每一行都有 `best_candidate_code`；低置信候选会在 `reason` 中标注“仅供人工审核”，不能直接作为自动替换依据。
