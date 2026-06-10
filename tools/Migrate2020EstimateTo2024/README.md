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
- `unit-differences.csv`
- `conflicts.csv`
- `rollback.sql`

工具只写当前工作区文件；只有 `Apply` 会写 `RecoData2024` 数据库。

## 注意事项

- 预检和验证需要访问 `192.168.2.13,1433`；在受限沙箱中可能出现 SQL “拒绝访问”，需用已授权的非沙箱命令运行。
- 工具必须按 x86 编译和运行，因为需要反射加载 2024 主程序中的 `RecoNet.Security`。
- 2020 消耗串解密后按“前 9 位电算代号 + 后续消耗量”解析，不能按固定小数位切分。
- 代码 `10`、`11` 在接触网概算中作为特殊人工代码保留，不按材料或机械补充。
