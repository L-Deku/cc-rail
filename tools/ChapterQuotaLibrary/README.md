# ChapterQuotaLibrary（推荐定额按章节条目分类）

把推荐定额学习库按编制办法的章节条目分类：只读扫描服务器上已完成项目，为每个条目建立"原始定额池"，并产出条目模板 Excel 供人工删减。

- 2020 版条目模板：国铁科法[2017]30号（库内 `编制办法文号='30号文'`）
- 2024 版条目模板：TB 10801—2024（注意是全角长破折号 U+2014）

## 构建

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Users\谢刚\Desktop\自动预算\tools\ChapterQuotaLibrary\build.ps1"
```

## 使用顺序

```powershell
$exe = "C:\Users\谢刚\Desktop\自动预算\tools\ChapterQuotaLibrary\bin\ChapterQuotaLibrary.exe"

# 1. 建库（饱和度驱动扫描，断点续扫，全程只读 SELECT）
& $exe BuildLibrary

# 2. 导出条目模板 Excel（带每条目学到的定额数），交给用户删减
& $exe ExportTemplate

# 3. 用户删完后导入（整行存在即保留；不要改"条目编号"列）
& $exe ImportTrimmed --in "<删减后的xlsx>"

# 4. 给已有对应框打条目标签
& $exe TagMappingBoxes
```

## BuildLibrary 扫描策略

1. **种子优先**：先完整读取种子项目，其定额无条件全部入池（不受饱和上限丢弃）：
   - 2020：`Reco20260511134731660`（新建南阳经信阳至合肥高速铁路（河南段）勘察设计）
   - 2024：`Reco20250506093156577`（上海经乍浦至杭州铁路初步设计）
2. **补充扫描**：其余项目按创建时间倒序读取；条目池达到 `--max-pool`（默认 50）即饱和不再收新码。
3. **停止**：连续 `--stale-stop`（默认 30）个项目没有给任何未饱和条目贡献新定额即停止。
4. `--limit N` 只扫 N 个项目（冒烟用）；重跑自动跳过 `reports/scan-state-*.csv` 里已完成的库。

## 数据清洗规则

- `定额编号` 为空或 `-`（分组占位行）的行排除。
- 归一化只存原始定额：截断 `*`/`×` 系数后缀，去掉 `参/换/借` 调整字。
- 全数字代号视为材料（material），其余为定额（quota）。
- 项目自定义条目编号按前缀逐级上溯归到最近的模板条目。
- 同一项目存了多个数据库副本时，按建设项目名称去重计 `project_count`。
- 定额名称/单位优先取 `定额库`/`材料单价库`/`台班定额库` 的原始名称；查不到的记入 `reports/unresolved-codes-*.csv`。

## 输出

| 文件 | 说明 |
| --- | --- |
| `RecoQuotaData\chapter-quota-library.jsonl` | 条目定额池（`method` 2020/2024 + `entry_code` + 原始定额；`source=seed/scan/user`） |
| `RecoQuotaData\chapter-entries.jsonl` | ImportTrimmed 后的保留条目树；**此文件存在后运行时条目过滤才生效** |
| `reports\章节条目模板.xlsx` | 交付用户删减的条目模板 |
| `reports\scan-state-*.csv` | 断点续扫状态（库名/状态/命中/新增） |
| `reports\raw-entry-quotas-*.jsonl` | 原始命中明细（重放可重建池） |
| `reports\entry-summary-*.csv` / `build-summary.txt` | 条目饱和度统计 |

## 注意

- 连接 `192.168.2.13,1433`（reco 账户），与主程序同源；本工具只发 SELECT，绝不写服务器。
- `TB 10801—2024QD`（清单变体）项目不在扫描范围。
- 改 `--max-pool` 后建议删除 `reports/raw-entry-quotas-*.jsonl` 与 `scan-state-*.csv` 重扫，否则续扫口径不一致。
