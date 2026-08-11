# AGENTS.md

本文件记录本仓库的项目规则，供后续编码代理和维护者参考。除非用户明确要求，所有改动都应遵守这些约定。

## 项目概况

- 本项目用于给铁路基本建设工程投资控制系统增加批量推荐定额能力。
- 核心插件源码位于 `RecoQuotaRecommend/`。
- 主要源码文件是 `RecoQuotaRecommend/QuotaRecommendPanel.cs`。
- 构建脚本是 `RecoQuotaRecommend/build.ps1`。
- 插件部署目标目录是 `铁路基本建设工程投资控制系统2020网络版V0503021201/`。
- 本地定额与学习数据缓存放在软件目录下的 `RecoQuotaData/`。
- 插件构建、工作区运行目录同步和同事发布包更新时，只允许以当前仓库 `main` 的 `RecoQuotaRecommend/bin/` 为插件 DLL 源头；不要从任何运行目录或发布包反向覆盖当前仓库输出。
- `D:\AI文件\同事模拟目录`（2026-07-20 由 `D:\AI文件\铁路工程云计价系统网络版V1.0` 改名而来）是“同事电脑模拟目录”，必须保持为最近一次实际发给同事的版本；日常构建、工作区部署、发布包刷新和提交均不得自动同步该目录，也不得把它写进 `build.ps1` 的部署目标。
- Claude Code 可能在 `.claude/worktrees/` 下并行修改；除非用户明确要求，不要主动把当前主工作区改动合入 Claude worktree，也不要从 Claude worktree 部署覆盖当前软件目录。

## 构建与验证

- 修改插件后优先运行：

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Users\谢刚\Desktop\自动预算\RecoQuotaRecommend\build.ps1"
```

- 构建成功后，脚本会生成并部署 `RecoQuotaRecommend.dll` 到软件目录。
- 修改 `tools/RecoExpandPanel/` 后，如果因为目标软件正在运行只编译了 `RecoQuotaRecommend/bin/RecoExpandPanel.dll` 而未部署，必须明确说明运行目录仍是旧 DLL；现场验证前先用 DLL 时间戳或 marker 确认运行目录 `RecoExpandPanel.dll` 已更新。
- `tools/DeployUnifiedPlugins.ps1` 当前会识别工作区内外的多个运行目录；日常构建和发布包更新时可先干跑检查，但不得直接加 `-Deploy`：

```powershell
powershell.exe -ExecutionPolicy Bypass -File "D:\AI文件\自动预算\tools\DeployUnifiedPlugins.ps1" -SkipBuild
```

- 只有用户明确要求同步干跑列出的全部目标，并确认目标软件已关闭后，才能显式加 `-Deploy`；脚本会复制四个插件 DLL 并备份旧 DLL，不同步 `RecoQuotaData`、`ExcelLinks` 或项目配置数据：

```powershell
powershell.exe -ExecutionPolicy Bypass -File "D:\AI文件\自动预算\tools\DeployUnifiedPlugins.ps1" -Deploy
```

- 每次构建或部署插件后，如果 `release/同事插件分层发布/` 已存在，必须同步更新对应功能的首次安装包和 `90-后续更新文件`，重新生成 `文件清单-SHA256.txt`，并验证发布包 DLL 与 `RecoQuotaRecommend/bin/` 源 DLL 哈希一致；这一步只更新工作区内的发布包，不得同步“同事电脑模拟目录”。仅更新 `RecoExpandPanel.dll` 时不得重建或复制 `RecoQuotaData`、其他插件 DLL 或同事数据；给已安装插件的同事发送时，只发送 `90-后续更新文件/综合扩展更新/RecoExpandPanel.dll`。
- 每次修改插件并部署后，必须同步刷新发布包 `release/同事插件分层发布`（`pwsh -File tools/BuildColleaguePluginRelease.ps1 -Force`），使发布包始终等于当前最新构建；发布包与同事模拟目录是两回事，刷新发布包时绝不触碰同事模拟目录。
- 核对或备份发布包 DLL 时，先用 `rg --files -uu release/同事插件分层发布` 枚举当前实际路径，不得硬编码或猜测历史分层目录名。
- 当用户明确说准备给同事发布，或询问“同步最新功能需要把发布包的哪些文件发给同事”时，先对比发布包与同事模拟目录的版本/哈希，列出应发文件；只有在用户明确同意后，才把这些文件同步到 `D:\AI文件\同事模拟目录`，并验证它与当次实际发送文件的哈希一致；未经明确同意不得复制 `RecoQuotaData` 或其他同事数据。同事模拟目录只有在“用户真的把文件发给同事”这一刻才更新，这样它才真实反映同事手里的版本。
- 运行构建部署前先确认目标软件主程序已关闭；`RejjNet2020.exe`、`ReJJGSNet2024.exe` 等进程占用插件 DLL 时，`Copy-Item` 会因文件被占用而部署失败。
- 构建用于部署或发布的 `RecoExpandPanel.dll` 前，如果 `tools/RecoExpandPanel/` 还存在本任务范围外的未提交修改，必须从当前已确认提交生成干净源码快照后编译；不得把并行工作或用户未确认的工作树改动混入发布 DLL，也不得还原这些未提交修改。
- 验证时优先用插件内部检索入口或实际软件界面验证，不只看文本文件搜索结果。
- 启动 `RejjNet2020.exe`、`ReJJGSNet2024.exe` 做插件加载冒烟验证后，先轮询进程是否已自行退出；需要关闭时优先用窗口关闭/`CloseMainWindow`，不要默认 `Stop-Process` 或 `taskkill /F` 一定有权限结束软件进程。
- 仅修改 UI 写入路径、窗口交互或软件原生写入调用时，优先做实际软件点击/日志验证；离线检索样例回归主要用于检索规则、索引、匹配权重或组件框逻辑变更。
- 不要要求用户手动复制 DLL；用户和代理在同一台机器上，构建脚本会处理部署。
- 迁移 2020 概算/估算到 2024 定额库时，辅助查询窗口会读取 `定额库.消耗` 的 2024 运行态加密串；只写 `定额库消耗` 明细表或使用旧版加密串会导致“显示定额消耗错误”。
- 生成 2024 `定额库.消耗` 时应通过已加载 2024 主程序中的 `RecoNet.Security.Encrypto` 运行态入口完成；不要把其他定额的密文临时替换结果当作最终验证。
- 2020 原定额消耗明文按前 9 位解析旧电算代号；写入 2024 时必须把替换后的电算代号格式化为固定 10 位，再拼接原消耗量。使用 9 位会导致数量首位被界面读入电算代号。
- 2024 云计价启动报 `Ahs Error: 0x3000102` 或所有 2024 目录都打不开而 2020/2022 可打开时，先检查系统 `AHS Service` 是否正常、服务注册路径 `D:\ahsProtector\ahs_service.exe` 对应目录是否存在；插件排查应先临时移走 `ReJJGSNet2024.exe.config`、`RecoPluginLoader.dll` 等加载入口，确认无插件残留后再判断授权服务问题。已验证在恢复 `D:\ahsProtector` 后，2024 可正常加载 `RecoPluginLoader.dll`、`RecoExpandPanel.dll` 和 `RecoQuotaRecommend.dll`。
- 参考定额池同时承载 2020/30号文、101号文估算和 2024 新办法记录；加载 `chapter-quota-library.jsonl` 时不能只按运行目录的 `method` 过滤，否则 2024 客户端中打开 30号文项目会把 `method=2020, method_no=30号文` 的旧参考池全部过滤掉。池隔离应使用 `method_no|entry_code`。
- `missing-resources.xlsx` 的人工审核规则每次重建都必须强制应用，不能因补充资源已存在而跳过：L=`0` 使用 C 列 `new_supplement_code`，L=`1` 使用 G 列 `best_candidate_code`，L 为其他数字时直接使用 L 列电算代号。
- 排查迁移定额消耗时，先用迁移工具的 `InspectQuotaRows --book <书号> --code <定额编号>` 查明细行数、重复代号和缺失资源，再结合软件界面验证。
- 清理已迁移的 2020 概算/估算定额时，优先运行 `tools/Migrate2020EstimateTo2024/CleanupMigratedEstimate2024.ps1 -WhatIfOnly` 生成只读报告，再正式执行脚本；补充材料/机械必须按 `applied-manual-decisions.csv` 的真实写入代码删除，并加 `not exists(select 1 from 定额库消耗 where 电算代号=@code)` 保护，不能直接使用旧 `rollback.sql`。
- Windows PowerShell 5 执行含中文表名/字段名、中文路径或中文测试文本的入口脚本和回归脚本时，脚本文件必须保存为带 BOM 的 UTF-8，否则可能把中文解析成乱码并导致语法错误；补 BOM 时只能改变编码标记，正文与换行必须保持不变。`pwsh` 的 PASS 可作辅助证据，但不能替代要求 Windows PowerShell 5 的验收。
- PowerShell 反射验证脚本如果包含中文路径或中文测试文本，优先在当前 shell 直接执行，或保存为 UTF-8 脚本文件后用 `-File` 执行；不要把 here-string 管道给子 `powershell.exe -`，否则中文路径可能被管道编码破坏。
- 反射加载 `RecoExpandPanel.dll` 并实际调用 `JavaScriptSerializer`/`System.Web.Extensions` 解析 JSONL 时使用 Windows PowerShell 5；`pwsh` 可能因 .NET Framework `System.Web` 类型不兼容把解析异常吞成空结果。纯源码检查或不触发该依赖的反射用例仍可使用 `pwsh`。
- 在 PowerShell 中再调用 `powershell.exe -Command` 且子命令包含 `$变量` 时，优先把子命令保存为脚本后用 `-File`，或用单引号整体隔离子命令；不要让父级 PowerShell 提前展开子命令变量。
- PowerShell 直接调用 `csc.exe` 时，先把输出和引用文件路径赋给变量，再传 `/out:$变量`、`/reference:$变量`；不要写 `/out:(Join-Path ...)` 或 `/reference:(Join-Path ...)`，否则 `csc` 会收到空路径参数。
- Windows PowerShell 5 中不得用 `Group-Object -Property <键名>` 对哈希表键值去重；其属性适配可能把不同键值并入空组。重建目标集合时应显式用索引取值和 `HashSet` 去重，并在 Windows PowerShell 5 下做多目标行为测试。
- SQL Server 结构迁移需要动态删除约束时，不要在 `EXEC(...)` 参数中直接拼接 `REPLACE`/`QUOTENAME` 等函数；先组装到 `NVARCHAR(MAX)` 变量，再用 `sys.sp_executesql` 执行，并重跑幂等 schema 脚本验证。
- Windows PowerShell 5 中用 `[IO.File]::Replace` 原子更新已有状态文件时，不要把备份路径传 `$null`；应使用同目录唯一临时旧版路径，提交后清理，并用“首次写入 + 已有文件替换 + 无临时残留”运行态测试验证。备份已成功而状态落盘失败的恢复流程必须重新验证服务器侧备份且禁止覆盖已有备份。
- `tools/RecoExpandPanel/tests/Test-TemplateFillNameMatch.ps1` 默认加载 `RecoQuotaRecommend/bin/RecoExpandPanel.dll`，不会自动编译当前源码；做源码级红绿回归时，应先把 `tools/RecoExpandPanel/` 当前全部 C# 源文件编译到工作区验证目录并设置 `RECO_EXPAND_DLL`，避免把旧 DLL 的结果误判为新代码结果。
- `build.ps1 -BuildOnly` 输出目录按白名单只含两个插件 DLL 和清单，不携带 NPOI 运行依赖；反射测试新 `RecoExpandPanel.dll` 前应在工作区 `artifacts/test-runtime/` 创建隔离测试目录，仅复制该 DLL 与既有 NPOI 依赖。测试依赖不得复制到运行目录或发布包。
- BuildOnly 清单中的源码哈希必须来自实际传给编译器的只读快照，不得在编译后重新哈希可能已被并行修改的工作树原文件；`source_commit` 只表示基线 HEAD，dirty 工作树必须另行显式标记。
- `tools/RecoExpandPanel/tests/Test-TemplateFillNameMatch.ps1` 在非交互 WinForms 环境可能卡在“定额候选下拉与组件组界面确认”之后的滚动视口用例；连续停在该位置时应按“综合回归未完成”报告，终止并核对本次测试启动的精确进程，不得把前半段 PASS 当作全部通过，也不要反复无上限重跑。
- 读取 UTF-8 附件或中文文本时，如 PowerShell `Get-Content` 输出乱码，先设置 `[Console]::OutputEncoding`，并优先用 `[System.Text.Encoding]::UTF8.GetString([System.IO.File]::ReadAllBytes(...))` 按字节解码验证内容。
- PowerShell 部署/验证脚本需要格式化多段 `foreach` 输出时，优先先收集到数组或 `List[object]` 再统一 `Format-Table`；不要把脚本块闭合后直接接管道，容易触发 `An empty pipe element is not allowed` 解析错误。
- PowerShell 创建目录时不要给 `New-Item` 使用不存在的 `-LiteralPath` 参数；中文或特殊字符路径优先调用 `[System.IO.Directory]::CreateDirectory($path)`，并在后续读写继续使用 `-LiteralPath`。
- Windows PowerShell 中用 `rg` 搜索指定目录文件时，优先写 `rg -n "pattern" -S path` 或用 `rg --files -g "*.cs"` 先列文件；不要把 `目录\*.cs` 当作路径参数传给 `rg`，容易被解析成非法路径。
- PowerShell 中用 `rg` 同时搜索含中文引号、括号或反斜杠的多个模式时，优先拆成多个简单的 `rg -n -F "literal" path`；不要在一条双引号命令里拼复杂分组正则，避免 PowerShell 截断引号后产生伪语法错误。
- GitHub 推送报 `Failed to connect ... via 127.0.0.1` 时，先分别检查 Git `http.proxy`/`https.proxy` 与 `HTTP_PROXY`/`HTTPS_PROXY`/`ALL_PROXY` 环境变量；若本地代理端口失效或不一致，先在单次命令中临时清空环境代理并用 `git -c http.proxy= -c https.proxy=` 做 `ls-remote` 直连验证；不要未经用户确认修改全局代理配置。
- Git 暂存范围包含中文或其他非 ASCII 路径时，使用 `git -c core.quotepath=false diff --cached --name-only` 获取可比较的真实路径；不要直接比较默认的八进制转义输出。
- PowerShell 接收原生命令输出后需要使用 `.Count` 或 `[0]` 校验时，应写成 `[string[]]$items = @(...)`；单行输出若保留为普通字符串，`[0]` 只会得到首字符，可能造成错误的范围校验失败。
- Windows PowerShell 5 中通过 `if/else` 表达式返回集合并赋值时，单元素结果也会被自动拆成标量；后续依赖 `.Count` 或乘法计数的变量应先声明为 `[object[]]`，再在分支内用 `@(...)` 赋值。
- Windows PowerShell 5 把 `System.Collections.Generic.List[object]` 放入哈希表/JSON 对象时，不要用 `@($list)` 包装，可能报 `Argument types do not match`；应显式调用 `$list.ToArray()`，并在 PS5 下运行真实序列化测试。
- Windows 上用 `git archive` + `tar` 生成干净构建快照时，如果仓库包含大量中文文件名，应把归档范围限制为实际参与构建的源码子树（如 `tools/RecoExpandPanel`），并核对源码文件数量；不要无条件归档整个仓库，避免 `tar` 因中文路径解码失败。
- 修改已含中文字符串的 C# 源码时，避免用 PowerShell `Set-Content` 默认编码整文件重写；优先用补丁方式，必要时用 `.NET UTF8Encoding(false)` 并把新增中文字符串写成 `\u` 转义，防止产生无关编码差异。
- 新增或修复绑定学习字段时，必须按“数据源 -> 预览对象 -> `ExcelQuotaLink` XML -> `mapping-boxes.jsonl` -> `BindingLog`/聚合表 -> 推荐读取”逐段核对；Excel 工程量单位与定额目标单位要分别做回归，不能因预览对象已有字段就认定持久化链已传递。
- 推荐定额/模板铺量的名字驱动组件“确认写入”即为接受推荐，必须回流 `source='plugin:apply-accept'`；只有整个 `TargetRow` 组件组全部写入成功才学习，残缺组不得产生 accepted，学习库写入失败不得阻断实际写入结果。
- `SignatureBoxMap.weight` 不设上限，只保留下限 0；调整公式时必须同步修改 SQL 增量写入、本机 `mapping-boxes.jsonl` 和 `Rebuild-Aggregates.ps1` 三端，并执行一次全量重算使历史聚合收敛。
- `Rebuild-Aggregates.ps1` 分配 `QuotaBox.box_id` 时，历史显式 `box_id` 可能以 `auto-` 开头并与其他目标集合的自动 MD5 前缀冲突。必须先确定性保留唯一的历史显式 ID，再延长自动哈希前缀直到唯一；不得合并不同 `target_set_hash` 或依赖遍历顺序。
- 修改 `NormalizeForSignature` 或 `Get-NormalizedPart` 时必须同步另一端并执行一次 `Rebuild-Aggregates.ps1`；包括 `-DryRun` 在内都会持有 `BindingLog` 独占锁，只能在冻结绑定写入的维护窗口执行。
- 推荐快照 `SmartLearningSnapshot` 视为只读；范围过滤和排序只能操作副本，不得就地修改快照中的集合。
- `SH`、`SQ`、`ZLF`、`LF`、`SF`、`TLF`、`YF`、`GF`、`JF`、`XGT1` 等通用辅助代码不得只按编号聚合或解析；本地组件、SQL 增量/全量聚合和推荐读取必须统一使用“类型+完整编号+规范化名称+规范化单位”身份。同一绑定事件内同码多义时只保留原始 `BindingLog`、不得晋升聚合；当前项目名称或单位不一致时整组过滤，纯辅助组件不进入普通推荐，`SF` 设备购置费例外。
- `SF` 只能写入名称含“设备购置费”的条目，名称含“设备购置费”的条目也只接受 `SF`；任一方向违反时必须阻断整组，禁止自动写入和手工确认，不产生 accepted。条目名判定顺序固定为“当前项目真实条目名 > 精确 `(LibraryMethod, MethodNo)` 分区的 `ChapterEntry` > 目标级历史名称”，不得跨办法回退。
- 组件框中的条目证据必须保存到每个目标；普通定额和 `SF` 可各自形成 `EngineeringTemplate` 专业范围，纯 `SF` 框也必须归入设备购置费范围。`ZLF`、`LF` 和材料可跟随普通主定额条目，但不得独立扩展专业范围；`SH` 保留自身目标级条目证据，但不参与工程前缀投票。推荐学习库“未归类”只由持久化 `EngineeringTemplate` 是否存在决定，本机临时上下文不得把框移出“未归类”。
- 绑定学习的聚合签名只使用归一化工程量名称（兼容 `名称|`），Excel 工程量单位仅作流水审计；推荐数量必须用当前 Excel 单位和当前运行版本定额单位现场换算，不得学习或复用原绑定表达式中的 `/1000`、`*1.05` 等运算。多单元格正向加项只拆工程量别名，各别名指向原表达式的完整组件框；组件内目标共用同一套编制办法，但稳定条目必须按目标分别保存，同一原始表达式不得因目标条目不同而拆散组件。
- 跨量纲业务换算不得简化成历史单元格地址或一次性结果；应保存 `V0/V1...` 参数公式及每个参数的名称、单位和名称级签名。推荐时只用当前表同章节、邻近行内精确且唯一的参数重新计算；缺参数、同名歧义或单位不兼容时必须取消自动勾选。`F10/1000+F11/1000` 仍按独立正向别名学习，不得因此生成共享公式或跨行合计。
- C# 5 代码中不要在 `||`/`&&` 短路条件里依赖 `out` 参数一定赋值；用于错误文案的 `out` 变量先给默认值，避免 `CS0165`。
- 插件访问当前项目数据库时，不得从宿主已脱敏的 `ConnectionString` 克隆新连接，不得恢复备用账号、保存密码或依赖 `Persist Security Info`。UI 线程路径只借用宿主当前 `SqlConnection`，不得 `using`、`Dispose`、`Close` 或 `ChangeDatabase`；后台任务只把短数据库阶段同步调度到 UI 线程，网络请求继续留在后台。预览、执行、撤销和重做必须同时保存并核对连接对象引用与不含密码的 `DataSource|Database` 身份，项目切换后拒绝旧计划。
- 用 Windows PowerShell 5 反射调用 WinForms 私有构造器做冒烟测试时，先设 `$ErrorActionPreference = 'Stop'`，并把 `New-Object` 返回控件的 `.PSObject.BaseObject` 传给反射 API；泛型 `List<T>`、`HashSet<T>` 等参数也要拆包，否则类型包装错误可能只产生非终止错误并让命令假通过。反射方法只接收一个泛型集合参数时，不要直接用 `[object[]]@($list)`，应先创建长度为 1 的 `object[]` 再将 `.PSObject.BaseObject` 赋给第 0 项，避免 PowerShell 把集合展开成多个参数。反射读取 `List<T>` 后若经辅助函数返回并继续使用 `.Count`/索引，辅助函数应使用 `Write-Output -NoEnumerate`，避免单元素集合被自动展开成标量。
- 2024 软件预算项目输入 2020 概算/估算定额时，首要检查 `项目设置 -> 定额选择` 是否勾选了迁移书号，或数据库 `项目信息.标准定额应用` 是否包含对应书号。未勾选时会出现“定额编号无效或费用类型不匹配”、计算单价为 0 或“无法找到定额消耗数据”等现象；勾选后 2024 原生辅助查询、输入和计算即可使用迁移定额。
- 2020 概算/估算定额迁移到 2024 后，正式方案不需要给 2024 客户端部署兼容插件或 Harmony/Prefix 钩子。不得把钩子部署到 `RecoNet.DEBase.FindDe` 作为正式方案；实测 `FindDe_Patch1` 会影响原预算查询/输入并触发运行时异常。每次发布前必须回归验证 `LY_2024`、`DY_2024` 等原预算定额的查询、输入和计算。
- 给新下载的 2024 客户端使用已迁移的概算/估算定额时，只要连接同一个已迁移的 `RecoData2024` 数据库，并在项目设置中勾选对应书号即可；插件 DLL 和 `RecoPluginLoader.AutoLoadDomainManager` 只属于推荐定额插件能力，不是概算/估算定额原生查询、输入、计算的必要条件。
- 101号文估算参考定额池的数据源是 `192.168.2.13` 的 `RecoData2020`，不要误用 `192.168.2.213`；模板条目来自 `RecoData2020.dbo.章节表 where 编制办法文号='101号文估算'`，历史项目中的细分条目可能按模板上溯归并，例如 `0616-02-01-01-04-01` 会归到 `0616-02-01-01`。
- 参考定额池只应保留可作为定额输入的原定额编号：普通条目必须是 `quota-index.jsonl`/全套定额库中存在的 `target_kind=quota` 原编号，扫描和手工添加都要去掉末尾 `/数字` 换算子目后缀；材料码和 `SH`、`SQ`、`ZLF`、`LF`、`YF`、`TLF`、`GF`、`JF`、`XGT1` 等辅助/伪代码不得进入参考定额池。唯一例外是 `entry_name` 包含 `设备购置费` 的条目，只允许保留/添加 `quota_code=SF`，用于软件原生设备费输入。
- 参考定额池界面的“当前条目”定位不能把“候选池为空”当成“无效条目”：只要项目界面能读到条目编号，就应显示当前条目并允许用户新增定额；已有候选池或可上溯命中父级池时仍按既有池显示。
- 参考定额池中的定额编号必须使用软件可原生输入的原定额编号；扫描历史项目或读取用户输入时，要去掉末尾 `/数字` 换算子目后缀（如 `PY-385/1` 应归一为 `PY-385`），避免参考框双击输入失败。

## 数据存储规则

- 推荐关系、绑定组件和换算公式只允许存入并从 `RecoLearning` SQL 学习库读取；不得恢复 `mapping-boxes.jsonl`、`learning.jsonl`、pending overlay、outbox、dead-letter 或其他本地学习/配对回退。
- `quota-index*.jsonl` 与 `material-index*.jsonl` 仅是定额/材料名称、单位等索引，不是学习关系库；可以保留用于元数据补齐或普通候选检索，但不得据此恢复本地绑定配对。
- 代理不得直接执行 SQL 修改铁路投资控制系统业务数据库。用户在软件界面主动点击插件写入功能并完成确认后，插件可以通过安全借用宿主当前项目连接或软件原生写入机制修改当前项目；代理不得代点。
- 已完成的 D0–D4、原数据库迁移、既有 DLL 部署和双软件分区数据生成均视为冻结证据；后续代码修复不得顺带重跑迁移、重建聚合或重新生成既有迁移数据，除非用户针对新的维护窗口再次明确授权。

## 检索规则

- 推荐组件和可信换算公式先按当前 `software_partition`、办法、条目及目标身份查询 `RecoLearning`；SQL 未命中或不可用时不得回退本地学习文件。
- `quota-index`/`material-index` 可用于目标编号已确定后的名称、单位补齐及普通检索，但不能生成或替代 SQL 学习组件关系。
- 没有可靠 SQL 关系或候选时返回空定额，等待人工扶正或既有 SQL 接受回流。
- 普通关键词检索只返回 1 条最优定额。
- 普通关键词检索不能返回材料；材料只能通过人工扶正后的组件框出现。
- 推荐展示以定额编号为主，同时展示定额名称、单位和换算后数量。
- 定额名称命中权重最高，工作内容次之，节名称和专业分类只能作为辅助证据。
- 单字匹配只能作为低权重辅助分，不能单独让定额通过阈值。
- 连续词优先级应高于离散单字：完整词 > 连续三字 > 连续两字 > 单字辅助。
- 对钢筋类工程量要避免被“钢筋混凝土构件”泛匹配抢走；`HPB/光圆/圆钢` 优先匹配圆钢筋类定额，`HRB/螺纹` 优先匹配螺纹钢筋或明确 HRB 的钢筋制作绑扎类定额。

## 组件框与反馈规则

- 一个工程量可以对应一条定额，也可以对应一个组件框中的多条定额/材料。
- 例如：
  - `土方外运 -> LY-21 + LY-34 + LY-35`
  - `混凝土检查井 -> ZY-41 + 商品混凝土材料`
- 组件框命中时，界面只在第一行显示工程量名称和扶正按钮。
- 同一组件框的后续行工程量名称留空，扶正按钮留空且不可点击。
- 用户点击“复制勾选内容”视为接受推荐，应提高对应样本权重。
- 用户扶正时，应降低原错误推荐权重，并把当前工程量名称加入目标定额或组件框。
- 每个框保存的工程量名称样本容量默认最多 30 个。
- 容量满时优先淘汰权重低且长期未使用的样本。

## UI 行为规则

- 不可靠推荐不要硬塞；允许定额编号、名称、单位为空，等待用户扶正。
- 空推荐行仍应显示工程量名称、单位、数量和扶正按钮。
- 组件框生成多行推荐时，只第一行承载工程量名称和扶正入口。
- 保持原有 Excel 读取、剪贴板读取、数量换算、复制粘贴格式和人工扶正按钮可用。
- Excel 点选即时填进入连续复用模式后，成功写入不得关闭当前定额目标的 Excel 选区监听；应以工作簿、工作表、地址、值和单位组成的基线判断是否变化，同一基线去重，新单元格则覆盖当前单选或多选目标。只有离开数量列、关闭功能、数据源切换或目标失效时才清理监听状态。
- Excel 点选即时填的 COM 热路径不得重复读取当前选区：首次监听只读取一次防误填基线；已有可复用选区时保留现有基线，由下一次轮询读取一次并写入新目标。进入监听所需状态应统一通过一个方法设置。
- Excel 点选即时填的快捷键应同时挂到主窗口 `KeyPreview` 和定额输入表，开启提示使用普通 `ToolTip`，不要使用可能抢焦点或干扰 WPS/Excel 点击识别的气泡提示。
- 绑定 Excel 工程量的表达式模式回归时，必须验证窗口保持打开后继续点击下一 Excel 单元格，表达式输入框会同步替换上一次自动带入的单元格，同时保留 `/100`、`*1.5` 等人工运算后缀。
- WinForms `SplitContainer` 新增分栏窗口时，不要在控件尚未布局出有效宽度前设置较大的 `Panel1MinSize`、`Panel2MinSize` 或 `SplitterDistance`；应在 `Shown`/布局后按当前宽度夹紧设置，并做反射构造冒烟验证。
- 自动匹配 Excel 和模板铺量读取工程名称时，合并单元格回填应使用同一套规则：优先流式读取工作表 XML 的 `mergeCells`/`mergeCell` 合并区域并按 `工作簿|工作表` 缓存，不要在 UI 线程为枚举合并区域整簿 `WorkbookFactory.Create`，也不要退回“向上回看最近文字”的启发式串名。
- Excel 快照中“值”来自实时 COM 而合并区域来自磁盘文件时，必须提示或假定匹配前已保存 Excel；若未保存导致行列错位，应优先要求保存后重试，不要把错位回填当作匹配规则问题。
- 合并单元格名称片段去重后，横向合并跨多列的片段应按稳定列归属排序，例如区域内离数量列最近的列或明确锚点列，不能依赖 `Dictionary` 枚举顺序决定取舍；工程量名称拼接应尽量满足至少 3 个片段且至少 15 字，自动匹配 Excel 和模板铺量必须一致。
- 自动匹配 Excel 预览中，数量为 0 需手动匹配的定额和多处匹配的定额应参照模板铺量用浅红底提示；行被勾选、手动匹配或从多处候选中选定工程量名后，应立即取消标红。
- 模板铺量读取目标数量单位时，不得只检查数量列紧邻左格；应在同一 Excel 快照中向左检查最多 6 个可见列，跳过隐藏列、空白和数值，避免 `D=单位、E=数值、F=数量` 时丢失 `/100` 等单位换算。
- 模板铺量同名候选的勾选确认、下拉切换和右键临时绑定应按 `TargetRow` 局部更新当前工程量组并保持滚动视口；不得调用 `FillGrid()` 清空重建整表，避免列表跳回顶部并丢失其他行尚未同步的勾选状态。
- 模板铺量名字驱动中，“同名多来源”或“重复工程量名称”的待确认文案属于非阻断 `AlignNote`，红色由 `NeedExactNameConfirmation` 表示；不得把该提示写入阻断性 `Status`。默认候选正确时必须允许用户直接勾选组首复选框确认，确认后整组勾选并取消红色；只有真实的单位、条目、公式或取数错误才能阻止确认。
- “推荐定额”虽然复用 `TemplateFillPanel` 的预览表格，但匹配分层、标红原因和预览汇总来自 `SmartFillFeature.BuildPreview_SmartFill`；诊断“推荐定额”截图时必须按状态列、候选提示以及 SQL 聚合读取链解释，不得套用模板铺量的同名来源规则。
- 推荐定额组件候选下拉只显示组件编号集合及“条目编号 + 当前项目条目名称”，学习权重和办法证据仅用于内部排序/自动采纳；同显示文案候选保留排序第一项。数量列人工修改必须同步到 `QuantityText`，直接勾选当前候选应原地确认以保留修改值，且确认后整组逐行刷新 `AlignNote`。
- 排查模板铺量匹配异常时，必须按界面当前选择的模板名调用 `LoadFillTemplate` 并核对实际模板 JSON；不得用名称相近的其他模板代替复现后推断匹配分支。
- 能用软件原生写入的尽量用原生写入，优先复用软件已有的定额查询、选择、写入和计算逻辑。
- 触发定额输入表原生补齐时，应在当前行“定额编号”单元格进入编辑后模拟键盘输入/粘贴编号并回车；直接设置单元格值或编辑控件 Text 可能不会触发软件填充单重、编制人、修改日期等派生字段。

## 代码编辑规则

- 优先保持现有 WinForms/C# 风格，不引入大型新框架。
- 手工改文件使用补丁方式，避免无关格式化。
- 不要提交或展示数据库密码、连接串等敏感信息。
- 排查凭据或连接串相关源码时，不要直接用 `rg` 输出匹配行；先只定位文件/行号，再对必要上下文中的密码、账号、服务器地址和连接串先脱敏后输出。
- 不要删除用户已有数据文件，尤其是学习库、索引库、日志和业务目录。
- 遇到用户或软件生成的未跟踪/已修改文件，不要擅自还原。
- 修改检索规则后必须至少验证以下代表样例：
  - `警示带` 应命中 `PY-738 警示（示踪）带铺设`
  - `警示桩` 不应误命中 `ZY-41 现浇检查坑、井 混凝土`
  - `HPB300钢筋` 不应返回材料，应返回较合适的钢筋制作类定额或留空
  - `HRB400钢筋` 普通检索只返回 1 条最优定额
  - `土方外运` 通过组件框时可返回多条组件推荐
- 同一运行目录同时包含 `RejjNet2020.exe` 和 `ReJJGSNet2024.exe`/`ReJJQDNet2024.exe` 时，定额/材料源数据库必须优先按当前实际进程选择；不得因目录中存在 2024 主程序就把 2020 进程误判为 `RecoData2024`。
- 定额/材料候选变更回归除了验证“查得到”，还必须核对定额书号、材料文号与当前运行版本一致，并在原生编号输入后验证名称、单位和价格能正常补齐；要特别覆盖 `LY-192` 这类在 2020/2024 编号相同但含义不同的定额，不得只以候选框显示成功作为通过。

## 沟通规则

- 向用户说明问题时，优先解释当前规则为什么得出该结果。
- 对定额匹配错误，要区分是索引缺失、关键词分词问题、权重问题、单位问题，还是人工扶正库影响。
- 用户询问是否改规则时，先分析利弊和风险，再等待用户决定。
