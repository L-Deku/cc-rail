# RecoQuotaRecommend

这是独立的“定额推荐”功能目录，不修改原乘系数 DLL。

当前包含两部分：

- `QuotaLearningImporter.exe`：离线读取“已编制预算表 + 工程量表”，生成学习库。
- `RecoQuotaRecommend.dll`：软件内推荐入口，读取学习库并展示候选定额。

## 生成学习库

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Users\谢刚\Desktop\自动预算\RecoQuotaRecommend\build.ps1"
& "C:\Users\谢刚\Desktop\自动预算\RecoQuotaRecommend\bin\QuotaLearningImporter.exe" "C:\Users\谢刚\Desktop\自动预算\阜阳北机务段新增股道工程" "C:\Users\谢刚\Desktop\自动预算\RecoQuotaData"
```

输出文件：

- `RecoQuotaData/learning.jsonl`
- `RecoQuotaData/learning.csv`
- `RecoQuotaData/learning-summary.txt`

## 软件内使用

`build.ps1` 会同时编译 `RecoQuotaRecommend.dll`，并复制到软件根目录。

主程序通过 `RecoPluginLoader.dll` 加载扩展：

- 继续加载原来的 `RecoExpandPanel.dll`，保留乘系数功能。
- 如果根目录存在 `RecoQuotaRecommend.dll`，则加载“推荐定额”功能。
- 删除 `RecoQuotaRecommend.dll` 和 `RecoQuotaData` 后，推荐功能即移除。

## 当前批量工作流

1. 在工程量 Excel/WPS 表格中框选三列：工程量名称、单位、工程量。
2. 回到软件定额输入表右键点击“推荐定额”。
3. 弹窗会按框选区域逐行识别工程量，并从学习库给出推荐定额。
4. 默认勾选置信度不低于 60 的推荐。
5. 点击“复制勾选内容”，插件会把可粘贴内容放入剪贴板。
6. 回到定额输入表目标位置，从第一列“定额编号”开始人工 `Ctrl+V` 粘贴。

复制格式为四列：

```text
推荐定额编号    空列    空列    工程数量
```

这样工程数量会对齐到定额输入表第 4 列“工程数量输入”。

复制时会按 Excel 单位和推荐定额单位做数量级换算：

- Excel 单位 `m3`，定额单位 `100m3`：工程量除以 100。
- Excel 单位 `m2`，定额单位 `100m2`：工程量除以 100。
- 单位数量级一致时不换算。

注意：当前写入采用人工复制粘贴方式，不直接绕开软件业务逻辑写项目数据库。推荐窗口为非模态窗口，可以和定额输入窗口同时打开。

## 人工扶正

推荐表的工程量名称前有“扶正”按钮。

使用方式：

1. 在定额输入表中选中一条或多条正确的定额。
2. 可先在批量推荐窗口中直接修改“工程量名称”，让名称更贴近实际定额匹配关系。
3. 点击对应工程量行的“扶正”。
4. 插件会把该工程量和所选定额的对应关系写回学习库。
5. 同一工程量可以对应多条定额。
6. 再次推荐时，人工扶正记录优先于自动学习记录。

人工扶正写入学习库时，定额侧只记录：

- `quota_code`
- `quota_name`
- `quota_unit`

## 定额编号归一化

- `LY-16参` 记录为 `LY-16`
- `FY-66参` 记录为 `FY-66`
- `LY-35*4参` 记录为 `LY-35`
- `109001003*1.02` 记录为 `109001003`
