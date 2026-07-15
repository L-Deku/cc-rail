# 绑定Excel直达自动匹配设计

## 1. 目标

定额输入表右键菜单点击“绑定Excel工程量”后，不再显示原 `ExcelSmartBindPanel` 手动绑定面板，直接显示现有 `AutoMatchDialog` 自动匹配窗口。

删除原手动绑定面板及只为该面板服务的缓存、清理、入口和界面代码，同时取消原来打开该面板的 `Ctrl+E` 快捷键。

## 2. 保留范围

- 保留 `AutoMatchDialog` 的自动匹配完整流程，包括 Excel 快照、工作表和目标列选择、开始匹配、预览、单个绑定、全部绑定、保存和取消。
- 保留 `AutoMatchDialog` 内部的“手动匹配”按钮、轮询、目标行选择、Excel 单元格选择、表达式生成和单个/全部绑定流程。
- 保留 `Ctrl+Shift+E` 打开的 `QuickBindPanel` 快速绑定窗口。
- 保留“打开Excel联动面板”、Excel 联动运行时、绑定存储、同步、单位换算、工程量名称写入对应框和模板铺量依赖。
- 保留 `PromptExcelCell` 等被其它现有绑定流程使用的代码；不把“删除原手动绑定面板”扩大为删除所有手动绑定能力。

## 3. 修改设计

### 3.1 右键入口

保留右键菜单文字“绑定Excel工程量”和现有图标，只把点击处理从 `SmartBindSelectedQuotasToExcel` 改为新的直接打开自动匹配窗口入口。

直接入口负责：

1. 获取当前项目数据库连接；连接不存在时显示原有语义的提示并停止。
2. 创建 `AutoMatchDialog`，挂接 `Accepted` 保存回调。
3. 以当前主窗体为所有者显示自动匹配窗口。

### 3.2 自动匹配结果保存

把原来位于 `ExcelSmartBindPanel` 内的自动匹配结果保存逻辑移到 `FormPanel` 的静态入口层，继续复用现有 `LoadStore`、`Upsert`、`SaveStore`、`EnsureExcelLinkRuntime`、`Reload` 和 `RefreshExcelLinkPanel`。

不修改 `AiMatchPreviewItem`、`AutoMatchDialog`、自动匹配算法、表达式、匹配状态或绑定文件格式。

### 3.3 删除旧面板

删除以下只属于旧面板的代码：

- `ExcelSmartBindPanels` 缓存字典及主窗口关闭时的清理分支。
- `SmartBindSelectedQuotasToExcel` 和 `ShowExcelSmartBindPanel`。
- 完整的 `ExcelSmartBindPanel` 类。
- 定额表 `Ctrl+E` 快捷键处理。

不删除 `QuickBindPanel`、`ExcelLinkPanel`、`AutoMatchDialog` 或其它仍有调用者的手动选择/绑定辅助方法。

## 4. 交互流程

修改后的主流程为：

`定额输入表右键 -> 绑定Excel工程量 -> AutoMatchDialog -> 自动匹配或窗口内手动匹配 -> 单个绑定/全部绑定 -> 保存并刷新Excel联动运行时`

`Ctrl+E` 不再触发任何绑定窗口；`Ctrl+Shift+E` 继续打开快速绑定窗口。

## 5. 错误处理

- 当前项目数据库连接不存在时，在打开自动匹配窗口前提示并停止。
- 自动匹配结果保存异常时记录日志并向用户显示错误，不改变现有绑定数据。
- 自动匹配窗口内部的 Excel 连接、快照、单位不匹配和表达式校验继续沿用现有处理。

## 6. 验证标准

1. 源码中不再存在 `ExcelSmartBindPanel`、`ExcelSmartBindPanels`、`SmartBindSelectedQuotasToExcel` 和 `ShowExcelSmartBindPanel`。
2. 定额输入表右键“绑定Excel工程量”直接打开 `AutoMatchDialog`，不经过旧面板。
3. `Ctrl+E` 处理已删除，`Ctrl+Shift+E` 仍调用 `ShowQuickBindPanel`。
4. `AutoMatchDialog` 的“手动匹配”、单个绑定、全部绑定和 `Accepted` 保存链路仍存在。
5. 项目构建通过；对相关入口和保留流程做源码检查或反射冒烟验证。
6. 不修改自动匹配算法、Excel 快照、单位换算、绑定存储结构、Excel 联动面板和模板铺量流程。

