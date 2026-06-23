# WPS 点选即填无响应修复设计

## 背景

`ExcelInstantQuantityInputFeature` 已能通过 `Ctrl+Shift+Q` 开启，也能捕获定额表中的“工程数量”单元格。运行日志显示，点击 WPS 数量单元格后没有读取或写入记录。

根因是新功能只使用 `Marshal.GetActiveObject` 获取 WPS/Excel 实例，而现有 `ExcelLinkFeature` 已经验证当 ProgID 方式失败时，需要通过 WPS 窗口对象获取实例。

## 目标行为

- 保留 `Ctrl+Shift+Q` 作为手动开关，不默认自动开启。
- 开启后，先点定额“工程数量”格，再点 WPS 数量单元格，数值立即写入当前定额行。
- 保留已有单位识别、数量换算、数据绑定提交和状态提示。
- 连接不到 WPS 时留下可诊断日志，避免静默无响应。

## 实现设计

`FormPanel` 的各功能文件是同一个 partial class。`ExcelInstantQuantityInputFeature` 将直接复用 `ExcelLinkFeature` 中的 `GetActiveSpreadsheetApplication()`，不再维护一套仅支持 ProgID 的重复入口。该共享方法依次尝试：

1. WPS/Excel ProgID 活动对象。
2. 通过表格窗口的 Native Object Model 获取应用实例。

读取到实例后，现有选区读取和数量写入流程保持不变。若连接失败，按时间节流记录失败原因，不在 250ms 轮询中持续刷日志。

## 异常与边界

- WPS 未打开、无活动工作簿或无选区时，不修改定额数据。
- WPS 当前单元格不是有效数值时，保留现有“已跳过”提示。
- 未选定定额数量列时，不写入任何单元格。
- 关闭 `Ctrl+Shift+Q` 后停止轮询，不改变现有开关语义。

## 验证

1. 构建 `RecoExpandPanel.dll`，确认 partial class 间的共享方法调用编译通过。
2. 启动 2020 投资控制系统和 WPS 表格，按 `Ctrl+Shift+Q` 开启功能。
3. 点定额“工程数量”格，再点 WPS 数量格，确认数值写入。
4. 用单位一致和需换算的样例各验证一次。
5. 关闭开关后再点 WPS，确认不写入。
6. 查看 `RecoExpandPanel.log`，确认 WPS 通过 ProgID 或窗口对象连接，且没有高频失败日志。
