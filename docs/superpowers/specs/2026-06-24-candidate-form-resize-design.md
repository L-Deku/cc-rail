# 候选定额 Form 弹窗设计

## 目标

- 彻底避开 `ToolStripDropDown` 拉伸后内容不同步的问题。
- 候选框改为普通 WinForms 小窗口，像推荐定额和联动 Excel 窗口一样依赖 Form 原生缩放。
- 表格 `Dock=Fill`，窗口尺寸变化时内容天然跟随变化。
- 点击其他窗口或主表格移动时隐藏候选窗口。
- 用户调整后的窗口尺寸只在本次软件运行期间保留；软件重启后恢复默认。

## 方案

- 新增轻量 `CandidatePopupForm : Form`，`FormBorderStyle = SizableToolWindow`，`ShowInTaskbar = false`。
- 候选 `DataGridView` 直接加入窗口并 `Dock = DockStyle.Fill`，不再经过 `ToolStripControlHost`。
- 显示时按当前输入单元格的屏幕位置设置 `Location`，并按工作区边界限幅。
- 窗口初始尺寸用 `SizeFromClientSize` 从目标表格内容尺寸换算，避免标题栏和边框压缩表格。
- `ShowWithoutActivation` 和 `WS_EX_NOACTIVATE` 尽量避免弹窗显示时抢走定额名称编辑框焦点。
- 外部点击、主表格滚动、当前单元格变化、结束编辑等仍沿用现有隐藏逻辑。

## 取舍

- 这会从“无标题下拉框”变成“小工具窗口”视觉，但换来稳定可拖拽、内容随窗口变化。
- 该方案贴近现有推荐定额和联动 Excel 插件窗口的实现方式。

## 验证

- 构建成功，构建产物与工作区内两个软件目录的 DLL 哈希一致。
- 实际软件中拖动窗口边缘和右下角，确认候选表格立即铺满窗口。
- 双击候选定额写入效果保持不变。
- 点击其他窗口或移动主表格焦点时，候选窗口隐藏。