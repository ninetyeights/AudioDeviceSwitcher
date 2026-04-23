# 音频切换助手 (Audio Device Switcher)

WPF (.NET 10 for Windows) 音频设备快速切换工具。常驻系统托盘，支持配置管理、全局快捷键、蓝牙检测。

- **用户可见名（中文）**：音频切换助手
- **内部标识（英文）**：AudioDeviceSwitcher（namespace / AssemblyName / 数据目录 / 注册表项）
- **仓库/文件夹/csproj 文件名**：仍为 `CKit`（未改，避免 sln/命令路径变动）
- **输出 exe**：`AudioDeviceSwitcher.exe`（由 csproj `<AssemblyName>` 决定）

## 运行与构建

```bash
dotnet build              # 构建
dotnet run --project CKit # 运行（项目文件夹仍叫 CKit）
```

构建失败提示"file is locked"时，说明已有实例在运行，需要先 `taskkill //F //IM AudioDeviceSwitcher.exe`（应用关闭主窗口只隐藏到托盘，退出需从托盘菜单）。

## 架构

- `App.xaml.cs` — 应用入口，初始化 `NotifyIcon` 系统托盘，`ShutdownMode = OnExplicitShutdown`
- `MainWindow` — 主窗口，显示配置列表 + 设备列表，关闭时 `Hide()` 到托盘
- `MiniWindow` — 极简小窗口，无边框可拖拽，置顶/透明度/锁定，2 秒轮询设备状态
- `AudioDeviceService` — NAudio 枚举设备，通过未文档化的 COM 接口 `IPolicyConfig.SetDefaultEndpoint` 切换默认设备；`HasBluetoothDevice()` 通过 `PKEY_Device_EnumeratorName == "BTHENUM"` 判断
- `DeviceProfile` + `ProfileService` — 配置模型和 JSON 持久化，存 `%AppData%/AudioDeviceSwitcher/profiles.json`
- `HotkeyService` — Win32 `RegisterHotKey` 全局热键，需要 `partial class` 配合 `LibraryImport`
- `ProfileEditDialog` — 配置的创建/编辑对话框（名称 + 快捷键录入）
- `DeviceIconHelper` — 从系统资源提取设备图标

## 约定

- **所有用户可见的文本使用中文**（按钮、菜单、对话框、错误提示、托盘菜单）
- **代码标识符用英文**
- **csproj 中 `<Using Remove>` 排除了 `System.Drawing` 和 `System.Windows.Forms` 的全局 using**，避免与 WPF 类冲突；需要时在文件内手动 `using` 或完整限定
- **配置文件、设置文件**都存在 `%AppData%/AudioDeviceSwitcher/` 下的 JSON
- **设备 ID 对比**用 `string.Equals`（NAudio 返回的 ID 是稳定的）
- **UI 警告色系**：红色 `#D32F2F`（配置不匹配）、橙色 `#E65100`（蓝牙）、蓝色 `#2196F3`（激活的配置/设备）

## 已知陷阱

- `HotkeyService` 构造时若窗口句柄已存在（如在 `SourceInitialized` 事件中创建）必须立即挂 `WndProc`，不能只依赖 `SourceInitialized` 事件订阅
- WPF `DragMove()` 会捕获鼠标导致按钮 Click 失效，小窗口用 `PreviewMouseLeftButtonDown` + 排除 Button/CheckBox/Slider 父节点来兼顾拖拽和点击
- `Process.Start` 打开 `mmsys.cpl` 不支持 toggle（外部进程，无法可靠关闭）
