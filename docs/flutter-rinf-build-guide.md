# Excel-Dr Flutter + Rinf 开发与构建说明

更新时间：2026-04-30

## 当前结论

当前仓库已经有可复用的 Rust core：

- `rust/excel_dr_core`
- 已通过 `cargo test`
- 已通过 `scripts/smoke_test_rust_backend.py`

要实现 Flutter + Rinf 桌面壳，缺的不是后端能力，而是本机 Flutter/Dart/Rinf 工具链和 Flutter 工程接入。

## 本机当前状态

已确认可用：

- Rust / Cargo
- Git for Windows
- Visual Studio 2026 Community
- Visual Studio 2026 Build Tools
- Visual Studio 2022 Build Tools
- CMake
- Flutter SDK：`C:\src\flutter`
- Dart：随 Flutter SDK 安装
- Rinf CLI：`C:\Users\LC.DESKTOP-SENLCL9\.cargo\bin\rinf.exe`

当前说明：

- Android SDK 未安装，但不影响 Windows 桌面开发。
- Flutter 下载源已切换为 `https://storage.flutter-io.cn`，Pub 源已切换为 `https://pub.flutter-io.cn`。
- Visual Studio 2026 Community 被 Flutter 优先选择；本机已用 VS 2026 Build Tools 补齐 Community 缺失的 MSBuild / VC bin / C++ targets。

## 安装要求

Windows 桌面 Flutter 开发至少需要：

- Windows 10/11 64-bit
- Git for Windows
- Visual Studio 2022，安装 `Desktop development with C++`
- Flutter SDK，并把 `flutter\bin` 加入 PATH
- Rust toolchain
- Rinf CLI

安装后用下面命令确认：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\flutter_rinf_env.ps1
```

## 推荐工程结构

建议不要把 Flutter 工程直接创建在仓库根目录，避免覆盖现有文件。推荐：

```text
F:\Excel-Dr
  apps/
    excel_dr_flutter/
  rust/
    excel_dr_core/
```

## 创建 Flutter + Rinf 壳

当前已创建工程：`apps/excel_dr_flutter`。

已执行过：

```powershell
flutter create apps\excel_dr_flutter --platforms=windows
flutter pub add rinf
cargo install rinf_cli --locked
rinf template
rinf gen
```

如果 `flutter doctor -v` 还有 Windows toolchain 错误，先修 doctor，不要继续接业务代码。

## 接入现有 Rust core

Rinf 模板生成后，会有：

```text
apps/excel_dr_flutter/native/hub/Cargo.toml
apps/excel_dr_flutter/native/hub/src/lib.rs
```

在 `apps/excel_dr_flutter/native/hub/Cargo.toml` 中加入本仓库 core 依赖：

```toml
[dependencies]
excel_dr_core = { path = "../../../../rust/excel_dr_core" }
```

该依赖当前已经加入，并通过：

```powershell
cargo check --manifest-path .\apps\excel_dr_flutter\native\hub\Cargo.toml
```

然后在 `native/hub/src/lib.rs` 里做 Rinf 消息桥接：

- Dart -> Rust：
  - `AnalyzeFile { path }`
  - `CleanFile { path, output }`
  - `AnalyzeFolder { path }`
  - `CleanFolder { path }`
  - `CancelTask { task_id }`
- Rust -> Dart：
  - `TaskStarted`
  - `TaskProgress`
  - `AnalyzeCompleted`
  - `CleanCompleted`
  - `TaskFailed`
  - `TaskCancelled`

当前已经完成首版业务桥接：

- Flutter 发送 `TaskRequest`。
- Rust 回传 `TaskProgress` 和 `TaskResult`。
- Rust hub 在后台线程中调用 `excel_dr_core`。
- 文件夹检测/清理按文件回传进度。

后续仍需补齐 `CancelTask` 和报告导出。

## Flutter UI 开发目标

UI 以最终原型为准：

- `docs/prototypes/excel-dr-flutter-rinf-ui-v13.html`
- GitHub 公开仓库不提交中间截图导出；以 HTML 原型和 Flutter 实现为准。

Flutter 只负责：

- 选择文件/文件夹
- 按钮状态机
- 展示进度和结构化报告
- 打开输出位置
- 使用说明弹窗

Flutter 不解析 `.xlsx`，也不通过日志字符串判断结果。

## Windows 构建命令

开发运行：

```powershell
cd F:\Excel-Dr\apps\excel_dr_flutter
flutter run -d windows
```

Release 构建：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_flutter_rinf_shell.ps1
```

产物通常在：

```text
apps/excel_dr_flutter/build/windows/x64/runner/Release/
```

当前已成功生成：

- `apps/excel_dr_flutter/build/windows/x64/runner/Debug/excel_dr_flutter.exe`
- `apps/excel_dr_flutter/build/windows/x64/runner/Release/excel_dr_flutter.exe`
- `dist/Excel-Dr-Flutter/Excel-Dr.exe`
- `dist/Excel-Dr-Flutter-portable.zip`

最终免安装包应复制整个 `Release` 目录，而不只是单个 `.exe`，因为 Flutter Windows release 旁边通常还有 DLL 和 data 目录。

当前整理后的便携目录：

```text
dist/Excel-Dr-Flutter/
  Excel-Dr.exe
  hub.dll
  flutter_windows.dll
  file_selector_windows_plugin.dll
  native_assets.json
  data/
```

## 当前已知提示

Rinf/cargokit 构建结束时会打印一条：

```text
Get-Item : 找不到项 C:\Users\LC.DESKTOP-SENLCL9\AppData
```

当前该信息没有导致构建失败，`flutter build windows --debug` 和 `flutter build windows --release` 均返回成功并生成 exe。后续如要清理，可继续调查 `windows/flutter/ephemeral/.plugin_symlinks/rinf/cargokit/cmake/resolve_symlinks.ps1` 对短路径或中文用户路径的处理。

## 验收顺序

1. `flutter doctor -v` 全部关键 Windows 项通过。
2. 空 Flutter Windows app 能 `flutter run -d windows`。
3. Rinf 模板 app 能 `flutter run -d windows`。
4. `native/hub` 能依赖 `rust/excel_dr_core` 并编译。
5. Flutter 发 `AnalyzeFile`，Rust 返回结构化报告。
6. Flutter 发 `CleanFile`，输出新文件且原文件 hash 不变。
7. 文件夹批量检测/清理通过。
8. 保持下面命令通过：

```powershell
python .\scripts\smoke_test_current_backend.py
python .\scripts\smoke_test_rust_backend.py
cargo test --manifest-path .\rust\excel_dr_core\Cargo.toml
```
