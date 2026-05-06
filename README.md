<h1 align="center">Excel-Dr</h1>

<p align="center">
  <a href="https://github.com/LCbeijing/Excel-Dr/releases"><img alt="Release" src="https://img.shields.io/github/v/release/LCbeijing/Excel-Dr?display_name=tag&style=flat-square"></a>
  <a href="./LICENSE"><img alt="License" src="https://img.shields.io/github/license/LCbeijing/Excel-Dr?style=flat-square"></a>
  <a href="https://github.com/LCbeijing/Excel-Dr"><img alt="Platform" src="https://img.shields.io/badge/platform-Windows-0F172A?style=flat-square"></a>
  <a href="https://github.com/LCbeijing/Excel-Dr"><img alt="Format" src="https://img.shields.io/badge/format-.xlsx-16A34A?style=flat-square"></a>
</p>

<p align="center">
  <strong>一款专门修复 Excel / WPS 报表隐藏垃圾数据问题的 Windows 便携工具。</strong>
</p>

![Excel-Dr Cover](./docs/images/cover.png)

## What Is It

Excel-Dr，中文可以理解为“Excel 医生”。

它不是一个泛泛而谈的“表格优化器”，而是一个专门面向真实办公报表场景的修复工具：  
**精准识别 `.xlsx` 内部异常膨胀的隐藏对象，并做保守清理。**

很多报表真正拖慢体验的，并不是可见的业务数据，而是长期复制、粘贴、继承模板、多人接力维护之后，悄悄积累在工作簿内部结构里的隐藏垃圾数据。

Excel-Dr 要解决的，就是这种问题：

- 报表打开越来越慢
- 保存、关闭、筛选明显卡顿
- 行数不算离谱，但体验非常糟糕
- 从旧模板复制表头后，新文件继续继承卡顿

## Why It Matters

Excel-Dr 的价值，不在于“清理得多”，而在于“清理得准”。

- **精准识别**：重点锁定异常膨胀的隐藏 `drawing` 对象
- **保守处理**：不粗暴破坏公式、样式、共享字符串和 `cellImages` 图片公式
- **便携易用**：Windows 解压即用，删除目录即卸载
- **适合推广**：支持单文件处理，也支持文件夹批量处理
- **面向现实工作**：为财务、运营、订单、仓储、行政等真实报表场景设计

如果要用一句话概括：

**它不是在优化表格外观，而是在修复表格内部结构里的隐藏垃圾。**

## Typical Use Cases

- 财务月度订单汇总表
- 电商发货清单、售后台账、对账表
- 微信群人工录入订单表
- 多人长期接力维护的历史模板
- WPS / Excel 里那种“看起来不算大，但就是越来越卡”的报表

## Core Capabilities

> 当前版本：V1.1。推荐下载 `Excel-Dr-Single.exe`，双击即可运行。

### 1. 单文件检测

快速告诉你：

- 有没有异常隐藏对象
- 预计清理多少
- 有没有损坏的数据有效性
- 问题主要命中在哪张工作表

### 2. 单文件清理

- 默认另存为 `_cleaned.xlsx`
- 不覆盖原文件
- 清理前先弹出确认提示

### 3. 文件夹批量检测

扫描整个文件夹内的 `.xlsx`：

- 扫描文件数
- 需要处理的文件数
- 异常隐藏对象总量
- 正常文件与异常文件分别列出

### 4. 文件夹批量清理

- 只处理真正命中异常的文件
- 正常文件自动跳过
- 每个异常文件各自输出一个 `_cleaned.xlsx`

## Detection Strategy

Excel-Dr 不是“见对象就删”。它只会在满足明显异常特征时才出手。

当前策略重点包括：

- 只处理 `.xlsx`
- 重点检查 `drawing` 绘图层
- 只关注隐藏对象
- 只有当同一隐藏图片资源在同一锚点或少数锚点上大量重复出现时，才判定为异常膨胀
- 同时清理明显损坏的 `#REF!` 数据有效性规则

这意味着：

- 正常可见图片默认不会删
- 正常图表和普通对象默认不会删
- 单元格公式不会改
- 样式不会改
- `cellImages.xml` 中的单元格图片公式不会改

## Product Interface

### 实际界面截图

![Excel-Dr Screenshot](./docs/images/app-screenshot.png)

### 交互原则

- 普通办公用户一看就能上手
- 先检测，再清理
- 所有风险动作前都先确认
- 文件夹模式默认跳过正常文件

## Why It Beats Manual Cleanup

手工处理这类问题，通常意味着：

- 解压 `.xlsx`
- 手动查 XML
- 人工判断哪些关系该删
- 反复试错

这对大多数办公用户来说几乎不可行。

Excel-Dr 把这件事产品化成了一个非常直接的流程：

1. 选文件或文件夹
2. 看检测结果
3. 确认清理
4. 输出新文件

## Who Should Use It

- 财务人员
- 运营团队
- 订单和仓储团队
- 行政、人事、客服
- 需要长期维护 Excel 模板的人
- 被“老报表越来越卡”困扰过的人

## Run It

### V1.1 Download

GitHub Release page: [v1.1.0](https://github.com/LCbeijing/Excel-Dr/releases/tag/v1.1.0)

Download these assets from the release page:

- `Excel-Dr-Single.exe`
- `Excel-Dr-Flutter-portable.zip`

If the release assets are missing, run the `Windows Release` GitHub Actions workflow with tag `v1.1.0`; it builds and uploads the Windows package on GitHub.

### Windows 单文件版（推荐）

从 [GitHub Releases](https://github.com/LCbeijing/Excel-Dr/releases) 下载：

- `Excel-Dr-Single.exe`

双击即可运行。首次启动会自动把 Flutter/Rust 运行文件解压到用户本地缓存目录，后续会复用缓存。

### Windows 便携目录版

如果你更希望看到完整文件结构，也可以下载：

- `Excel-Dr-Flutter-portable.zip`

解压后运行目录里的 `Excel-Dr.exe`。注意 Flutter Windows 版需要同目录的 DLL 和 `data/` 文件夹，不能只单独复制目录版里的 exe。

### Python

```powershell
python cleaner.py
```

### CLI

单文件：

```powershell
python cleaner.py --scan "C:\path\file.xlsx"
python cleaner.py --clean "C:\path\file.xlsx" --output "C:\path\file_cleaned.xlsx"
```

文件夹：

```powershell
python cleaner.py --scan-folder "C:\path\folder"
python cleaner.py --clean-folder "C:\path\folder"
```

## Build

### Flutter + Rust + Rinf 桌面版

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_flutter_rinf_shell.ps1
```

该脚本会生成：

- `dist/Excel-Dr-Flutter/Excel-Dr.exe`
- `dist/Excel-Dr-Flutter-portable.zip`
- `dist/Excel-Dr-Single.exe`

### 旧 Python/Tkinter 版

```powershell
build_exe.bat
```

或手动：

```powershell
pip install pyinstaller
pyinstaller --noconsole --onefile cleaner.py --name Excel-Dr
```

## Promotion Copy

仓库内附了一份可直接拿去发朋友圈、社群、公众号或项目介绍页的宣传素材：

- [宣传文案](./docs/promo-copy.md)

## Open Source

欢迎：

- 提交 issue
- 提交样本和复现描述
- 提出识别策略优化建议
- 贡献更强的规则和 UI 改进

如果你手里有特别卡的报表，而 Excel-Dr 没命中，或者命中了但你希望识别得更细，欢迎反馈。

## Vision

Excel-Dr 想做的，不只是清理几个对象。

它想成为一个真正面向普通办公用户的 Excel 报表修复工具：  
能解释问题、能定位问题、能处理问题，而且足够轻、足够稳、足够直接。

让那些本来应该服务工作的报表，不再反过来拖累工作。

## Current Limits

- 当前只支持 `.xlsx`。
- 当前重点处理异常膨胀的隐藏 drawing 对象和明显损坏的 `#REF!` 数据有效性。
- 不会清理公式、共享字符串、样式、外部链接、条件格式异常膨胀等更广泛问题；这些属于后续版本。
- V1.1 的 Flutter + Rinf 版暂未提供取消任务和报告导出。
