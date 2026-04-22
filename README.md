# Excel-Dr

[![Release](https://img.shields.io/github/v/release/LCbeijing/Excel-Dr?display_name=tag&style=flat-square)](https://github.com/LCbeijing/Excel-Dr/releases)
[![License](https://img.shields.io/github/license/LCbeijing/Excel-Dr?style=flat-square)](./LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-0F172A?style=flat-square)](https://github.com/LCbeijing/Excel-Dr)
[![Format](https://img.shields.io/badge/format-.xlsx-16A34A?style=flat-square)](https://github.com/LCbeijing/Excel-Dr)

<p align="center">
  <img src="./assets/icon.svg" alt="Excel-Dr icon" width="128" />
</p>

<p align="center">
  <strong>一款专门修复 Excel / WPS 报表隐藏垃圾数据问题的便携工具。</strong>
</p>

![Excel-Dr Banner](./assets/banner.svg)

Excel-Dr，中文可以理解为“Excel 医生”。

它专门解决一类非常真实、非常常见、却又特别难排查的问题：

- Excel / WPS 报表打开很慢
- 保存、关闭、筛选时明显卡顿
- 可见数据不算夸张，但文件体验越来越差
- 从旧模板复制表头后，新文件也会继承卡顿问题

很多时候，真正拖慢报表的并不是业务数据本身，而是 `.xlsx` 内部长期积累的隐藏对象、异常绘图层、损坏关系和看不见的垃圾数据。Excel-Dr 做的，就是把这些问题从“说不清、找不到、删不准”，变成一个普通办公用户也能完成的操作流程。

如果要用一句最准确的话描述它：

**它不是在“优化表格外观”，而是在“修复表格内部结构里的隐藏垃圾”。**

## 为什么这个工具值得用

Excel-Dr 的核心价值，不是“粗暴删对象”，而是**精准识别、保守清理、优先安全**。

- **精准**：重点识别异常膨胀的隐藏 `drawing` 对象
- **保守**：不乱动单元格公式、共享字符串、样式、`cellImages` 图片公式
- **轻量**：Windows 便携版，解压即用，删除目录即卸载
- **实用**：单文件、文件夹批量处理都能覆盖
- **面向办公现实**：不是实验室脚本，而是针对财务、运营、订单、仓储等报表场景

## 典型使用场景

- 财务月度订单汇总表
- 电商发货清单、对账表、售后台账
- 微信群 / 人工录入 / 多人接力维护的 Excel 报表
- 用了很久的历史模板
- 复制来复制去、越用越卡的 WPS 表格

## 它到底能解决什么

### 1. 单文件检测

快速告诉你：

- 有没有异常隐藏对象
- 预计清理多少
- 有没有损坏的数据有效性
- 命中问题主要在哪张工作表

### 2. 单文件清理

- 默认另存为 `_cleaned.xlsx`
- 不覆盖原文件
- 清理前会先弹出确认提示

### 3. 文件夹批量检测

直接扫描整个文件夹内的 `.xlsx`：

- 扫描文件数
- 需要处理的文件数
- 异常对象总量
- 正常文件和异常文件分别列出

### 4. 文件夹批量清理

- 只处理真正命中异常的文件
- 正常文件自动跳过
- 每个异常文件各自输出一个 `_cleaned.xlsx`

## 核心策略

Excel-Dr 并不是“看见对象就删”。它只会在满足明显异常特征时才出手。

当前策略主要包括：

- 只处理 `.xlsx`
- 重点检查 `drawing` 绘图层
- 只关注隐藏对象
- 当同一隐藏图片资源在同一锚点或极少数锚点上大量重复出现时，判定为异常膨胀
- 同时清理明显损坏的 `#REF!` 数据有效性规则

这意味着：

- 正常可见图片默认不会删
- 正常图表和普通对象默认不会删
- 单元格公式不会改
- 样式不会改
- `cellImages.xml` 中的单元格图片公式不会改

## 产品截图

### 主界面

![Excel-Dr Screenshot](./docs/images/app-screenshot.png)

### 界面设计思路

- 普通办公人员一看就会用
- 先检测，再清理
- 所有危险动作前都先确认
- 文件夹模式默认跳过正常文件

## 为什么它比手工排查更有效

手工处理这类问题，通常意味着：

- 解压 `.xlsx`
- 手动查 XML
- 人工判断哪些关系该删
- 反复试错

这对大多数办公用户来说几乎不可行。

Excel-Dr 把这件事产品化成了：

1. 选文件或文件夹
2. 看检测结果
3. 确认清理
4. 输出新文件

## 适合谁使用

- 财务人员
- 运营团队
- 订单和仓储团队
- 行政、人事、客服
- 需要长期维护 Excel 模板的人
- 被“老表格越来越卡”折磨过的人

## 安装与运行

### 便携版

推荐直接使用 Windows 便携版：

- 解压
- 双击 `Excel-Dr.exe`
- 用完直接删除目录即可

### Python 运行

```powershell
python cleaner.py
```

### 命令行

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

## 打包

```powershell
build_exe.bat
```

或手动：

```powershell
pip install pyinstaller
pyinstaller --noconsole --onefile cleaner.py --name Excel-Dr
```

## 开源说明

欢迎：

- 提交 issue
- 提交样本和复现描述
- 提出检测策略优化建议
- 贡献更强的规则和 UI 改进

如果你手里有特别卡的报表，而 Excel-Dr 没命中，或者命中了但你希望识别得更细，欢迎反馈。

## 宣传文案

仓库内附了一份可直接拿去发朋友圈、社群、公众号或项目介绍页的宣传素材：

- [宣传文案](./docs/promo-copy.md)

## 愿景

Excel-Dr 想做的，不只是清理几个对象。

它要成为一个真正面向普通办公用户的 Excel 报表修复工具：能解释问题、能定位问题、能处理问题，而且足够轻、足够稳、足够直接。

让那些本来应该服务工作的报表，不再反过来拖累工作。
