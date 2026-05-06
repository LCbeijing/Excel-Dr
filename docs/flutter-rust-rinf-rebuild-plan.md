# Excel-Dr Flutter + Rust + Rinf 重构方案

## 目标

这次重构的目标不是简单换 UI 技术栈，而是把 Excel-Dr 从“Python 单文件工具”升级为可维护、可验证、可扩展的桌面产品：

- Flutter 负责高精度界面、状态展示、文件选择和结果呈现。
- Rust 负责 `.xlsx` 解析、检测、清理、批量任务、进度计算和报告生成。
- Rinf 负责 Flutter 与 Rust 后端之间的类型化消息通信。

产品必须坚持三条底线：

- 不覆盖原文件。
- 先检测，再清理。
- 每一次清理都能解释“发现了什么、处理了什么、输出到哪里”。

## 当前能力盘点

现有 `cleaner.py` 已经具备可迁移价值：

- 支持 `.xlsx` 文件解包和重新打包。
- 支持读取 workbook、worksheet、relationship、drawing 等 OpenXML 结构。
- 支持检测异常隐藏 drawing 对象。
- 支持检测 `#REF!` 损坏数据有效性规则。
- 支持清理后输出 `_cleaned.xlsx` 副本。
- 支持单文件和文件夹批量处理。
- 支持后台线程和进度回调，避免界面假死。

这些能力应该迁移为 Rust 核心库，而不是在 Flutter 中重新写业务逻辑。

## 现有后端主要不足

正式重构前需要处理这些问题：

- 业务逻辑、GUI、CLI 混在同一个 Python 文件里，后续扩展成本高。
- 没有系统测试样本，无法证明清理不会误伤正常图片、图表、公式和样式。
- 检测规则阈值写死在代码中，缺少解释性和配置能力。
- 进度百分比是“任务单元进度”，还不是完全基于真实文件字节或对象数量的精细进度。
- 报告结构偏文本化，前端难以稳定渲染成结果卡片、表格和导出报告。
- 错误类型没有分层，例如文件损坏、密码保护、格式不支持、权限不足、磁盘写入失败应该分别处理。

## Rust 后端模块设计

建议 Rust workspace 拆成以下模块：

```text
excel_dr_core/
  src/
    lib.rs
    model.rs          # Report、Finding、TaskProgress、ErrorCode 等结构
    package.rs        # xlsx zip 包读取、路径解析、关系解析
    workbook.rs       # workbook / worksheet 索引
    drawing.rs        # drawing anchor 解析、异常规则检测
    validation.rs     # 数据有效性规则检测与修复
    cleaner.rs        # 根据检测报告执行清理并输出新文件
    batch.rs          # 文件夹批量扫描与清理
    report.rs         # JSON/Markdown/HTML 报告生成
    progress.rs       # 真实任务进度模型
    error.rs          # 错误分层
```

核心原则：

- `analyze()` 只读文件，不产生任何写入。
- `clean()` 必须基于 `analyze()` 的报告执行，不能绕过检测直接删除。
- 所有清理动作都必须写入新路径。
- Rust 返回结构化结果，Flutter 不解析日志文本。

## Rinf 消息协议

Flutter 与 Rust 之间建议只传这些稳定消息。

### Flutter -> Rust

```text
AnalyzeFile(path)
AnalyzeFolder(path)
CleanFile(source_path, output_path, report_id)
CleanFolder(path, report_id)
CancelTask(task_id)
OpenOutputFolder(path)
ExportReport(report_id, output_path)
```

### Rust -> Flutter

```text
TaskStarted(task_id, task_type)
TaskProgress(task_id, percent, stage, current_item)
AnalyzeCompleted(task_id, report)
CleanCompleted(task_id, report)
TaskFailed(task_id, error)
TaskCancelled(task_id)
```

### 报告结构

```text
Report
  source_path
  output_path
  file_count
  needs_cleanup
  summary
    scanned_files
    cleanable_objects
    broken_rules
    output_files
  findings[]
    worksheet
    issue_type
    count
    action
    severity
  warnings[]
```

这样前端才能稳定展示：

- 扫描文件数
- 可清理对象数
- 损坏规则数
- 输出文件数
- 明细表格
- 运行日志
- 导出报告

## 第一阶段交付边界

第一阶段不是 MVP，也不是只做单文件流程。第一阶段必须完成现有产品的完整重构、隐藏 bug 修复和工程质量补齐，达到可以替代当前 Python 版本的程度。

第一阶段必须包含：

- 单文件检测。
- 单文件清理。
- 文件夹批量检测。
- 文件夹批量清理。
- 正常文件自动跳过。
- 清理结果只生成新文件，不覆盖原文件。
- 真实进度显示。
- 任务取消与失败恢复。
- 结构化检测报告。
- 输出目录打开。
- 使用说明弹窗。
- Windows 免安装包。
- 当前已知路径解析、编码、进度、异常处理等隐藏 bug 修复。

第二阶段才进入“全量 Excel 修复平台”方向，扩展更多 Excel 问题类型和高级能力。

## 第一阶段必须支持的问题

正式版第一阶段不要贪多，先把最有把握、最有用户价值的问题做稳：

1. 隐藏 drawing 对象异常膨胀
   - 价值：解决“文件不大但打开、保存、筛选明显卡顿”的常见问题。
   - 策略：只处理满足异常阈值的隐藏对象。
   - 风险控制：默认不删除普通可见图片、图表、形状。

2. 损坏的数据有效性规则
   - 价值：解决模板复制、引用失效后带来的卡顿或报错。
   - 策略：检测 `#REF!` 等明显损坏规则。
   - 风险控制：只移除损坏规则，不改正常下拉规则。

3. 文件夹批量处理
   - 价值：真实办公用户常常有一批历史报表。
   - 策略：先检测整批文件，只清理命中问题的文件。
   - 风险控制：正常文件跳过，不生成无意义副本。

4. 结构化报告
   - 价值：用户需要知道“软件到底做了什么”。
   - 策略：报告同时服务 UI 展示和导出。

## 第二阶段可扩展的问题库

这些能力属于第二阶段“全量 Excel 修复平台”方向，不放进第一阶段完整重构范围：

- 无用样式膨胀检测。
- 共享字符串异常膨胀检测。
- 空白行列 usedRange 异常扩张检测。
- 外部链接失效检测。
- 过多条件格式检测。
- 损坏图片关系检测。
- WPS 特有兼容结构检测。
- 密码保护文件识别与明确提示。

每增加一种问题，都必须满足：

- 有真实样本。
- 能解释给普通用户听。
- 能只读检测。
- 能安全修复或明确只做提示。
- 有回归测试。

## 进度条设计

进度条不能是假进度。建议 Rust 后端按任务阶段计算：

- 读取文件结构：5%
- 扫描 workbook / worksheet / rels：15%
- 扫描 drawing / validation：40%
- 生成检测报告：5%
- 清理写入新文件：30%
- 校验输出文件：5%

对于大文件，进度事件必须包含：

- 当前阶段
- 当前处理对象
- 已处理对象数
- 总对象数
- 百分比

Flutter 只负责展示，不自己推算。

## 测试样本准备

重构前必须准备样本集，否则无法保证产品可用。

建议建立：

```text
tests/fixtures/
  clean/
    normal_images.xlsx
    normal_chart.xlsx
    normal_validation.xlsx
  dirty/
    hidden_drawing_1000.xlsx
    hidden_drawing_50000.xlsx
    broken_validation_ref.xlsx
    mixed_hidden_drawing_and_validation.xlsx
  edge/
    password_protected.xlsx
    corrupt_zip.xlsx
    missing_relationship.xlsx
    wps_saved.xlsx
```

每个样本要有说明：

- 文件来源或构造方式。
- 预期检测结果。
- 预期清理结果。
- 不允许被修改的内容。

## 验收标准

后端验收：

- 正常文件检测结果为无需清理。
- 命中问题的文件能生成新文件。
- 原文件哈希值保持不变。
- 输出文件能被 Excel / WPS 打开。
- 正常图片、图表、公式、样式不被破坏。
- 大文件处理期间 UI 不假死。
- 用户取消任务后不会留下半成品输出文件。

前端验收：

- 未选择文件时，清理按钮不可用。
- 未完成检测前，清理按钮不可用。
- 检测完成且发现问题后，清理按钮可用。
- 检测完成但无问题时，清理按钮不可用，并明确提示无需处理。
- 处理中显示真实进度、当前阶段和当前对象。
- 失败时显示人能理解的错误原因。

## 正式开发前必须完成的工作

1. 固定最终 UI 原型
   - 当前 v13 原型已确定，可作为 Flutter 实现依据。

2. 输出组件清单
   - 顶栏、文件选择卡片、步骤操作卡片、进度卡片、结果卡片、明细表格、输出位置、使用说明弹窗、确认弹窗、错误提示。

3. 定义 Rust 数据模型
   - 先定 `Report`、`Finding`、`Progress`、`TaskError`，再写实现。

4. 准备测试样本
   - 没有样本前，不建议开始大规模 Rust 重写。

5. 建立回归测试
   - 每个规则都必须有 positive 和 negative 样本。

6. 建立打包流程
   - Windows 免安装 `.exe`。
   - 明确是否需要 VC++ 运行库。
   - 明确输出目录、日志目录和临时目录。

7. 建立日志策略
   - 默认只记录本地运行日志。
   - 不上传文件。
   - 日志不能包含用户敏感表格内容。

## 推荐开发顺序

1. 建 Rust core，实现 `analyze_file` 和 `analyze_folder`。
2. 用 clean / dirty 样本验证检测结果和误判风险。
3. 实现 `clean_file`，保证输出副本可打开且原文件不变。
4. 实现 `clean_folder`，保证正常文件跳过、问题文件输出副本、失败文件不中断整批任务。
5. 增加真实进度事件和取消机制。
6. 接入 Rinf，固定前后端结构化消息协议。
7. Flutter 实现最终 v13 UI。
8. 接入单文件完整流程。
9. 接入文件夹批量完整流程。
10. 加导出报告、使用说明、错误提示和确认弹窗。
11. 补齐隐藏 bug 修复和回归测试。
12. 打包 Windows 免安装版本。

## 不建议做的事

- 不要一开始就把所有 Excel 问题都做进去。
- 不要让 Flutter 直接处理 `.xlsx`。
- 不要让 UI 根据日志字符串判断状态。
- 不要在没有检测报告的情况下允许清理。
- 不要覆盖原文件。
- 不要用“万能优化”这类模糊文案承诺无法验证的效果。
