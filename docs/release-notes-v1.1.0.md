# Excel-Dr v1.1.0

V1.1 是从旧 Python/Tkinter 工具向 Flutter + Rust + Rinf 桌面产品迁移后的 MVP 版本。

## 新增

- 新增 Flutter + Rinf 桌面界面，按最终 v13 原型实现。
- 新增 Rust core 后端：
  - 单文件检测
  - 单文件清理
  - 文件夹批量检测
  - 文件夹批量清理
- 新增 Windows 单文件启动版：`Excel-Dr-Single.exe`。
- 新增 Flutter 便携目录版：`Excel-Dr-Flutter-portable.zip`。
- 新增结构化任务结果和进度回传。
- 新增 Rust 回归测试和 Flutter UI smoke test。

## 修复和优化

- 清理输出不覆盖原文件，存在同名输出时自动追加序号。
- 批量任务会跳过 `_cleaned.xlsx`、`_cleaned_2.xlsx` 等输出副本。
- 正常文件执行清理时不再生成无意义副本。
- 批量失败文件会计入扫描总数，并保留失败文件名。
- 全失败场景会在 UI 明细表中展示错误行。
- 修复 relationship target 使用 `/xl/...` 绝对包路径时的解析问题。
- zip 解包路径增加基础归一化，降低异常条目越出临时目录的风险。
- 单文件外壳增加缓存完整性检查和并发启动互斥。

## 下载

推荐下载：

- `Excel-Dr-Single.exe`

备选：

- `Excel-Dr-Flutter-portable.zip`

## 校验

- `Excel-Dr-Single.exe` SHA256：`6BCF8F44C0CAE0563B2563FAB9A8C4421BA4B2560A8707401B4EBC460F225AEC`
- `Excel-Dr-Flutter-portable.zip` SHA256：`46E89C72D265B54A31E8452D8704D188D1C5F441C806C8001EE3E7D8D2FF0DC4`

## 已知限制

- 当前只支持 `.xlsx`。
- 当前重点处理异常隐藏 drawing 对象和明显损坏的 `#REF!` 数据有效性。
- 取消任务和报告导出将在后续版本继续补齐。
