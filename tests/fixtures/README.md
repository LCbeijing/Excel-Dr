# Excel-Dr 测试样本说明

公开仓库不提交任何真实业务 Excel 工作簿，也不提交由真实业务文件清理得到的输出文件。

## 当前策略

- `tests/fixtures/dirty/`：仅用于本地私有回归，真实问题样本不进入 GitHub。
- `tests/output/`：测试生成目录，已加入 `.gitignore`，不进入 GitHub。
- `tests/fixtures/clean/` 和 `tests/fixtures/batch/`：公开仓库默认不携带 `.xlsx` 文件；需要时可在本地生成合成样本。
- 测试脚本在缺少公开样本或私有样本时会显示 `SKIP`，不会因为仓库未携带业务数据而失败。

## 生成合成样本

如需本地跑完整 clean 样本回归，可执行：

```powershell
python .\scripts\generate_clean_fixtures.py
New-Item -ItemType Directory -Force .\tests\fixtures\batch
Copy-Item .\tests\fixtures\clean\normal_basic.xlsx .\tests\fixtures\batch\normal_basic.xlsx
Copy-Item .\tests\fixtures\clean\normal_chart.xlsx .\tests\fixtures\batch\normal_chart.xlsx
```

这些文件由脚本生成，不包含真实业务数据。

## 私有问题样本

第一阶段开发期间曾使用用户提供的真实办公文件验证：

- 大量隐藏 drawing 对象可检测并清理。
- 损坏数据有效性引用可检测并清理。
- 清理输出为新文件，原文件 SHA256 保持不变。

该类样本包含真实业务数据风险，已从发布副本和公开提交中移除。后续如果需要保留问题样本，应优先制作脱敏后的最小复现文件，或使用脚本合成结构性问题样本。
