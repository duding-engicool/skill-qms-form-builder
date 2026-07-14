# 质量管理表单生成（现场数据收集表）qms-form-builder

面向现场检验员 / 试验员 / 点检员，生成 **SIP/SOP/CP 的下一级执行文件**——用来现场填数据、做记录的空白表，输出 **Excel(.xlsx)** 与 **Word(.docx)** 模板（无网页版）。

## 三类表单
| 类型 | 说明 |
|------|------|
| 检验记录表 | 首检/巡检/终检，记录实测值与判定（合格/不合格下拉） |
| 试验测试报告表 | 试验/测试原始数据记录与结论（合格/不合格下拉） |
| 点检表 | 设备班/日/周点检，记录结果（正常/异常下拉） |

## 用法
```bash
# 依赖
pip install openpyxl python-docx

# 内置样本，产出三类演示模板（xlsx+docx）
python scripts/build_form.py

# 指定类型 + 表头 + 行数
python scripts/build_form.py --type inspection --format both --rows 15 \
    --title "XX产品首件检验记录表" \
    --meta "产品名称=XX" "零件号=ABC-001" "检验阶段(首检/巡检/终检)=首检"

# 仅 Word
python scripts/build_form.py --type checklist --format docx --rows 20
```

产物位于 `outputs/`：
- `form_inspection.xlsx / .docx`
- `form_test.xlsx / .docx`
- `form_checklist.xlsx / .docx`

均为带空白数据行的可填写模板，含表头信息区、签字栏、判定/结果下拉、打印设置，可直接打印下发。

## 联动
- `inspection-record-digester`：本技能生成的空白表→现场填完→拍照/扫描→该技能识别整理分析。
- SIP/SOP/CP 类技能：本技能的检验项目与标准应取自那些指导文件（本技能是其下一级）。
