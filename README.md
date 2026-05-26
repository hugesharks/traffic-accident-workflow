# 🚗 交通事故案件AI办案工作流

> 基于Python的交通事故损害赔偿案件全流程自动化工具，从证据到文书一键生成

## ✨ 核心功能

- **赔偿自动计算** — 13项赔偿项目，严格依据河北省官方标准
- **三期智能评定** — 60+种损伤类型，引用GA/T 1193-2014
- **证据清单生成** — 8大类27项规则库，根据案件类型自动推导
- **要素式起诉状** — 基于法院模板XML替换，样式100%还原
- **勾选框自动填充** — ☐→☑ 根据案件条件自动判断

## 🚀 快速开始

### 安装依赖

```bash
pip install python-docx
```

### 运行示例

```bash
python scripts/workflow_main.py --input examples/case_data_example.json
```

### Python调用

```python
from scripts.workflow_main import run_workflow

case_input = {
    "plaintiff_name": "张某某",
    "case_type": "civil",
    "injury_type": "disability",
    "liability_type": "主责",
    "standard_year": "2025",
    "hospital_days": 45,
    "medical_fee": 52340.80,
    "disability_grade": 10,
    "plaintiff_age": 45,
    # ... 更多字段见 examples/case_data_example.json
}

result = run_workflow(case_input, template_path="模板_民事起诉状.docx")
print(f"标的总额: {result['result']['total_amount']:,.2f}元")
```

## 📄 输出文件

| 文件 | 用途 |
|------|------|
| 赔偿明细报告.docx | 给当事人看，含计算公式+法律依据+保险拆分 |
| 证据清单.docx | 立案提交，6列标准表格 |
| 要素式起诉状.docx | 立案提交，样式100%还原 |
| 赔偿明细.md | 快速查看 |

## 🔧 支持的案件类型

- **伤情**：死亡 / 伤残（1-10级）/ 仅财产损失
- **案件**：纯民事 / 刑事附带民事（自动排除精神损害抚慰金）
- **责任**：全责(100%) / 主责(70%) / 同责(50%) / 次责(30%) / 无责

## 📊 验证案例

| 案号 | 类型 | 标的额 | 核对结果 |
|------|------|--------|----------|
| (2024)冀0531刑初20号 | 死亡·刑事附带民事 | 983,809.83元 | 10项全部精确匹配 ✅ |

## 📋 三期评定数据（GA/T 1193-2014）

```python
from scripts.three_period_data import lookup_three_period

result = lookup_three_period("左股骨颈骨折")
# → 误工180-270天 护理60-120天 营养60-90天
```

覆盖：头部/面部/颈部/胸部/脊柱/四肢/手足/神经等60+种损伤类型

## ⚠️ 重要提示

1. **必须使用法院原件模板** — 要素式起诉状不支持从零生成，样式无法还原
2. **日标准可自定义** — 不同法院伙食/营养标准可能不同
3. **非河北省用户** — 需修改 `HEBEI_STANDARD` 中的赔偿参数为本省数据

## 📁 项目结构

```
├── SKILL.md                      # Coze Skill定义
├── scripts/
│   ├── workflow_main.py          # 主入口
│   ├── lawsuit_generator_v2.py   # 起诉状生成器
│   ├── evidence_list_generator.py # 证据清单生成器
│   ├── compensation_report_generator.py # 赔偿报告生成器
│   └── three_period_data.py      # 三期评定数据
├── references/                   # 法律依据文档
├── examples/                     # 输入数据示例
└── README.md
```

## 📜 许可证

MIT License

## 🙏 致谢

- 赔偿标准依据：《河北省道路交通事故损害赔偿项目计算标准(试行)》
- 三期评定依据：公安部GA/T 1193-2014《人身损害误工期、护理期、营养期评定规范》
- 伤残分级依据：《人体损伤致残程度分级》
