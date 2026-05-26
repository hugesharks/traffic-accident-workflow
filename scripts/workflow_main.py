#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
交通事故案件工作流 - 端到端主入口
一键输入案件数据 → 输出三份文件（赔偿明细报告 + 证据清单 + 要素式起诉状）

使用方式：
    1. Python调用：from workflow_main import run_workflow; run_workflow(case_input)
    2. 命令行：python workflow_main.py --input case_data.json
"""

import os
import sys
import json
from datetime import datetime

# 确保能找到同目录模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from lawsuit_generator_v2 import CaseData, HEBEI_STANDARD, LawsuitGenerator, generate_compensation_report
from evidence_list_generator import collect_evidence, generate_evidence_list_docx, generate_evidence_list_text
from compensation_report_generator import generate_compensation_report_docx


# ============================================================
# 案件数据输入结构
# ============================================================
CASE_INPUT_TEMPLATE = {
    # === 必填 ===
    "plaintiff_name": "原告姓名",
    "plaintiff_gender": "男/女",
    "plaintiff_birthdate": "1970年6月15日",
    "plaintiff_ethnicity": "汉族",
    "plaintiff_phone": "13831910001",
    "plaintiff_id_number": "130531197006150011",
    "plaintiff_address": "河北省邢台市××区××路××号",
    "plaintiff_age": 55,  # 受害人年龄（计算赔偿年限用）

    # === 案件类型 ===
    "case_type": "civil",           # civil=纯民事 | criminal_attached=刑事附带民事
    "injury_type": "death",         # death=死亡 | disability=伤残 | property=仅财产损失
    "liability_type": "主责",        # 全责/主责/同责/次责/无责

    # === 赔偿标准 ===
    "standard_year": "2024",        # 适用哪年标准
    "region": "urban",              # urban=城镇 | rural=农村

    # === 医疗信息 ===
    "hospital_name": "××医院",
    "hospital_days": 11,
    "medical_fee": 102309.83,
    "hospital_meal_daily": 50,      # 可选，默认100
    "nutrition_days": 11,
    "nutrition_daily": 20,          # 可选，默认30
    "nursing_days": 11,
    "nursing_persons": 1,
    "nursing_custom_daily": 139,    # 可选，默认130
    "lost_work_days": 11,
    "lost_work_daily": 178,         # 可选，默认178

    # === 其他费用 ===
    "traffic_fee": 3000,
    "property_damage": 2000,
    "other_fee": 0,
    "other_fee_desc": "",

    # === 伤残信息（伤残案件填） ===
    "disability_grade": None,       # 1-10级

    # === 被扶养人 ===
    "dependents": [],               # [{age, years, share_ratio, relation}]

    # === 事故信息 ===
    "accident_date": "2023年10月28日",
    "accident_detail": "2023年10月28日13时50分，被告驾驶××号小型轿车与被害人发生交通事故...",
    "responsibility_doc_number": "第××××××××××号",
    "responsibility_result": "本次事故经××交警大队认定，被告负主要责任，被害人负次要责任。",
    "insurance_info": "被告驾驶的车辆在××保险公司投保交强险和商业三者险200万元。",

    # === 被告信息 ===
    "defendants_person": [          # 自然人被告
        {"name": "潘某坤", "gender": "男", "id_number": "130531199001010011", "address": "河北省邢台市××县"}
    ],
    "defendants_company": [         # 法人被告（保险公司等）
        {"name": "中国某某财产保险股份有限公司邢台市中心支公司", "address": "河北省邢台市××区",
         "legal_person": "×××", "credit_code": "91130500××××××××××"}
    ],

    # === 法院 ===
    "court": "河北省××县人民法院",

    # === 自定义证据（可选追加） ===
    "custom_evidence": [],
}


def build_case_data(case_input: dict) -> CaseData:
    """将输入字典转为CaseData对象"""
    case = CaseData()

    # 案件类型
    case_type_map = {'civil': CaseData.CASE_TYPE_CIVIL, 'criminal_attached': CaseData.CASE_TYPE_CRIMINAL_ATTACHED}
    case.case_type = case_type_map.get(case_input.get('case_type', 'civil'), CaseData.CASE_TYPE_CIVIL)

    injury_map = {'death': CaseData.INJURY_DEATH, 'disability': CaseData.INJURY_DISABILITY, 'property': CaseData.INJURY_PROPERTY}
    case.injury_type = injury_map.get(case_input.get('injury_type', 'disability'), CaseData.INJURY_DISABILITY)

    case.standard_year = case_input.get('standard_year', '2024')
    case.region = case_input.get('region', 'urban')
    case.liability_type = case_input.get('liability_type', '全责')

    # 原告
    case.plaintiffs = [{
        'name': case_input.get('plaintiff_name', '×××'),
        'gender': case_input.get('plaintiff_gender', '男'),
        'birthdate': case_input.get('plaintiff_birthdate', ''),
        'ethnicity': case_input.get('plaintiff_ethnicity', '汉族'),
        'phone': case_input.get('plaintiff_phone', ''),
        'id_number': case_input.get('plaintiff_id_number', ''),
        'address': case_input.get('plaintiff_address', ''),
    }]

    # 被告
    case.defendants_person = case_input.get('defendants_person', [])
    case.defendants_company = case_input.get('defendants_company', [])

    # 医疗
    case.hospital_name = case_input.get('hospital_name', '')
    case.hospital_days = case_input.get('hospital_days', 0)
    case.medical_fee = case_input.get('medical_fee', 0)
    case.hospital_meal_daily = case_input.get('hospital_meal_daily')
    case.nutrition_days = case_input.get('nutrition_days', 0)
    case.nutrition_daily = case_input.get('nutrition_daily')
    case.nursing_days = case_input.get('nursing_days', 0)
    case.nursing_persons = case_input.get('nursing_persons', 1)
    case.nursing_custom_daily = case_input.get('nursing_custom_daily')
    case.lost_work_days = case_input.get('lost_work_days', 0)
    case.lost_work_daily = case_input.get('lost_work_daily')

    # 其他费用
    case.traffic_fee = case_input.get('traffic_fee', 0)
    case.property_damage = case_input.get('property_damage', 0)
    case.other_fee = case_input.get('other_fee', 0)
    case.other_fee_desc = case_input.get('other_fee_desc', '')

    # 伤残
    case.disability_grade = case_input.get('disability_grade')
    case.plaintiff_age = case_input.get('plaintiff_age', 0)

    # 被扶养人
    case.dependents = case_input.get('dependents', [])

    # 事故信息
    case.accident_detail = case_input.get('accident_detail', '')
    case.responsibility_doc_number = case_input.get('responsibility_doc_number', '')
    case.responsibility_result = case_input.get('responsibility_result', '')
    case.insurance_info = case_input.get('insurance_info', '')
    case.evidence_list = case_input.get('custom_evidence', [])
    case.court = case_input.get('court', '')
    case.filing_date = case_input.get('filing_date', '')

    return case


def _case_to_report_dict(case: CaseData, case_input: dict) -> dict:
    """转为赔偿报告所需的字典"""
    return {
        'injury_type': case.injury_type,
        'case_type': case.case_type,
        'standard_year': case.standard_year,
        'region': case.region,
        'liability_type': case.liability_type,
        'hospital_days': case.hospital_days,
        'hospital_meal_daily': case.hospital_meal_daily,
        'nutrition_days': case.nutrition_days,
        'nutrition_daily': case.nutrition_daily,
        'nursing_days': case.nursing_days,
        'nursing_persons': case.nursing_persons,
        'nursing_custom_daily': case.nursing_custom_daily,
        'lost_work_days': case.lost_work_days,
        'lost_work_daily': case.lost_work_daily,
        'disability_grade': case.disability_grade,
        'plaintiff_age': case.plaintiff_age,
        'dependents': case.dependents,
        'other_fee_desc': case.other_fee_desc,
        'plaintiff_name': case_input.get('plaintiff_name', ''),
        'defendant_text': '、'.join(
            [d['name'] for d in case_input.get('defendants_person', [])] +
            [d['name'] for d in case_input.get('defendants_company', [])]
        ),
        'accident_date': case_input.get('accident_date', ''),
    }


def _case_to_evidence_dict(case: CaseData, case_input: dict) -> dict:
    """转为证据清单所需的字典"""
    return {
        'injury_type': case.injury_type,
        'case_type': case.case_type,
        'medical_fee': case.medical_fee,
        'hospital_days': case.hospital_days,
        'nursing_days': case.nursing_days,
        'nursing_custom_daily': case.nursing_custom_daily,
        'lost_work_days': case.lost_work_days,
        'lost_work_daily': case.lost_work_daily,
        'traffic_fee': case.traffic_fee,
        'property_damage': case.property_damage,
        'dependents': case.dependents,
        'custom_evidence': case.evidence_list,
        'plaintiff_name': case_input.get('plaintiff_name', ''),
        'defendant_text': '、'.join(
            [d['name'] for d in case_input.get('defendants_person', [])] +
            [d['name'] for d in case_input.get('defendants_company', [])]
        ),
        'court': case_input.get('court', ''),
    }


def run_workflow(case_input: dict, output_dir: str = None, template_path: str = None) -> dict:
    """
    端到端工作流：一键生成三份文件

    参数:
        case_input: 案件数据字典（参照 CASE_INPUT_TEMPLATE）
        output_dir: 输出目录（默认为 交通事故案件工作流/案件名/）
        template_path: 起诉状模板路径

    返回:
        {
            'success': True,
            'files': {
                'compensation_report': '路径',
                'evidence_list': '路径',
                'lawsuit': '路径',
            },
            'result': {...计算结果...},
            'summary': '摘要文本',
        }
    """
    # 构建CaseData
    case = build_case_data(case_input)

    # 计算赔偿
    result = case.calculate()

    # 确定输出目录
    plaintiff_name = case_input.get('plaintiff_name', '未命名')
    if output_dir is None:
        output_dir = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            'output',
            plaintiff_name
        )
    os.makedirs(output_dir, exist_ok=True)

    # 被告名称列表
    defendant_names = (
        [d['name'] for d in case_input.get('defendants_person', [])] +
        [d['name'] for d in case_input.get('defendants_company', [])]
    )

    # ========================================
    # 1. 赔偿明细报告
    # ========================================
    report_dict = _case_to_report_dict(case, case_input)
    report_path = os.path.join(output_dir, f'赔偿明细报告_{plaintiff_name}.docx')
    generate_compensation_report_docx(
        report_dict, result,
        output_path=report_path,
        plaintiff_name=plaintiff_name,
        defendant_names=defendant_names,
        accident_date=case_input.get('accident_date', ''),
    )

    # ========================================
    # 2. 证据清单
    # ========================================
    evidence_dict = _case_to_evidence_dict(case, case_input)
    evidence_path = os.path.join(output_dir, f'证据清单_{plaintiff_name}.docx')
    generate_evidence_list_docx(
        evidence_dict,
        output_path=evidence_path,
        plaintiff_name=plaintiff_name,
        defendant_names=defendant_names,
        court=case_input.get('court', ''),
    )

    # ========================================
    # 3. 要素式起诉状
    # ========================================
    if template_path is None:
        template_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            '模板_民事起诉状.docx'
        )

    if os.path.exists(template_path):
        # 更新case的证据列表为规则库生成的版本
        case.evidence_list = generate_evidence_list_text(evidence_dict).rstrip('。').split('；')
        case.evidence_list = [e.strip() for e in case.evidence_list]

        lawsuit_path = os.path.join(output_dir, f'要素式起诉状_{plaintiff_name}.docx')
        gen = LawsuitGenerator(template_path, output_dir)
        lawsuit_output = gen.generate(case, f'要素式起诉状_{plaintiff_name}.docx')
    else:
        lawsuit_output = None

    # ========================================
    # 4. 赔偿明细Markdown（轻量版，方便快速查看）
    # ========================================
    md_path = os.path.join(output_dir, f'赔偿明细_{plaintiff_name}.md')
    md_report = generate_compensation_report(case, result)
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(md_report)

    # ========================================
    # 汇总
    # ========================================
    files = {
        'compensation_report': report_path,
        'evidence_list': evidence_path,
        'compensation_md': md_path,
    }
    if lawsuit_output:
        files['lawsuit'] = lawsuit_output

    # 生成摘要
    injury_text = {'death': '死亡', 'disability': '伤残', 'property': '仅财产损失'}.get(case.injury_type, '未知')
    liability_text = case.liability_type
    summary = (
        f"案件：{plaintiff_name}机动车交通事故责任纠纷\n"
        f"类型：{injury_text} | {liability_text}\n"
        f"标的总额：{result['total_amount']:,.2f}元\n"
        f"保险公司赔付：{result['insurance_split']['total_insurance']:,.2f}元\n"
        f"生成文件：{len(files)}份"
    )

    return {
        'success': True,
        'files': files,
        'result': result,
        'summary': summary,
    }


# ============================================================
# 命令行入口
# ============================================================
if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='交通事故案件工作流 - 一键生成')
    parser.add_argument('--input', '-i', help='案件数据JSON文件路径')
    parser.add_argument('--output', '-o', help='输出目录')
    parser.add_argument('--template', '-t', help='起诉状模板路径')
    parser.add_argument('--demo', action='store_true', help='使用冉某江案演示数据')
    args = parser.parse_args()

    if args.demo or not args.input:
        # 演示数据：冉某江案
        case_input = {
            "plaintiff_name": "冉某江",
            "plaintiff_gender": "男",
            "plaintiff_birthdate": "1970年6月15日",
            "plaintiff_ethnicity": "汉族",
            "plaintiff_phone": "13831910001",
            "plaintiff_id_number": "130531197006150011",
            "plaintiff_address": "河北省邢台市广宗县塘疃乡常阜村",
            "plaintiff_age": 55,

            "case_type": "criminal_attached",
            "injury_type": "death",
            "liability_type": "主责",

            "standard_year": "2023",
            "region": "urban",

            "hospital_days": 11,
            "medical_fee": 102309.83,
            "hospital_meal_daily": 50,
            "nutrition_days": 11,
            "nutrition_daily": 20,
            "nursing_days": 11,
            "nursing_custom_daily": 139,
            "lost_work_days": 11,
            "lost_work_daily": 178,

            "traffic_fee": 3000,
            "property_damage": 2000,
            "other_fee": 0,

            "accident_date": "2023年10月28日",
            "accident_detail": "2023年10月28日13时50分，被告驾驶冀E1××某某号小型轿车由北向南行驶至广宗县常阜村南十字路口处时，与前方同向行驶向左转弯的被害人高某芝驾驶的电动三轮车相碰撞，致高某芝受伤，后经医院抢救无效于2023年11月8日死亡。",
            "responsibility_result": "本次事故经广宗县公安局交通警察大队道路交通事故认定书认定，被告人潘某坤负本起事故的主要责任，被害人高某芝负次要责任。",
            "insurance_info": "被告驾驶的冀E1××某某号小型轿车在中国某某财产保险股份有限公司邢台市中心支公司投保机动车交强险和商业第三者责任保险（保险金额200万元）。",

            "defendants_person": [
                {"name": "潘某坤", "gender": "男"}
            ],
            "defendants_company": [
                {"name": "中国某某财产保险股份有限公司邢台市中心支公司"}
            ],

            "court": "河北省广宗县人民法院",
            "filing_date": "2024年2月5日",
        }
    else:
        with open(args.input, 'r', encoding='utf-8') as f:
            case_input = json.load(f)

    # 执行工作流
    print("=" * 60)
    print("交通事故案件工作流 - 端到端生成")
    print("=" * 60)

    result = run_workflow(case_input, output_dir=args.output, template_path=args.template)

    print(f"\n{result['summary']}")
    print(f"\n生成文件：")
    for name, path in result['files'].items():
        print(f"  ✓ {name}: {path}")

    print(f"\n标的总额: {result['result']['total_amount']:,.2f}元")
    ins = result['result']['insurance_split']
    print(f"保险公司赔付: {ins['total_insurance']:,.2f}元")
    print(f"  - 交强险: {ins['compulsory']:,.2f}元")
    print(f"  - 商业险: {ins['commercial']:,.2f}元")
    print(f"原告自担: {ins['self_bear']:,.2f}元")
