#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
交通事故要素式起诉状自动生成脚本 v2.0
支持：死亡/伤残/仅车损 | 纯民事/刑事附带民事 | 多原告 | 责任比例 | 结构化赔偿计算
"""

import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from copy import deepcopy

# 注册命名空间
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)

# ============================================================
# 河北省赔偿标准参数
# ============================================================
HEBEI_STANDARD = {
    # 2023年标准（死亡案件适用事故发生时标准）
    '2023': {
        'urban_income': 41278,      # 城镇居民人均可支配收入
        'urban_consumption': 26252, # 城镇居民人均消费支出
        'rural_income': 19394,
        'rural_consumption': 16447,
        'funeral': 46683,           # 丧葬费（全省在岗职工半年平均工资）
    },
    # 2024年标准（河北省统计局2025.6.26发布分行业工资）
    '2024': {
        'urban_income': 45610,      # 城镇居民人均可支配收入（2024年度）
        'urban_consumption': 29310, # 城镇居民人均消费支出（2024年度）
        'rural_income': 21044,
        'rural_consumption': 18088,
        'funeral': 49545.50,        # 丧葬费 = 99091/2（在岗职工半年平均工资）
    },
    # 2025年标准（河北省2025年统计公报，2026.4.13发布）
    # 注：分行业工资仍采用2024年数据（2025年分行业数据预计2026年6月发布）
    '2025': {
        'urban_income': 47544,      # 城镇居民人均可支配收入
        'urban_consumption': 30522, # 城镇居民人均消费支出
        'rural_income': 23319,
        'rural_consumption': 19488,
        'funeral': 49545.50,        # 丧葬费 = 99091/2（2024年在岗职工平均工资/2）
    },

    # === 通用参数（依据河北省标准） ===
    'hospital_meal': 100,    # 住院伙食补助 元/天（邢台标准，各地市不同）
    'nutrition': 20,         # 营养费 元/天（河北省标准：20元/天×住院天数）
    'nursing': 163.34,       # 护理费 元/天（居民服务业59618元/年÷365）
    'default_wage': 163.34,  # 默认误工费 元/天（无固定收入时按居民服务业标准）
    'city_traffic': 20,      # 市内交通费 元/天（河北省标准：20元/天×门诊次数或住院天数）
    'non_medical_deduction_max': 0.10,  # 非医保用药扣减比例上限（河北省标准：最高10%）

    # 伤残赔偿系数
    'disability_ratio': {
        10: 0.10, 9: 0.20, 8: 0.30, 7: 0.40, 6: 0.50,
        5: 0.60, 4: 0.70, 3: 0.80, 2: 0.90, 1: 1.0
    },

    # 精神损害抚慰金（纯民事案件适用，河北省标准）
    # 按伤残十个等级且最高额为50000元对照赔付
    # 十级1000-5000，九级500-10000，...一级40000-50000，死亡50000
    'mental_damage': {
        10: 5000, 9: 10000, 8: 15000, 7: 20000, 6: 25000,
        5: 30000, 4: 40000, 3: 50000, 2: 50000, 1: 50000,
        'death': 50000,  # 死亡案件统一按50000元
    },

    # 交强险分项限额
    'compulsory_insurance': {
        'medical': 18000,
        'death_disability': 180000,
        'property': 2000
    },

    # 河北省在岗职工分行业年平均工资（2024年度，河北省统计局2025.6.26发布）
    # 用于误工费按行业计算、丧葬费计算
    'industry_wages_2024': {
        '农林牧渔业': 83654,
        '采矿业': 110565,
        '制造业': 90515,
        '电力热力燃气及水': 148962,
        '建筑业': 91955,
        '交通运输仓储邮政': 112643,
        '信息传输软件IT': 135956,
        '批发和零售业': 79452,
        '住宿和餐饮业': 55539,
        '金融业': 160961,
        '房地产业': 73020,
        '租赁和商务服务业': 66603,
        '科学研究技术服务业': 117863,
        '水利环境公共设施': 48530,
        '居民服务修理其他服务业': 59618,
        '教育': 103999,
        '卫生和社会工作': 106496,
        '文化体育娱乐业': 88649,
        '公共管理社会保障': 86864,
        '在岗职工平均': 99091,
    },

    # 责任比例默认值
    'liability_ratio': {
        '全责': 1.0,
        '主责': 0.7,
        '同责': 0.5,
        '次责': 0.3,
        '无责': 0.0,
    },
}


# ============================================================
# 赔偿计算引擎
# ============================================================
class CompensationCalculator:
    """赔偿计算引擎"""

    def __init__(self, standard='2024', region='urban'):
        self.standard = HEBEI_STANDARD.get(standard, HEBEI_STANDARD['2024'])
        self.region = region  # urban / rural

    def calc_medical_fee(self, amount):
        """医疗费：按实际票据"""
        return amount

    def calc_hospital_meal(self, days):
        """住院伙食补助费"""
        return days * HEBEI_STANDARD['hospital_meal']

    def calc_nutrition(self, days):
        """营养费"""
        return days * HEBEI_STANDARD['nutrition']

    def calc_nursing(self, days, persons=1, custom_daily=None):
        """护理费"""
        daily = custom_daily or HEBEI_STANDARD['nursing']
        return days * persons * daily

    def calc_lost_wage(self, days, daily_wage=None):
        """误工费"""
        daily = daily_wage or HEBEI_STANDARD['default_wage']
        return days * daily

    def calc_traffic_fee(self, amount):
        """交通费：按实际票据"""
        return amount

    def calc_disability_compensation(self, grade, age, years=20):
        """残疾赔偿金"""
        income = self.standard['urban_income'] if self.region == 'urban' else self.standard['rural_income']
        ratio = HEBEI_STANDARD['disability_ratio'].get(grade, 0.10)
        if age >= 60:
            years = max(5, years - (age - 60))
        elif age < 60:
            years = 20
        return income * ratio * years

    def calc_death_compensation(self, age, years=20):
        """死亡赔偿金"""
        income = self.standard['urban_income'] if self.region == 'urban' else self.standard['rural_income']
        if age >= 60:
            years = max(5, years - (age - 60))
        elif age < 60:
            years = 20
        return income * years

    def calc_funeral(self):
        """丧葬费"""
        return self.standard['funeral']

    def calc_mental_damage(self, case_type, injury_type, grade=None):
        """精神损害抚慰金"""
        # 刑事附带民事不支持
        if case_type == 'criminal_attached':
            return 0
        if injury_type == 'death':
            return HEBEI_STANDARD['mental_damage']['death']
        elif injury_type == 'disability' and grade:
            return HEBEI_STANDARD['mental_damage'].get(grade, 5000)
        return 0

    def calc_dependent_living(self, dependents):
        """
        被扶养人生活费（分段计算）
        dependents: [{age, years, share_ratio}, ...]
        年赔偿总额不超过人均消费支出
        """
        consumption = self.standard['urban_consumption'] if self.region == 'urban' else self.standard['rural_consumption']
        total = 0
        for dep in dependents:
            total += consumption * dep['share_ratio'] * dep['years']
        return total

    def calc_insurance_split(self, loss_detail, liability_type='全责'):
        """
        交强险/商业险分项拆解（严格按分项限额）
        
        loss_detail: 分项损失明细，格式：
            {
                'medical': 医疗费+住院伙食补助费+营养费,
                'death_disability': 死亡/残疾赔偿金+误工费+护理费+交通费+丧葬费+精神抚慰金+被扶养人生活费+辅助器具费,
                'property': 财产损失+其他财产性费用
            }
        liability_type: 被告（侵权方）的责任类型，全责/主责/同责/次责/无责
        
        返回: {compulsory, commercial, self_bear, compulsory_detail, total_insurance}
        """
        ratio = HEBEI_STANDARD['liability_ratio'].get(liability_type, 1.0)
        limits = HEBEI_STANDARD['compulsory_insurance']

        # 交强险严格按分项限额赔付（不分责任比例）
        comp_medical = min(loss_detail.get('medical', 0), limits['medical'])
        comp_death_dis = min(loss_detail.get('death_disability', 0), limits['death_disability'])
        comp_property = min(loss_detail.get('property', 0), limits['property'])
        compulsory = comp_medical + comp_death_dis + comp_property

        # 剩余部分按被告责任比例由商业险赔付
        total_loss = loss_detail.get('medical', 0) + loss_detail.get('death_disability', 0) + loss_detail.get('property', 0)
        remaining = total_loss - compulsory
        commercial = remaining * ratio
        self_bear = remaining * (1 - ratio)

        return {
            'compulsory': compulsory,
            'commercial': commercial,
            'self_bear': self_bear,
            'total_insurance': compulsory + commercial,
            'compulsory_detail': {
                'medical': comp_medical,
                'death_disability': comp_death_dis,
                'property': comp_property,
            },
        }


# ============================================================
# 案件数据结构
# ============================================================
class CaseData:
    """案件数据"""

    # 案件类型
    CASE_TYPE_CIVIL = 'civil'                    # 纯民事
    CASE_TYPE_CRIMINAL_ATTACHED = 'criminal_attached'  # 刑事附带民事

    # 伤情类型
    INJURY_DEATH = 'death'          # 死亡
    INJURY_DISABILITY = 'disability'  # 伤残
    INJURY_PROPERTY = 'property'     # 仅车损/财产损失
    INJURY_PROPERTY_PERSONAL = 'property_personal'  # 车损+人身伤害（无伤残）

    def __init__(self):
        # 案件类型
        self.case_type = self.CASE_TYPE_CIVIL
        self.injury_type = self.INJURY_DISABILITY

        # 赔偿标准
        self.standard_year = '2024'
        self.region = 'urban'  # urban / rural

        # 当事人
        self.plaintiffs = []   # [{name, gender, birthdate, ethnicity, work, position, phone, address, residence, id_number, relation}]
        self.defendants_person = []  # [{name, gender, birthdate, address, id_number}]
        self.defendants_company = []  # [{name, address, legal_person, position, phone, credit_code, type}]
        self.agents = []  # [{name, unit, position, phone, authority}]

        # 责任（指被告/侵权方的责任类型）
        self.liability_type = '全责'  # 被告责任：全责/主责/同责/次责/无责（指被告对事故承担的责任比例）
        self.liability_detail = ''    # 责任认定书编号及内容

        # 事故信息
        self.accident_time = ''
        self.accident_location = ''
        self.accident_detail = ''
        self.responsibility_doc_number = ''
        self.responsibility_result = ''

        # 保险信息
        self.insurance_info = ''

        # 医疗信息
        self.hospital_name = ''
        self.hospital_start_date = ''
        self.hospital_end_date = ''
        self.hospital_days = 0
        self.medical_fee = 0

        # 日标准（可自定义，覆盖默认值）
        self.hospital_meal_daily = None   # 默认100元/天，可改50等
        self.nutrition_daily = None       # 默认30元/天，可改20等

        # 三期
        self.nursing_days = 0
        self.nursing_persons = 1
        self.nursing_custom_daily = None
        self.nutrition_days = 0
        self.lost_work_days = 0
        self.lost_work_daily = None

        # 伤残信息
        self.disability_grade = None   # 1-10
        self.plaintiff_age = 0

        # 被扶养人
        self.dependents = []  # [{age, years, share_ratio, relation}]

        # 其他费用
        self.traffic_fee = 0
        self.property_damage = 0
        self.assistive_device_fee = 0
        self.other_fee_desc = ''
        self.other_fee = 0

        # 证据
        self.evidence_list = []

        # 法院
        self.court = ''
        self.filing_date = ''

    def calculate(self):
        """计算全部赔偿项目"""
        calc = CompensationCalculator(self.standard_year, self.region)
        result = {}

        # 1. 医疗费
        result['medical_fee'] = calc.calc_medical_fee(self.medical_fee)

        # 2. 住院伙食补助费
        meal_daily = self.hospital_meal_daily or HEBEI_STANDARD['hospital_meal']
        result['hospital_meal_fee'] = self.hospital_days * meal_daily

        # 3. 营养费
        nutri_daily = self.nutrition_daily or HEBEI_STANDARD['nutrition']
        result['nutrition_fee'] = (self.nutrition_days or self.hospital_days) * nutri_daily

        # 4. 护理费
        result['nursing_fee'] = calc.calc_nursing(
            self.nursing_days or self.hospital_days,
            self.nursing_persons,
            self.nursing_custom_daily
        )

        # 5. 误工费
        result['lost_wage'] = calc.calc_lost_wage(
            self.lost_work_days or self.hospital_days,
            self.lost_work_daily
        )

        # 6. 交通费
        result['traffic_fee'] = calc.calc_traffic_fee(self.traffic_fee)

        # 7-9. 根据伤情类型
        if self.injury_type == self.INJURY_DEATH:
            result['death_compensation'] = calc.calc_death_compensation(self.plaintiff_age)
            result['funeral_fee'] = calc.calc_funeral()
            result['disability_compensation'] = 0
            result['assistive_device_fee'] = 0
        elif self.injury_type == self.INJURY_DISABILITY:
            result['disability_compensation'] = calc.calc_disability_compensation(
                self.disability_grade, self.plaintiff_age
            )
            result['death_compensation'] = 0
            result['funeral_fee'] = 0
            result['assistive_device_fee'] = self.assistive_device_fee
        else:
            # 仅财产损失 或 车损+人身伤害（无伤残）
            result['disability_compensation'] = 0
            result['death_compensation'] = 0
            result['funeral_fee'] = 0
            result['assistive_device_fee'] = 0

        # 被扶养人生活费
        if self.dependents:
            result['dependent_living'] = calc.calc_dependent_living(self.dependents)
        else:
            result['dependent_living'] = 0

        # 10. 精神损害抚慰金
        result['mental_damage_fee'] = calc.calc_mental_damage(
            self.case_type, self.injury_type, self.disability_grade
        )

        # 11. 财产损失
        result['property_damage'] = self.property_damage

        # 12. 其他费用
        result['other_fee'] = self.other_fee

        # 13. 标的总额
        result['total_amount'] = (
            result['medical_fee'] +
            result['hospital_meal_fee'] +
            result['nutrition_fee'] +
            result['nursing_fee'] +
            result['lost_wage'] +
            result['traffic_fee'] +
            result['disability_compensation'] +
            result['death_compensation'] +
            result['funeral_fee'] +
            result['dependent_living'] +
            result['assistive_device_fee'] +
            result['mental_damage_fee'] +
            result['property_damage'] +
            result['other_fee']
        )

        # 保险拆分（按分项限额严格拆解）
        # 医疗费分项：医疗费+住院伙食补助费+营养费
        loss_medical = result['medical_fee'] + result['hospital_meal_fee'] + result['nutrition_fee']
        # 死亡伤残分项：残疾/死亡赔偿金+误工费+护理费+交通费+丧葬费+精神抚慰金+被扶养人生活费+辅助器具费
        loss_death_disability = (
            result['disability_compensation'] + result['death_compensation'] +
            result['lost_wage'] + result['nursing_fee'] + result['traffic_fee'] +
            result['funeral_fee'] + result['mental_damage_fee'] +
            result['dependent_living'] + result['assistive_device_fee']
        )
        # 财产损失分项：财产损失+其他费用（其他费用按财产性质处理）
        loss_property = result['property_damage'] + result['other_fee']

        result['insurance_split'] = calc.calc_insurance_split(
            {'medical': loss_medical, 'death_disability': loss_death_disability, 'property': loss_property},
            self.liability_type
        )

        return result


# ============================================================
# 要素式起诉状生成器
# ============================================================
class LawsuitGenerator:
    """要素式起诉状生成器（区域定位填充版 v3.0）"""

    # 注册命名空间
    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    }
    EXTRA_NS = {
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        'xml': 'http://www.w3.org/XML/1998/namespace',
        'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    }

    def __init__(self, template_path, output_dir='.'):
        self.template_path = template_path
        self.output_dir = output_dir
        self.unpacked_dir = None
        self.root = None
        self._xml_decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

    def generate(self, case: CaseData, output_filename='民事起诉状.docx'):
        """一键生成"""
        self._unpack()
        self._parse_xml()

        # 1. 计算赔偿结果
        result = case.calculate()

        # 2. 按需复制当事人区块（多原告/多被告场景）
        self._duplicate_sections(case)

        # 3. 重新解析（复制后XML已变）
        doc_path = os.path.join(self.unpacked_dir, 'word/document.xml')
        self.root = ET.parse(doc_path).getroot()

        # 4. 构建区域索引
        section_map = self._build_section_map()

        # 5. 构建填充数据
        fill_data = self._build_fill_data(case, result)

        # 6. 区域内字段填充
        for sf in fill_data.get('section_fills', []):
            self._fill_section(sf, section_map)

        # 7. 勾选框处理
        for cb in fill_data.get('checkbox_ops', []):
            self._do_checkbox(cb, section_map)

        # 8. 保存输出
        output = self._save_and_pack(output_filename)
        self._cleanup()
        return output, result

    def _duplicate_sections(self, case: CaseData):
        """
        按需复制当事人区块，使模板能容纳多个原告/被告/代理人
        法〔2025〕82号要素式模板结构：每个当事人占一个表格行(<w:tr>)
        复制策略：找到包含该当事人标记的表格行，在其后插入复制的行
        """
        ns = self.NAMESPACES

        # 计算需要几个区块
        needed = {
            '被告（自然人）': len(case.defendants_person) if case.defendants_person else 1,
            '被告（法人、非法人组织）': len(case.defendants_company) if case.defendants_company else 1,
            '原告（自然人）': len(case.plaintiffs) if case.plaintiffs else 1,
            '委托诉讼代理人': len(case.agents) if case.agents else 1,
        }

        # 找到包含当事人标记的表格行
        def _find_row_containing(text_marker):
            """找到包含指定文本的表格行"""
            for tbl in self.root.findall('.//w:tbl', ns):
                for row in tbl.findall('w:tr', ns):
                    row_text = ''.join([t.text for p in row.findall('.//w:p', ns)
                                       for t in p.findall('.//w:t', ns) if t.text])
                    if text_marker in row_text:
                        return row, tbl
            return None, None

        # 按从后往前的顺序处理（避免行号偏移）
        dup_tasks = []
        for marker, count in needed.items():
            if count > 1:
                dup_tasks.append((marker, count))

        # 倒序处理（后面的行先复制，避免索引偏移）
        dup_tasks.reverse()

        for marker, count in dup_tasks:
            source_row, tbl = _find_row_containing(marker)
            if source_row is None or tbl is None:
                continue

            rows = list(tbl.findall('w:tr', ns))
            source_idx = rows.index(source_row)

            for copy_num in range(1, count):
                new_row = deepcopy(source_row)
                tbl.insert(source_idx + copy_num, new_row)

        # 保存修改后的XML
        doc_path = os.path.join(self.unpacked_dir, 'word/document.xml')
        xml_str = ET.tostring(self.root, encoding='unicode')
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(self._xml_decl + xml_str)

    # ================================================================
    # XML 解析与保存
    # ================================================================
    def _unpack(self):
        self.unpacked_dir = self.template_path.replace('.docx', '_work')
        if os.path.exists(self.unpacked_dir):
            shutil.rmtree(self.unpacked_dir)
        with zipfile.ZipFile(self.template_path, 'r') as zf:
            zf.extractall(self.unpacked_dir)

    def _parse_xml(self):
        for prefix, uri in self.NAMESPACES.items():
            ET.register_namespace(prefix, uri)
        for prefix, uri in self.EXTRA_NS.items():
            ET.register_namespace(prefix, uri)

        doc_path = os.path.join(self.unpacked_dir, 'word/document.xml')
        self.root = ET.parse(doc_path).getroot()

    def _save_and_pack(self, filename):
        doc_path = os.path.join(self.unpacked_dir, 'word/document.xml')
        xml_str = ET.tostring(self.root, encoding='unicode')
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(self._xml_decl + xml_str)

        # 去除最高法模板自带的documentProtection标签（enforcement=0未启用但WPS/Word可能弹保护提示）
        settings_path = os.path.join(self.unpacked_dir, 'word/settings.xml')
        if os.path.exists(settings_path):
            with open(settings_path, 'r', encoding='utf-8') as f:
                settings = f.read()
            settings = re.sub(r'<w:documentProtection[^/]*/>', '', settings)
            with open(settings_path, 'w', encoding='utf-8') as f:
                f.write(settings)

        output_path = os.path.join(self.output_dir, filename)
        if os.path.exists(output_path):
            os.remove(output_path)
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(self.unpacked_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arc_name = os.path.relpath(file_path, self.unpacked_dir)
                    zf.write(file_path, arc_name)
        return output_path

    def _cleanup(self):
        if self.unpacked_dir and os.path.exists(self.unpacked_dir):
            shutil.rmtree(self.unpacked_dir)

    # ================================================================
    # 辅助方法
    # ================================================================
    def _get_para_text(self, p) -> str:
        """获取段落完整文本"""
        ns = self.NAMESPACES
        texts = p.findall('.//w:t', ns)
        return ''.join([t.text for t in texts if t.text])

    # ================================================================
    # 区域索引
    # ================================================================
    def _build_section_map(self) -> dict:
        """
        构建段落区域索引（支持同名多实例）
        返回: {"原告_自然人_1": (start_idx, end_idx), "被告_法人_1": ..., "被告_自然人_2": ..., ...}
        同类型区域按出现顺序编号，单实例也带编号，保证key唯一
        """
        ns = self.NAMESPACES
        paragraphs = list(self.root.findall('.//w:p', ns))

        # 当事人标记（可跟子类型）
        PERSON_MARKERS = {'原告', '被告', '第三人'}
        # 代理人标记
        AGENT_MARKERS = {'委托诉讼代理人'}
        # 独立区域标记
        SECTION_MARKERS = {
            '诉讼请求': '诉讼请求',
            '诉前保全及鉴定申请': '诉前保全及鉴定申请',
            '事实与理由': '事实与理由',
            '对纠纷解决方式的意愿': '对纠纷解决方式的意愿',
        }

        # 第一轮：收集所有区域（可能有同名覆盖）
        raw_sections = []  # [(section_name, start_idx, end_idx), ...]
        current_section = None
        section_start = None

        for i, p in enumerate(paragraphs):
            text = self._get_para_text(p).strip()
            new_section = None

            # 子类型行
            if text == '（自然人）':
                if current_section in PERSON_MARKERS:
                    new_section = f'{current_section}_自然人'
            elif text == '（法人、非法人组织）':
                if current_section in PERSON_MARKERS:
                    new_section = f'{current_section}_法人'
            # 当事人标记
            elif text in PERSON_MARKERS:
                new_section = text
            # 代理人标记
            elif text in AGENT_MARKERS:
                new_section = text
            # 独立区域标记
            else:
                for marker, section_name in SECTION_MARKERS.items():
                    if text == marker or text.startswith(marker):
                        new_section = section_name
                        break
                # 签名区
                if text.startswith('具状人'):
                    new_section = '签名区'

            if new_section:
                # 保存前一个区域
                if current_section and section_start is not None:
                    raw_sections.append((current_section, section_start, i))
                current_section = new_section
                section_start = i

        # 最后一个区域
        if current_section and section_start is not None:
            raw_sections.append((current_section, section_start, len(paragraphs)))

        # 后处理1：无子类型的当事人自动加_自然人
        # 检查：某个当事人标记后面是否紧跟了子类型行
        person_with_subtype = set()
        for idx, (name, _, end) in enumerate(raw_sections):
            if name in PERSON_MARKERS:
                # 检查紧邻的下一个区域是否是该当事人的子类型
                if idx + 1 < len(raw_sections):
                    next_name = raw_sections[idx + 1][0]
                    if next_name.startswith(name + '_'):
                        person_with_subtype.add(idx)

        # 重命名无子类型的当事人标记
        for idx in range(len(raw_sections)):
            name = raw_sections[idx][0]
            if name in PERSON_MARKERS and idx not in person_with_subtype:
                raw_sections[idx] = (f'{name}_自然人', raw_sections[idx][1], raw_sections[idx][2])

        # 后处理2：同名区域按出现顺序编号
        # 同时删除当事人标记自身的行（只有子类型行保留）
        # 即：如果存在"原告_自然人"，则删除纯"原告"行
        has_subtypes = {name.split('_')[0] for name, _, _ in raw_sections
                       if '_' in name and name.split('_')[0] in PERSON_MARKERS}
        filtered_sections = []
        for name, start, end in raw_sections:
            if name in has_subtypes:
                continue  # 跳过纯当事人标记行（已有子类型行替代）
            filtered_sections.append((name, start, end))

        # 编号
        section_map = {}
        counter = {}  # section_name -> 当前编号
        for name, start, end in filtered_sections:
            if name not in counter:
                counter[name] = 1
            numbered_name = f'{name}_{counter[name]}'
            section_map[numbered_name] = (start, end)
            counter[name] += 1

        return section_map

    # ================================================================
    # 填充数据构建
    # ================================================================
    def _build_fill_data(self, case: CaseData, result: dict) -> dict:
        """从CaseData和计算结果生成结构化填充数据"""
        section_fills = []
        checkbox_ops = []

        # === 原告_自然人（支持多个） ===
        for idx, p in enumerate(case.plaintiffs or []):
            fields = {}
            checkboxes = {}

            if p.get('name'):
                fields['姓名：'] = ' ' + p['name']
            if p.get('gender'):
                # 性别互斥勾选：用checkbox_ops精确控制
                if p['gender'] == '男':
                    checkbox_ops.append({
                        'section': f'原告_自然人_{idx + 1}',
                        'paragraph_contains': '性别',
                        'before_checkbox': '男',
                        'check': True,
                    })
                elif p['gender'] == '女':
                    checkbox_ops.append({
                        'section': f'原告_自然人_{idx + 1}',
                        'paragraph_contains': '性别',
                        'before_checkbox': '女',
                        'check': True,
                    })
            if p.get('birthdate'):
                fields['出生日期：'] = ' ' + p['birthdate']
            if p.get('ethnicity'):
                fields['民族：'] = p['ethnicity']
            if p.get('work'):
                fields['工作单位：'] = p['work']
            if p.get('position'):
                fields['职务：'] = p['position']
            if p.get('phone'):
                fields['联系电话：'] = ' ' + p['phone']
            if p.get('address'):
                fields['住所地（户籍所在地）：'] = ' ' + p['address']
            if p.get('residence'):
                fields['经常居住地：'] = p['residence']
            if p.get('id_number'):
                fields['证件号码：'] = ' ' + p['id_number']

            section_fills.append({
                'section': f'原告_自然人_{idx + 1}',
                'fields': fields,
                'checkboxes': checkboxes,
            })

        # === 被告_自然人（支持多个） ===
        for idx, d in enumerate(case.defendants_person or []):
            fields = {}
            checkboxes = {}

            if d.get('name'):
                fields['姓名：'] = ' ' + d['name']
            if d.get('gender'):
                if d['gender'] == '男':
                    checkbox_ops.append({
                        'section': f'被告_自然人_{idx + 1}',
                        'paragraph_contains': '性别',
                        'before_checkbox': '男',
                        'check': True,
                    })
                elif d['gender'] == '女':
                    checkbox_ops.append({
                        'section': f'被告_自然人_{idx + 1}',
                        'paragraph_contains': '性别',
                        'before_checkbox': '女',
                        'check': True,
                    })
            if d.get('birthdate'):
                fields['出生日期：'] = d['birthdate']
            if d.get('address'):
                fields['住所地（户籍所在地）：'] = ' ' + d['address']
            if d.get('id_number'):
                fields['证件号码：'] = ' ' + d['id_number']
            if d.get('phone'):
                fields['联系电话：'] = ' ' + d['phone']

            section_fills.append({
                'section': f'被告_自然人_{idx + 1}',
                'fields': fields,
                'checkboxes': checkboxes,
            })

        # === 被告_法人（支持多个） ===
        for idx, c in enumerate(case.defendants_company or []):
            fields = {}
            checkboxes = {}

            if c.get('name'):
                fields['名称：'] = ' ' + c['name']
            if c.get('address'):
                fields['住所地（主要办事机构所在地）：'] = ' ' + c['address']
            if c.get('legal_person'):
                fields['法定代表人 / 负责人：'] = c['legal_person']
            if c.get('position'):
                fields['职务：'] = c['position']
            if c.get('phone'):
                fields['联系电话：'] = c['phone']
            if c.get('credit_code'):
                fields['统一社会信用代码：'] = ' ' + c['credit_code']
            # 公司类型勾选
            company_type = c.get('type', '')
            if company_type:
                type_map = {
                    '有限责任公司': '有限责任公司',
                    '股份有限公司': '股份有限公司',
                    '上市公司': '上市公司',
                }
                if company_type in type_map:
                    checkboxes[type_map[company_type]] = True
            # 所有制性质
            ownership = c.get('ownership', '')
            if ownership:
                own_map = {'国有': '国有', '民营': '民营'}
                if ownership in own_map:
                    checkboxes[own_map[ownership]] = True

            section_fills.append({
                'section': f'被告_法人_{idx + 1}',
                'fields': fields,
                'checkboxes': checkboxes,
            })

        # === 委托诉讼代理人（支持多个） ===
        if case.agents:
            for idx, a in enumerate(case.agents):
                agent_fields = {}
                agent_checkboxes = {'有': True}

                if a.get('name'):
                    agent_fields['姓名：'] = a['name']
                if a.get('unit'):
                    agent_fields['单位：'] = a['unit']
                if a.get('position'):
                    agent_fields['职务：'] = a['position']
                if a.get('phone'):
                    agent_fields['联系电话：'] = a['phone']
                authority = a.get('authority', '')
                if authority == '特别授权':
                    agent_checkboxes['特别授权'] = True
                elif authority == '一般授权':
                    agent_checkboxes['一般授权'] = True

                section_fills.append({
                    'section': f'委托诉讼代理人_{idx + 1}',
                    'fields': agent_fields,
                    'checkboxes': agent_checkboxes,
                })
        else:
            # 无代理人
            section_fills.append({
                'section': '委托诉讼代理人_1',
                'fields': {},
                'checkboxes': {'有': False, '无': True},
            })

        # === 诉讼请求区 ===
        request_fields = {}
        request_checkboxes = {}

        # 1. 医疗费
        if result.get('medical_fee', 0) > 0:
            request_fields['累计发生医疗费'] = f' {result["medical_fee"]:.2f} 元'
            if case.hospital_name:
                request_fields['医院住院（门'] = f' {case.hospital_name}'
            request_checkboxes['医疗费发票、医疗费清单、病历资料：有'] = True

        # 2. 护理费
        if result.get('nursing_fee', 0) > 0:
            nursing_days = case.nursing_days or case.hospital_days
            request_fields['住院护理'] = f' {nursing_days} 天'
            request_fields['支付护理费'] = f' {result["nursing_fee"]:.2f} 元'
            request_checkboxes['住院证明、医嘱等：有'] = True

        # 3. 营养费
        if result.get('nutrition_fee', 0) > 0:
            request_fields['营养费'] = f' {result["nutrition_fee"]:.2f} 元'
            request_checkboxes['病历资料：有'] = (result['nutrition_fee'] > 0)

        # 4. 住院伙食补助费
        if result.get('hospital_meal_fee', 0) > 0:
            request_fields['住院伙食补助费'] = f' {result["hospital_meal_fee"]:.2f} 元'

        # 5. 误工费
        if result.get('lost_wage', 0) > 0:
            request_fields['误工费'] = f' {result["lost_wage"]:.2f} 元'

        # 6. 交通费
        if result.get('traffic_fee', 0) > 0:
            request_fields['交通费'] = f' {result["traffic_fee"]:.2f} 元'
            request_checkboxes['交通费凭证：有'] = True

        # 7. 残疾赔偿金
        if case.injury_type == CaseData.INJURY_DISABILITY and result.get('disability_compensation', 0) > 0:
            request_fields['残疾赔偿金'] = f' {result["disability_compensation"]:.2f} 元'
            if result.get('dependent_living', 0) > 0:
                request_fields['被扶养人生活费'] = f' {result["dependent_living"]:.2f} 元'

        # 8. 残疾辅助器具费
        if result.get('assistive_device_fee', 0) > 0:
            request_fields['残疾辅助器具费'] = f' {result["assistive_device_fee"]:.2f} 元'

        # 9. 死亡赔偿金、丧葬费
        if case.injury_type == CaseData.INJURY_DEATH:
            if result.get('death_compensation', 0) > 0:
                request_fields['死亡赔偿金'] = f' {result["death_compensation"]:.2f} 元'
            if result.get('funeral_fee', 0) > 0:
                request_fields['丧葬费'] = f' {result["funeral_fee"]:.2f} 元'

        # 10. 精神损害抚慰金
        if result.get('mental_damage_fee', 0) > 0:
            request_fields['精神损害抚慰金'] = f' {result["mental_damage_fee"]:.2f} 元'

        # 11. 财产损失
        if result.get('property_damage', 0) > 0:
            request_fields['车辆损失：'] = f' {result["property_damage"]:.2f} 元'
        if case.other_fee_desc and result.get('other_fee', 0) > 0:
            request_fields['其他损失：'] = f' {case.other_fee_desc}{result["other_fee"]:.2f} 元'

        # 13. 标的总额
        request_fields['标的总额'] = f' {result["total_amount"]:.2f} 元'

        section_fills.append({
            'section': '诉讼请求_1',
            'fields': request_fields,
            'checkboxes': request_checkboxes,
        })

        # === 诉前保全及鉴定申请区 ===
        checkbox_ops.append({
            'section': '诉前保全及鉴定申请_1',
            'paragraph_contains': '诉前保全',
            'before_checkbox': '否',
            'check': True,
        })
        checkbox_ops.append({
            'section': '诉前保全及鉴定申请_1',
            'paragraph_contains': '鉴定',
            'before_checkbox': '是',
            'check': case.injury_type in (CaseData.INJURY_DISABILITY, CaseData.INJURY_DEATH),
        })

        # === 事实与理由区 ===
        fact_fields = {}
        if case.accident_detail:
            fact_fields['1. 交通事故发生情况'] = case.accident_detail
        if case.responsibility_result:
            fact_fields['2. 交通事故责任认定'] = case.responsibility_result
        if case.insurance_info:
            fact_fields['3. 机动车投保情况'] = case.insurance_info
        # 请求依据：自动生成
        liability_pct = HEBEI_STANDARD['liability_ratio'].get(case.liability_type, 1.0) * 100
        request_basis = f'被告{case.liability_type}（{liability_pct:.0f}%），应当承担赔偿责任。'
        if case.defendants_company:
            request_basis += '保险公司应在交强险和商业三者险限额内承担赔偿责任。'
        fact_fields['4. 请求依据'] = request_basis

        # 证据清单
        if case.evidence_list:
            evidence_items = [f'{i+1}. {item}' for i, item in enumerate(case.evidence_list)]
            fact_fields['5. 证据清单'] = '；'.join(evidence_items) + '。'

        section_fills.append({
            'section': '事实与理由_1',
            'fields': fact_fields,
            'checkboxes': {},
        })

        # === 签名区 ===
        signature_fields = {}
        if case.plaintiffs:
            signature_fields['具状人（签字、盖章）：'] = ' ' + case.plaintiffs[0].get('name', '')
        filing_date = case.filing_date or datetime.now().strftime('%Y 年 %m 月 %d 日')
        signature_fields['日期：'] = filing_date

        section_fills.append({
            'section': '签名区_1',
            'fields': signature_fields,
            'checkboxes': {},
        })

        # === 对纠纷解决方式的意愿区 ===
        checkbox_ops.append({
            'section': '对纠纷解决方式的意愿_1',
            'paragraph_contains': '是否了解调解作为非诉',
            'before_checkbox': '了解',
            'check': True,
        })
        checkbox_ops.append({
            'section': '对纠纷解决方式的意愿_1',
            'paragraph_contains': '是否考虑先行调解',
            'before_checkbox': '是',
            'check': True,
        })

        return {
            'section_fills': section_fills,
            'checkbox_ops': checkbox_ops,
        }

    # ================================================================
    # 区域内字段填充
    # ================================================================
    def _fill_section(self, sf: dict, section_map: dict):
        """在指定区域内填充字段"""
        section_name = sf.get('section', '')
        fields = sf.get('fields', {})
        checkboxes = sf.get('checkboxes', {})

        ns = self.NAMESPACES
        paragraphs = list(self.root.findall('.//w:p', ns))

        # 查找区域
        region = section_map.get(section_name)
        if not region:
            for key in section_map:
                if key.startswith(section_name + '_') or section_name.startswith(key + '_'):
                    region = section_map[key]
                    break
                if section_name in key or key in section_name:
                    region = section_map[key]
                    break
        if not region:
            return

        start_idx, end_idx = region
        region_paragraphs = paragraphs[start_idx:end_idx + 1]

        # 填充字段
        for label, value in fields.items():
            if not value:
                continue
            self._fill_field_in_region(region_paragraphs, label, str(value))

        # 处理勾选框
        for context, should_check in checkboxes.items():
            self._fill_checkbox_in_region(region_paragraphs, context, should_check)

    def _fill_field_in_region(self, paragraphs: list, label: str, value: str):
        """
        在区域内段落中填充字段（空白模板版）
        
        空白模板特点：
        - 简单字段：标签后直接为空（如"姓名："）
        - 金额字段：标签后有空白位+单位（如"营养费                元"）
        - 日期字段：标签后有"年月日"模板结构（如"出生日期：      年  月  日  民族："）
        
        填充策略：
        1. 找到包含标签的段落
        2. 在标签文本末尾追加值
        3. 清除标签后到下一个"字段标签"（冒号结尾短文本）之前的所有内容
        """
        ns = self.NAMESPACES

        # 找最佳匹配段落：标签在段首的优先，跳过纯标题段
        matched = []
        for p in paragraphs:
            full_text = self._get_para_text(p)
            if label not in full_text:
                continue
            label_pos = full_text.index(label)
            is_heading = full_text.strip().startswith(
                ('1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.', '10.', '11.', '12.', '13.'))
            # 标题段处理：如果后续段落也包含相同标签，跳过标题段
            # （如"3. 营养费"后有"营养费                元"，标题段跳过）
            # 如果标题段是唯一包含该标签的段落（如"13. 标的总额"），则保留
            if is_heading:
                # 检查是否有非标题段也包含此标签
                has_data_para = False
                for p2 in paragraphs:
                    ft2 = self._get_para_text(p2)
                    if label in ft2 and p2 is not p:
                        is_heading2 = ft2.strip().startswith(
                            ('1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.', '10.', '11.', '12.', '13.'))
                        if not is_heading2:
                            has_data_para = True
                            break
                if has_data_para:
                    continue  # 有独立数据段，跳过标题段
            priority = 0 if label_pos <= 5 else 50 + label_pos
            matched.append((priority, p, full_text))

        if not matched:
            return

        matched.sort(key=lambda x: x[0])
        _, p, full_text = matched[0]
        
        # === 特殊处理：标题+空段落模式 ===
        # 如"1. 交通事故发生情况"标题后跟空段落，应填入空段落而非标题
        p_idx = None
        for pi, pp in enumerate(paragraphs):
            if pp is p:
                p_idx = pi
                break
        
        if p_idx is not None and p_idx + 1 < len(paragraphs):
            next_p = paragraphs[p_idx + 1]
            next_text = self._get_para_text(next_p).strip()
            if not next_text:
                # 下一段是空段落，将内容填入空段落
                next_t_elems = list(next_p.findall('.//w:t', ns))
                if next_t_elems:
                    next_t_elems[0].text = value
                    next_t_elems[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    for t in next_t_elems[1:]:
                        if not (t.text and t.text.strip()):
                            t.text = ''
                else:
                    # 空段落没有<w:t>，创建一个
                    from xml.etree.ElementTree import SubElement as _Sub
                    r = _Sub(next_p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                    t_new = _Sub(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                    t_new.text = value
                    t_new.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                return  # 填入空段落后直接返回

        t_elems = list(p.findall('.//w:t', ns))

        # 构建字符偏移映射
        offsets = []
        pos = 0
        for t in t_elems:
            text = t.text or ''
            offsets.append((t, pos, pos + len(text)))
            pos += len(text)

        # 定位标签结束位置
        label_start = full_text.index(label)
        label_end = label_start + len(label)

        # 定位清除边界：下一个字段标签（以中文冒号结尾的2-6字短文本）
        clear_end = len(full_text)  # 默认清除到段落末尾
        for m in re.finditer(r'[一-鿿/]{2,6}[：:]', full_text[label_end:]):
            clear_end = label_end + m.start()
            break

        # 找到包含标签结束位置的<w:t>
        label_end_t = None
        label_end_pos_in_t = None
        for t, start, end in offsets:
            if start <= label_end <= end:
                label_end_t = t
                label_end_pos_in_t = label_end - start
                break

        if label_end_t is None:
            return

        # 在标签结束位置截断，追加新值
        # 如果清除边界后有保留内容，值末尾加空格分隔
        display_value = value
        if clear_end < len(full_text):
            display_value = value.rstrip() + ' '
        label_end_t.text = label_end_t.text[:label_end_pos_in_t] + display_value
        label_end_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

        # 清除标签<w:t>之后到clear_end之间的所有<w:t>内容
        found_label_t = False
        for t, start, end in offsets:
            if t is label_end_t:
                found_label_t = True
                continue
            if not found_label_t:
                continue
            if start >= clear_end:
                break  # 超出清除范围
            if end <= clear_end:
                # 完全在清除范围内
                t.text = ''
            else:
                # 跨越清除边界，保留边界之后的部分
                keep_start = clear_end - start
                if t.text and len(t.text) > keep_start:
                    t.text = t.text[keep_start:]
                    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                else:
                    t.text = ''


    def _fill_checkbox_in_region(self, paragraphs: list, context: str, should_check: bool):
        """在区域内处理勾选框"""
        ns = self.NAMESPACES

        # 收集匹配段落，优先匹配context在段首的（避免子串误匹配）
        matched = []
        for p in paragraphs:
            full_text = self._get_para_text(p)
            if context not in full_text:
                continue
            # 判断context是否在段首（更精确匹配）
            at_start = full_text.strip().startswith(context)
            matched.append((0 if at_start else 1, p, full_text))

        if not matched:
            return

        matched.sort(key=lambda x: x[0])

        for _, p, full_text in matched:

            t_elems = p.findall('.//w:t', ns)

            # 同一个<w:t>中 context + □
            for t in t_elems:
                if not t.text:
                    continue
                if context in t.text:
                    after = t.text[t.text.index(context) + len(context):]
                    if '□' in after or '☐' in t.text:
                        if should_check:
                            ctx_end = t.text.index(context) + len(context)
                            before = t.text[:ctx_end]
                            after_part = t.text[ctx_end:]
                            after_part = after_part.replace('□', '☑', 1).replace('☐', '☑', 1)
                            t.text = before + after_part
                        return

            # context在一个<w:t>，□在后续<w:t>
            context_found = False
            for t in t_elems:
                if not context_found:
                    if t.text and context in t.text:
                        context_found = True
                    continue
                # context已找到，找后续的□
                if t.text and ('□' in t.text or '☐' in t.text):
                    if should_check:
                        t.text = t.text.replace('□', '☑', 1).replace('☐', '☑', 1)
                    return

    # ================================================================
    # 精确勾选框操作
    # ================================================================
    def _do_checkbox(self, cb: dict, section_map: dict):
        """
        精确勾选框操作
        cb: {"section": "原告_自然人", "paragraph_contains": "性别", "before_checkbox": "男", "check": True}
        
        para_contains: 上下文线索，在区域内搜索时用来缩小范围。
        如果勾选框所在段落不含para_contains，会向前查找附近段落是否包含。
        """
        section_name = cb.get('section', '')
        para_contains = cb.get('paragraph_contains', '')
        before_checkbox = cb.get('before_checkbox', '')
        should_check = cb.get('check', True)

        ns = self.NAMESPACES
        paragraphs = list(self.root.findall('.//w:p', ns))

        # 确定搜索范围
        search_paragraphs = paragraphs
        if section_name:
            region = section_map.get(section_name)
            if region:
                start_idx, end_idx = region
                search_paragraphs = paragraphs[start_idx:end_idx + 1]

        # 策略：在区域内找 before_checkbox+□ 的段落
        # para_contains作为上下文线索，在目标段落或其附近段落中出现即可
        for idx, p in enumerate(search_paragraphs):
            full_text = self._get_para_text(p)
            
            # 必须包含 before_checkbox + □/☐
            target_text = before_checkbox + '□'
            target_text2 = before_checkbox + '☐'
            has_target = (target_text in full_text or target_text2 in full_text)
            
            # 或者 □ 紧跟在 before_checkbox 后面（跨<w:t>情况）
            has_target_cross = False
            if not has_target:
                t_elems = p.findall('.//w:t', ns)
                prev_text = ''
                for t in t_elems:
                    if t.text and ('□' in t.text or '☐' in t.text):
                        combined = prev_text
                        if combined.rstrip().endswith(before_checkbox):
                            has_target_cross = True
                            break
                    if t.text:
                        prev_text += t.text
            
            if not has_target and not has_target_cross:
                continue
            
            # 如果指定了para_contains，验证上下文（当前段落或前3个段落中包含）
            if para_contains:
                context_found = para_contains in full_text
                if not context_found:
                    # 检查前面的段落
                    for prev_idx in range(max(0, idx - 1), max(0, idx - 4), -1):
                        prev_text = self._get_para_text(search_paragraphs[prev_idx])
                        if para_contains in prev_text:
                            context_found = True
                            break
                if not context_found:
                    continue

            t_elems = p.findall('.//w:t', ns)

            # 策略1：before_checkbox + □/☐ 在同一个<w:t>
            target1 = before_checkbox + '□'
            target2 = before_checkbox + '☐'

            for t in t_elems:
                if not t.text:
                    continue
                if target1 in t.text:
                    if should_check:
                        t.text = t.text.replace(target1, before_checkbox + '☑', 1)
                    return
                if target2 in t.text:
                    if should_check:
                        t.text = t.text.replace(target2, before_checkbox + '☑', 1)
                    return

            # 策略2：before_checkbox在一个<w:t>末尾，□在下一个<w:t>
            prev_text = ''
            for t in t_elems:
                if t.text and ('□' in t.text or '☐' in t.text):
                    # 找到□的位置
                    box_pos = t.text.find('□') if '□' in t.text else t.text.find('☐')
                    combined = prev_text + t.text[:box_pos] if box_pos >= 0 else prev_text
                    if combined.rstrip().endswith(before_checkbox):
                        if should_check:
                            t.text = t.text.replace('□', '☑', 1).replace('☐', '☑', 1)
                        return
                if t.text:
                    prev_text += t.text




# ============================================================
# 赔偿明细报告生成
# ============================================================
def generate_compensation_report(case: CaseData, result: dict) -> str:
    """生成赔偿明细Markdown报告"""

    lines = [
        '# 交通事故赔偿明细计算报告',
        '',
        f'**赔偿标准**：河北省{case.standard_year}年标准（{"城镇" if case.region == "urban" else "农村"}）',
        f'**案件类型**：{"刑事附带民事" if case.case_type == CaseData.CASE_TYPE_CRIMINAL_ATTACHED else "纯民事"}',
        f'**伤情类型**：{"死亡" if case.injury_type == CaseData.INJURY_DEATH else "伤残" if case.injury_type == CaseData.INJURY_DISABILITY else "车损+人身伤害" if case.injury_type == CaseData.INJURY_PROPERTY_PERSONAL else "仅财产损失"}',
        f'**责任比例**：被告{case.liability_type}（{HEBEI_STANDARD["liability_ratio"].get(case.liability_type, 1.0)*100:.0f}%）',
        '',
        '## 赔偿项目明细',
        '',
        '| 序号 | 项目 | 金额(元) | 计算方式 |',
        '|------|------|----------|----------|',
    ]

    items = [
        ('1', '医疗费', result['medical_fee'], '按实际票据'),
        ('2', '护理费', result['nursing_fee'], f'{case.nursing_days or case.hospital_days}天×{case.nursing_persons}人×{case.nursing_custom_daily or HEBEI_STANDARD["nursing"]}元/天'),
        ('3', '营养费', result['nutrition_fee'], f'{case.nutrition_days or case.hospital_days}天×{case.nutrition_daily or HEBEI_STANDARD["nutrition"]}元/天'),
        ('4', '住院伙食补助费', result['hospital_meal_fee'], f'{case.hospital_days}天×{HEBEI_STANDARD["hospital_meal"]}元/天'),
        ('5', '误工费', result['lost_wage'], f'{case.lost_work_days or case.hospital_days}天×{case.lost_work_daily or HEBEI_STANDARD["default_wage"]}元/天'),
        ('6', '交通费', result['traffic_fee'], '按实际票据'),
    ]

    if case.injury_type == CaseData.INJURY_DEATH:
        items.append(('7', '死亡赔偿金', result['death_compensation'], f'{HEBEI_STANDARD[case.standard_year]["urban_income" if case.region == "urban" else "rural_income"]}元/年×20年'))
        items.append(('7-1', '丧葬费', result['funeral_fee'], f'全省在岗职工半年平均工资'))
    elif case.injury_type == CaseData.INJURY_DISABILITY:
        items.append(('7', '残疾赔偿金', result['disability_compensation'], f'{case.disability_grade}级伤残×{HEBEI_STANDARD[case.standard_year]["urban_income" if case.region == "urban" else "rural_income"]}元/年×20年×{HEBEI_STANDARD["disability_ratio"][case.disability_grade]*100:.0f}%'))

    if result['dependent_living'] > 0:
        items.append(('7-2', '被扶养人生活费', result['dependent_living'], '分段计算'))

    if result['mental_damage_fee'] > 0:
        items.append(('8', '精神损害抚慰金', result['mental_damage_fee'], f'{"死亡" if case.injury_type == "death" else f"{case.disability_grade}级伤残"}标准'))
    elif case.case_type == CaseData.CASE_TYPE_CRIMINAL_ATTACHED:
        items.append(('8', '精神损害抚慰金', 0, '❌ 刑事附带民事不支持'))

    if result['property_damage'] > 0:
        items.append(('9', '财产损失', result['property_damage'], '车辆损失等'))

    items.append(('10', '其他费用', result['other_fee'], case.other_fee_desc or '鉴定费、后续治疗费等'))

    for seq, name, amount, calc_method in items:
        lines.append(f'| {seq} | {name} | {amount:,.2f} | {calc_method} |')

    lines.append(f'| **合计** | **标的总额** | **{result["total_amount"]:,.2f}** | |')

    # 保险拆分
    ins = result['insurance_split']
    lines.extend([
        '',
        '## 保险赔付拆分',
        '',
        '| 项目 | 金额(元) |',
        '|------|----------|',
        f'| 交强险合计 | {ins["compulsory"]:,.2f} |',
    ])
    # 分项明细
    if 'compulsory_detail' in ins:
        cd = ins['compulsory_detail']
        if cd['medical'] > 0:
            lines.append(f'| └ 医疗费用限额（18,000） | {cd["medical"]:,.2f} |')
        if cd['death_disability'] > 0:
            lines.append(f'| └ 死亡伤残限额（180,000） | {cd["death_disability"]:,.2f} |')
        if cd['property'] > 0:
            lines.append(f'| └ 财产损失限额（2,000） | {cd["property"]:,.2f} |')
    lines.extend([
        f'| 商业三者险（被告{case.liability_type}{HEBEI_STANDARD["liability_ratio"].get(case.liability_type, 1.0)*100:.0f}%） | {ins["commercial"]:,.2f} |',
        f'| 原告自行承担（{(1-HEBEI_STANDARD["liability_ratio"].get(case.liability_type, 1.0))*100:.0f}%） | {ins["self_bear"]:,.2f} |',
        f'| **保险公司合计** | **{ins["total_insurance"]:,.2f}** |',
    ])

    # 重要提示
    lines.extend([
        '',
        '## ⚠️ 重要提示',
    ])

    if case.case_type == CaseData.CASE_TYPE_CRIMINAL_ATTACHED:
        lines.append('- 刑事附带民事案件不支持精神损害抚慰金（刑诉法解释第192条）')
        lines.append('- 处理丧葬事宜误工费已不被支持（民法典已删除该项目）')

    if case.injury_type == CaseData.INJURY_DEATH:
        lines.append('- 死亡赔偿金按20年计算，60岁以上每增1岁减1年，最低5年')

    if case.liability_type != '全责':
        lines.append(f'- 被告负{case.liability_type}，商业险部分按被告{HEBEI_STANDARD["liability_ratio"].get(case.liability_type, 1.0)*100:.0f}%比例赔付')

    return '\n'.join(lines)


# ============================================================
# 证据清单生成
# ============================================================
def generate_evidence_list(case: CaseData) -> str:
    """根据案件类型自动生成证据清单（调用统一规则库）"""
    from evidence_list_generator import generate_evidence_list_text
    case_dict = _case_to_dict(case)
    return generate_evidence_list_text(case_dict)


def _case_to_dict(case: CaseData) -> dict:
    """将CaseData转为证据规则库所需的字典格式"""
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
    }


# ============================================================
# 测试
# ============================================================
if __name__ == '__main__':
    # === 测试1：死亡案件（冉某江案） ===
    print("=" * 60)
    print("测试1：冉某江等交通肇事案（刑事附带民事·死亡案件）")
    print("=" * 60)

    case1 = CaseData()
    case1.case_type = CaseData.CASE_TYPE_CRIMINAL_ATTACHED
    case1.injury_type = CaseData.INJURY_DEATH
    case1.standard_year = '2023'
    case1.liability_type = '主责'

    case1.plaintiffs = [
        {'name': '冉某江', 'phone': '13831910001', 'id_number': '130531197006150011',
         'birthdate': '1970 年 6 月 15 日', 'address': '河北省邢台市广宗县塘疃乡常阜村'},
    ]

    case1.hospital_days = 11
    case1.medical_fee = 102309.83
    case1.nursing_days = 11
    case1.nursing_custom_daily = 139
    case1.nutrition_days = 11
    case1.lost_work_days = 11
    case1.lost_work_daily = 178
    case1.traffic_fee = 3000
    case1.property_damage = 2000
    case1.plaintiff_age = 55  # 被害人高某芝年龄（推算）

    case1.accident_detail = '2023年10月28日13时50分，被告驾驶冀E1××某某号小型轿车由北向南行驶至广宗县常阜村南十字路口处时，与前方同向行驶向左转弯的被害人高某芝驾驶的电动三轮车相碰撞，致高某芝受伤，后经医院抢救无效于2023年11月8日死亡。'
    case1.responsibility_result = '本次事故经广宗县公安局交通警察大队道路交通事故认定书认定，被告人潘某坤负本起事故的主要责任，被害人高某芝负次要责任。'
    case1.insurance_info = '被告驾驶的冀E1××某某号小型轿车在中国某某财产保险股份有限公司邢台市中心支公司投保机动车交强险和商业第三者责任保险（保险金额200万元）。'
    case1.evidence_list = ['医疗费凭证、病历资料一宗', '交通事故责任认定书', '居民死亡医学证明书、死亡殡葬证', '户口本、法定继承人证明']
    case1.filing_date = '2024 年 2 月 5 日'

    result1 = case1.calculate()
    print(f"\n赔偿计算结果:")
    for k, v in result1.items():
        if isinstance(v, (int, float)):
            print(f"  {k}: {v:,.2f}")
        elif isinstance(v, dict):
            for k2, v2 in v.items():
                print(f"  {k}.{k2}: {v2:,.2f}")

    print(f"\n标的总额: {result1['total_amount']:,.2f}元")
    print(f"精神损害抚慰金: {result1['mental_damage_fee']:,.2f}元（刑事附带民事→0）")
    print(f"保险公司合计: {result1['insurance_split']['total_insurance']:,.2f}元")

    # 生成起诉状
    gen = LawsuitGenerator('交通事故案件工作流/模板_民事起诉状.docx', '交通事故案件工作流')
    output = gen.generate(case1, '要素式起诉状_冉某江案_v2.docx')
    print(f"\n✓ 起诉状已生成: {output}")

    # 生成赔偿明细报告
    report = generate_compensation_report(case1, result1)
    with open('交通事故案件工作流/赔偿明细_冉某江案.md', 'w', encoding='utf-8') as f:
        f.write(report)
    print("✓ 赔偿明细已生成")

    # === 测试2：伤残案件（纯民事） ===
    print("\n" + "=" * 60)
    print("测试2：普通伤残案件（纯民事·10级伤残）")
    print("=" * 60)

    case2 = CaseData()
    case2.case_type = CaseData.CASE_TYPE_CIVIL
    case2.injury_type = CaseData.INJURY_DISABILITY
    case2.standard_year = '2024'
    case2.liability_type = '全责'

    case2.plaintiffs = [
        {'name': '李明', 'phone': '13931912345', 'id_number': '130502198805121234',
         'birthdate': '1988 年 5 月 12 日', 'address': '河北省邢台市桥西区团结路12号'},
    ]

    case2.hospital_days = 45
    case2.medical_fee = 52340.80
    case2.nursing_days = 45
    case2.nursing_persons = 1
    case2.nutrition_days = 60
    case2.lost_work_days = 180
    case2.lost_work_daily = 178
    case2.traffic_fee = 680
    case2.disability_grade = 10
    case2.plaintiff_age = 36

    result2 = case2.calculate()
    print(f"\n标的总额: {result2['total_amount']:,.2f}元")
    print(f"残疾赔偿金: {result2['disability_compensation']:,.2f}元")
    print(f"精神损害抚慰金: {result2['mental_damage_fee']:,.2f}元")

    output2 = gen.generate(case2, '要素式起诉状_李明伤残案.docx')
    print(f"✓ 起诉状已生成: {output2}")
