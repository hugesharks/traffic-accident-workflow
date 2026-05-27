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
    """要素式起诉状生成器"""

    def __init__(self, template_path, output_dir='.'):
        self.template_path = template_path
        self.output_dir = output_dir
        self.unpacked_dir = None

    def generate(self, case: CaseData, output_filename='民事起诉状.docx'):
        """一键生成"""
        self._unpack()
        self._fill(case)
        output = self._pack(output_filename)
        self._cleanup()
        return output

    def _unpack(self):
        self.unpacked_dir = self.template_path.replace('.docx', '_work')
        if os.path.exists(self.unpacked_dir):
            shutil.rmtree(self.unpacked_dir)
        with zipfile.ZipFile(self.template_path, 'r') as zf:
            zf.extractall(self.unpacked_dir)

    def _pack(self, filename):
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

    def _fill(self, case: CaseData):
        """填充模板"""
        result = case.calculate()

        doc_path = os.path.join(self.unpacked_dir, 'word/document.xml')
        with open(doc_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # === 金额替换 ===
        amount_replacements = [
            ('>393155.55<', f'>{result["total_amount"]:.2f}<'),
            ('>167034.15<', f'>{result["medical_fee"]:.2f}<'),
            ('>21367<', f'>{result["nursing_fee"]:.2f}<'),
            ('>4500<', f'>{result["nutrition_fee"]:.2f}<'),
            ('>3900<', f'>{result["hospital_meal_fee"]:.2f}<'),
            ('>30150<', f'>{result["lost_wage"]:.2f}<'),
            ('>1211.5<', f'>{result["traffic_fee"]:.2f}<'),
            ('>139992.9<', f'>{result["disability_compensation"] or result["death_compensation"]:.2f}<'),
            ('>5000<', f'>{result["mental_damage_fee"]:.2f}<'),
            ('>20000<', f'>{result["other_fee"]:.2f}<'),
        ]

        for old, new in amount_replacements:
            content = content.replace(old, new)

        # === 当事人信息替换 ===
        if case.plaintiffs:
            p = case.plaintiffs[0]
            content = content.replace('>荆<', f'>{p["name"][0]}<')
            content = content.replace('> ××<', f'> {p["name"][1:] if len(p["name"]) > 1 else ""}<')
            content = content.replace('>×××××××××××<', f'>{p.get("phone", "")}<')
            content = content.replace('>×××××××××××××××××<', f'>{p.get("id_number", "")}<')
            content = content.replace('>19×× 年  × 月  × 日<', f'>{p.get("birthdate", "")}<')
            content = content.replace('>山东省齐河县 ××× 镇  ×× 村  ××× 号<', f'>{p.get("address", "")}<')
            content = content.replace('>具状人（签字、盖章）： 荆 ××<', f'>具状人（签字、盖章）： {p["name"]}<')

            # 事实与理由中的原告姓名
            content = content.replace(
                '荆 ×× 因与肇事车辆发生交通事故受伤',
                f'{p["name"]}因与肇事车辆发生交通事故受伤'
            )

        # === 日期 ===
        filing_date = case.filing_date or datetime.now().strftime('%Y 年 %m 月 %d 日')
        content = content.replace('>×× 年  ×× 月  ××  日<', f'>{filing_date}<')

        # === 事实与理由 ===
        if case.accident_detail:
            content = content.replace(
                '2023 年 1 月 9 日 19 时 18 分在齐河县  ×× 路  ××× 镇  ×× 路段，被告驾 驶的车牌号为鲁 A××××× 的车辆与原告发生交通事故，导致原告受伤。',
                case.accident_detail
            )

        if case.responsibility_result:
            content = content.replace(
                '本次事故经齐河县公安局交通警察大队出具第 371425××××××××× ××× 号道路交通事故认定书，认定在本次事故中原告无责任、被告负全 部责任。',
                case.responsibility_result
            )

        if case.insurance_info:
            content = content.replace(
                '被告驾驶车牌号为鲁 A××××× 的车辆在被告  ×× 财产保险有限责任公 司济南分公司投保，交强险 20 万元，期限自 2022 年 3 月 1  日起至 2023 年 3 月 1 日止；在  ×××× 保险股份有限公司济南中心支公司投保第三者责 任险 100 万元，期限自 2022 年 3 月 1  日起至 2023 年 3 月 1 日止。',
                case.insurance_info
            )

        # 证据清单
        if case.evidence_list:
            evidence_text = '；'.join([f'{i+1}. {item}' for i, item in enumerate(case.evidence_list)])
            evidence_text += '。'
            content = content.replace(
                '1. 医疗费凭证、病历资料一宗；2. 交通费凭证 5 张；3. 交通事故责任认定 书；4. 鉴定意见。',
                evidence_text
            )

        # === 死亡案件特殊处理 ===
        if case.injury_type == CaseData.INJURY_DEATH:
            # "残疾赔偿金"改为"死亡赔偿金"
            content = content.replace(
                '7. 残疾赔偿金（含被扶养人生活费）',
                '7. 死亡赔偿金、丧葬费'
            )
            # 填入死亡赔偿金+丧葬费
            death_total = result['death_compensation'] + result['funeral_fee']
            content = content.replace(
                f'残疾赔偿金 {result["disability_compensation"] or result["death_compensation"]:.2f} 元',
                f'死亡赔偿金 {result["death_compensation"]:.2f} 元、丧葬费 {result["funeral_fee"]:.2f} 元'
            )

        # === 诉讼请求总额描述 ===
        if case.injury_type == CaseData.INJURY_DEATH:
            content = content.replace(
                '判决被告支付医疗费、护理费、营养费、住院伙食补助费、误工费、交通费、残疾赔偿金、精神损害抚慰金、后续治疗费、鉴定费等共计',
                '判决被告支付医疗费、护理费、营养费、住院伙食补助费、误工费、交通费、死亡赔偿金、丧葬费等共计'
            )
        elif case.injury_type == CaseData.INJURY_DISABILITY:
            content = content.replace(
                '精神损害抚慰金、后续治疗费、鉴定费等共计',
                '精神损害抚慰金、鉴定费等共计'
            )

        # === 勾选框自动填充 ===
        content = self._fill_checkboxes(content, case, result)

        # === 保存 ===
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(content)

        return result

    def _fill_checkboxes(self, content, case, result):
        """
        自动填充勾选框
        核心逻辑：在XML中，☐是独立的<w:t>☐</w:t>元素
        通过前后文本段落内容判断每个☐的归属，然后☐→☑
        """
        plaintiff_gender = case.plaintiffs[0].get('gender', '男') if case.plaintiffs else '男'
        defendants = case.defendants_person

        # 策略：用"段落文本"来定位每个☐，然后替换
        # 先提取所有包含☐的段落及其上下文
        import xml.etree.ElementTree as ET
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

        try:
            root = ET.fromstring(content)
        except:
            return content  # XML解析失败就不处理

        # 构建替换映射：每个☐根据其段落上下文决定是否替换为☑
        replacements_done = 0

        for p in root.findall('.//w:p', ns):
            texts = p.findall('.//w:t', ns)
            full_text = ''.join([t.text for t in texts if t.text])

            if '☐' not in full_text:
                continue

            # 根据段落内容判断每个☐的归属
            for t in texts:
                if t.text != '☐':
                    continue

                should_check = False  # 是否改为☑

                # 获取此☐在段落文本中的位置
                # 通过拼接前序文本判断
                prefix = ''
                for pt in texts:
                    if pt is t:
                        break
                    if pt.text:
                        prefix += pt.text

                # === 判断规则 ===

                # 1. 原告性别：prefix含"性别：男 " → 男☑
                if '性别：男 ' in prefix and prefix.count('性别') == 1:
                    # 判断是第几个"性别：男"
                    # 如果prefix中有"荆"或原告名，说明是原告区域
                    if plaintiff_gender == '男':
                        should_check = True

                # 2. 被告1性别：prefix含"性别：男□    女 " → 女☐→☑
                elif '女 ' in prefix[-20:] and '性别：' in prefix:
                    if defendants and defendants[0].get('gender') == '女':
                        should_check = True

                # 3. 被告2性别：第二个"性别：男 " + 被告2为男 → 男☑
                elif '性别：男 ' in prefix and prefix.count('性别') >= 2:
                    if defendants and len(defendants) > 1 and defendants[1].get('gender') == '男':
                        should_check = True

                # 4. 证据勾选："有 ☐" → 有☑
                elif '有 ' in prefix[-10:]:
                    should_check = True

                # 5. 营养费无证据："无 ☐" → 无☑
                elif '无 ' in prefix[-20:] and '病历资料' in prefix:
                    should_check = True

                # 6. 公司类型：有限责任公司☑
                elif '有限责任公司 ' in prefix[-40:]:
                    should_check = True

                # 7. 所有制：国有☑
                elif '国有 ' in prefix[-10:] and '所有制' in prefix:
                    should_check = True

                # 8. 委托诉讼代理人：无☑
                elif prefix.endswith('无 ') and '代理权限' in prefix:
                    should_check = True

                # 9. 诉前保全：否☑
                elif prefix.endswith('否 '):
                    should_check = True

                # 10. 鉴定申请：是☑
                elif prefix.endswith('是 '):
                    if case.injury_type in (CaseData.INJURY_DISABILITY, CaseData.INJURY_DEATH):
                        should_check = True

                if should_check:
                    t.text = '☑'
                    replacements_done += 1

        # 将修改后的XML写回
        content = ET.tostring(root, encoding='unicode')
        # 重新添加XML声明
        content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + content

        return content


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
