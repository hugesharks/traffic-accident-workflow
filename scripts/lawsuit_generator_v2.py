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
        self.plaintiffs_company = []   # [{name, address, legal_person, position, phone, credit_code, type}]
        self.third_parties_person = [] # [{name, gender, birthdate, address, id_number}]
        self.third_parties_company = [] # [{name, address, legal_person, position, phone, credit_code, type}]
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
    """要素式起诉状生成器（书签定位版 v4.0）
    
    核心设计：用OOXML书签(<w:bookmarkStart>/<w:bookmarkEnd>)作为定位符，
    填充时直接按书签名定位插值，不再依赖文本匹配。
    """
    
    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    }
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    
    def __init__(self, template_path, output_dir='.'):
        self.template_path = template_path
        self.output_dir = output_dir
        self.unpacked_dir = None
        self.root = None
        self._xml_decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    
    # ================================================================
    # 主流程
    # ================================================================
    def generate(self, case, output_filename='民事起诉状.docx'):
        import xml.etree.ElementTree as ET, os
        self._unpack()
        self._parse_xml()
        
        # 1. 添加书签
        self._add_bookmarks()
        
        # 2. 计算
        result = case.calculate()
        
        # 3. 复制当事人行
        self._duplicate_party_rows(case)
        
        # 4. 重新解析
        doc_path = os.path.join(self.unpacked_dir, 'word/document.xml')
        self.root = ET.parse(doc_path).getroot()
        
        # 5. 构建填充数据
        fill_map = self._build_fill_map(case, result)
        
        # 6. 填充
        self._fill_by_bookmarks(fill_map)
        
        # 7. 保存
        output = self._save_and_pack(output_filename)
        self._cleanup()
        return output, result
    
    # ================================================================
    # 模板操作基础
    # ================================================================
    def _unpack(self):
        import zipfile, tempfile, shutil
        if self.unpacked_dir and os.path.exists(self.unpacked_dir):
            shutil.rmtree(self.unpacked_dir)
        self.unpacked_dir = tempfile.mkdtemp(prefix='lawsuit_')
        with zipfile.ZipFile(self.template_path) as zf:
            zf.extractall(self.unpacked_dir)
    
    def _parse_xml(self):
        import xml.etree.ElementTree as ET, os
        doc_path = os.path.join(self.unpacked_dir, 'word/document.xml')
        self.root = ET.parse(doc_path).getroot()
        self._remove_protection()
    
    def _remove_protection(self):
        import os, xml.etree.ElementTree as ET
        settings_path = os.path.join(self.unpacked_dir, 'word/settings.xml')
        if os.path.exists(settings_path):
            tree = ET.parse(settings_path)
            sroot = tree.getroot()
            for dp in sroot.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}documentProtection'):
                sroot.remove(dp)
            tree.write(settings_path, xml_declaration=False, encoding='UTF-8')
    
    def _save_and_pack(self, filename):
        import xml.etree.ElementTree as ET, os, zipfile
        output_path = os.path.join(self.output_dir, filename)
        os.makedirs(self.output_dir, exist_ok=True)
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root_dir, dirs, files in os.walk(self.unpacked_dir):
                for f in files:
                    full = os.path.join(root_dir, f)
                    arc = os.path.relpath(full, self.unpacked_dir)
                    zf.write(full, arc)
        return output_path
    
    def _write_xml(self, path, root_elem):
        import xml.etree.ElementTree as ET
        for prefix, uri in self.NAMESPACES.items():
            ET.register_namespace(prefix, uri)
        ET.register_namespace('', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        tree = ET.ElementTree(root_elem)
        with open(path, 'wb') as f:
            tree.write(f, xml_declaration=False, encoding='UTF-8')
    
    def _cleanup(self):
        import shutil
        if self.unpacked_dir and os.path.exists(self.unpacked_dir):
            shutil.rmtree(self.unpacked_dir)
    
    def _get_para_text(self, p):
        ns = self.NAMESPACES
        return ''.join(t.text or '' for t in p.findall('.//w:t', ns))
    
    # ================================================================
    # XML辅助：run拆分
    # ================================================================
    def _split_run_at_text(self, para, run, split_text):
        """将一个<w:r>按指定文本位置拆分成两个run"""
        import copy, xml.etree.ElementTree as ET
        ns = self.NAMESPACES
        W = self.W
        t = run.find('w:t', ns)
        if t is None or not t.text or split_text not in t.text:
            return None
        text = t.text
        idx = text.index(split_text)
        before = text[:idx]
        after = text[idx:]
        t.text = before
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        new_run = copy.deepcopy(run)
        new_t = new_run.find('w:t', ns)
        new_t.text = after
        new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        run_idx = list(para).index(run)
        para.insert(run_idx + 1, new_run)
        return new_run
    
    # ================================================================
    # 书签操作
    # ================================================================
    def _next_bm_id(self):
        W = self.W
        max_id = 0
        for bm in self.root.iter(f'{{{W}}}bookmarkStart'):
            bid = int(bm.get('id', '0'))
            max_id = max(max_id, bid)
        return max_id + 1
    
    def _add_bm_after_run(self, para, run_idx, bm_name):
        """在段落中指定run之后插入空书签（用于文本插入点）"""
        import xml.etree.ElementTree as ET
        W = self.W
        bm_id = self._next_bm_id()
        children = list(para)
        # 找到第run_idx个<w:r>的位置
        r_count = -1
        insert_pos = None
        for i, child in enumerate(children):
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'r':
                r_count += 1
                if r_count == run_idx:
                    insert_pos = i + 1
                    break
        if insert_pos is None:
            insert_pos = len(children)
        bm_start = ET.Element(f'{{{W}}}bookmarkStart')
        bm_start.set('id', str(bm_id))
        bm_start.set('name', bm_name)
        bm_end = ET.Element(f'{{{W}}}bookmarkEnd')
        bm_end.set('id', str(bm_id))
        para.insert(insert_pos, bm_start)
        para.insert(insert_pos + 1, bm_end)
        return bm_id
    
    def _add_bm_wrap_runs(self, para, start_run, end_run, bm_name):
        """用书签包裹指定范围的run（用于替换内容，如'营养费 [空白] 元'中的空白）"""
        import xml.etree.ElementTree as ET
        W = self.W
        bm_id = self._next_bm_id()
        children = list(para)
        r_count = -1
        start_pos = None
        end_pos = None
        for i, child in enumerate(children):
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'r':
                r_count += 1
                if r_count == start_run and start_pos is None:
                    start_pos = i
                if r_count == end_run:
                    end_pos = i + 1
        if start_pos is None or end_pos is None:
            return None
        bm_start = ET.Element(f'{{{W}}}bookmarkStart')
        bm_start.set('id', str(bm_id))
        bm_start.set('name', bm_name)
        bm_end = ET.Element(f'{{{W}}}bookmarkEnd')
        bm_end.set('id', str(bm_id))
        para.insert(start_pos, bm_start)
        para.insert(end_pos + 1, bm_end)
        return bm_id
    
    def _add_bm_empty_para(self, para, bm_name):
        """在空段落中添加书签"""
        import xml.etree.ElementTree as ET
        W = self.W
        bm_id = self._next_bm_id()
        bm_start = ET.Element(f'{{{W}}}bookmarkStart')
        bm_start.set('id', str(bm_id))
        bm_start.set('name', bm_name)
        bm_end = ET.Element(f'{{{W}}}bookmarkEnd')
        bm_end.set('id', str(bm_id))
        para.append(bm_start)
        para.append(bm_end)
        return bm_id
    
    def _add_bm_checkbox(self, para, bm_name, before_text):
        """在段落中包裹指定文字后面的□字符"""
        import xml.etree.ElementTree as ET, copy
        ns = self.NAMESPACES
        W = self.W
        for run in para.findall('w:r', ns):
            t = run.find('w:t', ns)
            if t is not None and t.text and before_text in t.text:
                text = t.text
                idx = text.index(before_text) + len(before_text)
                after = text[idx:]
                if '□' not in after:
                    continue
                box_idx = after.index('□')
                before_box = text[:idx + box_idx]
                after_box = text[idx + box_idx + 1:]
                
                t.text = before_box
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                
                bm_id = self._next_bm_id()
                box_run = copy.deepcopy(run)
                box_t = box_run.find('w:t', ns)
                box_t.text = '□'
                box_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                # 删除box_run中除rPr和t外的子元素
                to_remove = [c for c in box_run if c.tag.split('}')[-1] not in ('rPr', 't')]
                for c in to_remove:
                    box_run.remove(c)
                
                bm_start = ET.Element(f'{{{W}}}bookmarkStart')
                bm_start.set('id', str(bm_id))
                bm_start.set('name', bm_name)
                bm_end = ET.Element(f'{{{W}}}bookmarkEnd')
                bm_end.set('id', str(bm_id))
                
                run_idx = list(para).index(run)
                insert_items = [bm_start, box_run, bm_end]
                if after_box:
                    after_run = copy.deepcopy(run)
                    after_t = after_run.find('w:t', ns)
                    after_t.text = after_box
                    after_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    insert_items.append(after_run)
                
                for j, item in enumerate(insert_items):
                    para.insert(run_idx + 1 + j, item)
                return bm_id
        return None
    
    # ================================================================
    # 添加所有书签到模板
    # ================================================================
    def _add_bookmarks(self):
        import xml.etree.ElementTree as ET, os
        ns = self.NAMESPACES
        W = self.W
        tables = self.root.findall('.//w:tbl', ns)
        
        def add_person_bms(tbl_idx, row_idx, prefix):
            """自然人行书签"""
            cell = tables[tbl_idx].findall('w:tr', ns)[row_idx].findall('w:tc', ns)[1]
            paras = cell.findall('.//w:p', ns)
            
            # P0: 姓名：
            self._add_bm_after_run(paras[0], 0, f'{prefix}_name')
            
            # P1: 性别checkbox
            self._add_bm_checkbox(paras[1], f'{prefix}_gender_m', '男')
            self._add_bm_checkbox(paras[1], f'{prefix}_gender_f', '女')
            
            # P2: 出生日期+民族
            # Run0:"出生日期：" Run1-5:"年月日空白" Run6:"日...民族："
            # 书签1：wrap Run1-5（替换年月日模板）
            self._add_bm_wrap_runs(paras[2], 1, 5, f'{prefix}_birthdate')
            # 书签2：after Run6（民族后插入）
            self._add_bm_after_run(paras[2], 6, f'{prefix}_ethnicity')
            
            # P3: 工作单位+职务+联系电话
            # 拆分R0中的"职务："，拆分"联系电话："，然后每个标签run后加书签
            runs = paras[3].findall('w:r', ns)
            # R0可能含"工作单位：...职务："
            t0 = runs[0].find('w:t', ns)
            if t0 is not None and '职务：' in (t0.text or ''):
                self._split_run_at_text(paras[3], runs[0], '职务：')
            # R1可能含"联系电话："
            runs = paras[3].findall('w:r', ns)
            for r in runs:
                t = r.find('w:t', ns)
                if t is not None and '联系电话：' in (t.text or ''):
                    self._split_run_at_text(paras[3], r, '联系电话：')
                    break
            # 现在run结构：R0:工作单位空白, R1:职务空白, R2:空白, R3:联系电话：
            # 在每个标签run后加书签
            self._add_bm_after_run(paras[3], 0, f'{prefix}_work')
            runs = paras[3].findall('w:r', ns)
            for ri, r in enumerate(runs):
                t = r.find('w:t', ns)
                if t is not None and '职务：' in (t.text or ''):
                    self._add_bm_after_run(paras[3], ri, f'{prefix}_position')
                elif t is not None and '联系电话：' in (t.text or ''):
                    self._add_bm_after_run(paras[3], ri, f'{prefix}_phone')
            
            # P4: 住所地
            runs = paras[4].findall('w:r', ns)
            self._add_bm_after_run(paras[4], len(runs) - 1, f'{prefix}_address')
            
            # P5: 经常居住地
            self._add_bm_after_run(paras[5], 0, f'{prefix}_residence')
            
            # P6: 证件类型
            self._add_bm_after_run(paras[6], 0, f'{prefix}_id_type')
            
            # P7: 证件号码
            self._add_bm_after_run(paras[7], 0, f'{prefix}_id_number')
        
        def add_company_bms(tbl_idx, row_idx, prefix):
            """法人行书签"""
            cell = tables[tbl_idx].findall('w:tr', ns)[row_idx].findall('w:tc', ns)[1]
            paras = cell.findall('.//w:p', ns)
            
            # P0: 名称
            self._add_bm_after_run(paras[0], 0, f'{prefix}_name')
            
            # P1: 住所地
            runs = paras[1].findall('w:r', ns)
            self._add_bm_after_run(paras[1], len(runs) - 1, f'{prefix}_address')
            
            # P2: 注册地
            self._add_bm_after_run(paras[2], 0, f'{prefix}_reg_addr')
            
            # P3: 法定代表人+职务+联系电话
            # 拆分：Run0含"法定代表人 / 负责人："，Run1空白，Run2含"职务：...联系电话："
            runs = paras[3].findall('w:r', ns)
            for r in runs:
                t = r.find('w:t', ns)
                if t is not None and '职务：' in (t.text or ''):
                    self._split_run_at_text(paras[3], r, '职务：')
            runs = paras[3].findall('w:r', ns)
            for r in runs:
                t = r.find('w:t', ns)
                if t is not None and '联系电话：' in (t.text or ''):
                    self._split_run_at_text(paras[3], r, '联系电话：')
                    break
            
            # 在法定代表人run后、职务run后、联系电话run后加书签
            runs = paras[3].findall('w:r', ns)
            for ri, r in enumerate(runs):
                t = r.find('w:t', ns)
                if t is not None:
                    text = t.text or ''
                    if '法定代表人' in text:
                        self._add_bm_after_run(paras[3], ri, f'{prefix}_legal_person')
                    elif text.strip().startswith('职务'):
                        self._add_bm_after_run(paras[3], ri, f'{prefix}_position')
                    elif text.strip().startswith('联系电话'):
                        self._add_bm_after_run(paras[3], ri, f'{prefix}_phone')
            
            # P4: 统一社会信用代码
            self._add_bm_after_run(paras[4], 0, f'{prefix}_credit_code')
            
            # P5-P9: 类型checkbox
            self._add_bm_checkbox(paras[5], f'{prefix}_type_llc', '有限责任公司')
            self._add_bm_checkbox(paras[5], f'{prefix}_type_jsc', '股份有限公司')
            self._add_bm_checkbox(paras[5], f'{prefix}_type_listed', '上市公司')
            self._add_bm_checkbox(paras[6], f'{prefix}_type_other_ent', '其他企业法人')
            self._add_bm_checkbox(paras[6], f'{prefix}_type_institution', '事业单位')
            self._add_bm_checkbox(paras[6], f'{prefix}_type_social', '社会团体')
            self._add_bm_checkbox(paras[7], f'{prefix}_type_foundation', '基金会')
            self._add_bm_checkbox(paras[7], f'{prefix}_type_service', '社会服务机构')
            self._add_bm_checkbox(paras[7], f'{prefix}_type_organ', '机关法人')
            self._add_bm_checkbox(paras[8], f'{prefix}_type_rural', '农村集体经济组织法人')
            self._add_bm_checkbox(paras[8], f'{prefix}_type_urban_rural', '城镇农村的合作经济组织法人')
            self._add_bm_checkbox(paras[8], f'{prefix}_type_grassroots', '基层群众性自治组织法人')
            self._add_bm_checkbox(paras[9], f'{prefix}_type_sole', '个人独资企业')
            self._add_bm_checkbox(paras[9], f'{prefix}_type_partner', '合伙企业')
            self._add_bm_checkbox(paras[9], f'{prefix}_type_pro_svc', '不具有法人资格的专业服务机构')
            self._add_bm_checkbox(paras[9], f'{prefix}_own_state', '国有')
            self._add_bm_checkbox(paras[9], f'{prefix}_own_private', '民营')
        
        # ====== 开始添加 ======
        
        # 表0：原告自然人(R2)、原告法人(R3)
        add_person_bms(0, 2, 'pl1')
        add_company_bms(0, 3, 'plc1')
        
        # 表1：代理人(R0)、被告自然人(R1)、被告法人(R2)、第三人自然人(R3)
        # 代理人
        ag_cell = tables[1].findall('w:tr', ns)[0].findall('w:tc', ns)[1]
        ag_paras = ag_cell.findall('.//w:p', ns)
        self._add_bm_checkbox(ag_paras[0], 'ag1_has', '有')
        self._add_bm_after_run(ag_paras[1], 0, 'ag1_name')
        # P2: 单位+职务+联系电话
        runs = ag_paras[2].findall('w:r', ns)
        for r in runs:
            t = r.find('w:t', ns)
            if t is not None and '职务：' in (t.text or ''):
                self._split_run_at_text(ag_paras[2], r, '职务：')
        runs = ag_paras[2].findall('w:r', ns)
        for r in runs:
            t = r.find('w:t', ns)
            if t is not None and '联系电话：' in (t.text or ''):
                self._split_run_at_text(ag_paras[2], r, '联系电话：')
                break
        runs = ag_paras[2].findall('w:r', ns)
        for ri, r in enumerate(runs):
            t = r.find('w:t', ns)
            if t is not None:
                text = t.text or ''
                if text.strip().startswith('单位'):
                    self._add_bm_after_run(ag_paras[2], ri, 'ag1_dept')
                elif text.strip().startswith('职务'):
                    self._add_bm_after_run(ag_paras[2], ri, 'ag1_position')
                elif text.strip().startswith('联系电话'):
                    self._add_bm_after_run(ag_paras[2], ri, 'ag1_phone')
        # P3: 代理权限
        self._add_bm_checkbox(ag_paras[3], 'ag1_auth_general', '一般授权')
        self._add_bm_checkbox(ag_paras[3], 'ag1_auth_special', '特别授权')
        self._add_bm_checkbox(ag_paras[3], 'ag1_auth_none', '无')
        
        add_person_bms(1, 1, 'dp1')
        add_company_bms(1, 2, 'dc1')
        add_person_bms(1, 3, 'tp1')
        
        # 表2：第三人法人(R0)、诉讼请求(R3-R12)
        add_company_bms(2, 0, 'tc1')
        
        t2_rows = tables[2].findall('w:tr', ns)
        
        # R3: 1.医疗费
        cell = t2_rows[3].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        self._add_bm_after_run(ps[0], 0, 'sq1_medical_period')
        self._add_bm_after_run(ps[1], 0, 'sq1_medical_fee')
        self._add_bm_checkbox(ps[2], 'sq1_medical_cert_y', '有')
        self._add_bm_checkbox(ps[2], 'sq1_medical_cert_n', '无')
        
        # R4: 2.护理费
        cell = t2_rows[4].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        self._add_bm_after_run(ps[0], 0, 'sq1_nursing_detail')
        self._add_bm_after_run(ps[1], 0, 'sq1_nursing_fee')
        self._add_bm_checkbox(ps[2], 'sq1_nursing_cert_y', '有')
        self._add_bm_checkbox(ps[2], 'sq1_nursing_cert_n', '无')
        
        # R5: 3.营养费 - "营养费[空白]元"
        cell = t2_rows[5].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        self._add_bm_wrap_runs(ps[0], 1, 1, 'sq1_nutrition')
        self._add_bm_checkbox(ps[1], 'sq1_nutrition_cert_y', '有')
        self._add_bm_checkbox(ps[1], 'sq1_nutrition_cert_n', '无')
        
        # R6: 4.住院伙食补助费 - single run
        cell = t2_rows[6].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        self._add_bm_after_run(ps[0], 0, 'sq1_hospital_meal')
        self._add_bm_checkbox(ps[1], 'sq1_meal_cert_y', '有')
        self._add_bm_checkbox(ps[1], 'sq1_meal_cert_n', '无')
        
        # R7: 5.误工费
        cell = t2_rows[7].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        self._add_bm_after_run(ps[0], 0, 'sq1_lost_wage')
        
        # R8: 6.交通费 - "交通费[空白]元"
        cell = t2_rows[8].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        self._add_bm_wrap_runs(ps[0], 1, 1, 'sq1_traffic')
        self._add_bm_checkbox(ps[1], 'sq1_traffic_cert_y', '有')
        self._add_bm_checkbox(ps[1], 'sq1_traffic_cert_n', '无')
        
        # R9: 7.残疾赔偿金+被扶养人生活费 - single run
        cell = t2_rows[9].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        # 拆分"被扶养人生活费"
        runs = ps[0].findall('w:r', ns)
        for r in runs:
            t = r.find('w:t', ns)
            if t is not None and '被扶养人生活费' in (t.text or ''):
                self._split_run_at_text(ps[0], r, '被扶养人生活费')
                break
        runs = ps[0].findall('w:r', ns)
        for ri, r in enumerate(runs):
            t = r.find('w:t', ns)
            if t is not None and '残疾赔偿金' in (t.text or ''):
                self._add_bm_after_run(ps[0], ri, 'sq1_disability')
            elif t is not None and '被扶养人生活费' in (t.text or ''):
                self._add_bm_after_run(ps[0], ri, 'sq1_dependent')
        
        # R10: 8.残疾辅助器具费 - single run
        cell = t2_rows[10].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        self._add_bm_after_run(ps[0], 0, 'sq1_assistive')
        
        # R11: 9.死亡赔偿金+丧葬费
        cell = t2_rows[11].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        runs = ps[0].findall('w:r', ns)
        for r in runs:
            t = r.find('w:t', ns)
            if t is not None and '丧葬费' in (t.text or ''):
                self._split_run_at_text(ps[0], r, '丧葬费')
                break
        runs = ps[0].findall('w:r', ns)
        for ri, r in enumerate(runs):
            t = r.find('w:t', ns)
            if t is not None and '死亡赔偿金' in (t.text or ''):
                self._add_bm_after_run(ps[0], ri, 'sq1_death')
            elif t is not None and '丧葬费' in (t.text or ''):
                self._add_bm_after_run(ps[0], ri, 'sq1_funeral')
        
        # R12: 10.精神损害抚慰金 - single run
        cell = t2_rows[12].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        self._add_bm_after_run(ps[0], 0, 'sq1_mental')
        
        # ====== 表3 ======
        t3_rows = tables[3].findall('w:tr', ns)
        
        # R0: 11.财产损失
        cell = t3_rows[0].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        self._add_bm_after_run(ps[0], 0, 'sq1_property_vehicle')
        self._add_bm_after_run(ps[1], 0, 'sq1_property_stopped')
        self._add_bm_after_run(ps[2], 0, 'sq1_property_other')
        
        # R1: 12.其他费用
        cell = t3_rows[1].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        self._add_bm_after_run(ps[0], 0, 'sq1_other')
        
        # R2: 13.标的总额 - C1空段落
        cell = t3_rows[2].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        self._add_bm_empty_para(ps[0], 'sq1_total')
        
        # R4: 诉前保全
        cell = t3_rows[4].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        self._add_bm_checkbox(ps[0], 'pb1_yes', '是')
        # 拆分"保全法院："和"保全时间："
        runs = ps[0].findall('w:r', ns)
        for r in runs:
            t = r.find('w:t', ns)
            if t is not None and '保全时间：' in (t.text or ''):
                self._split_run_at_text(ps[0], r, '保全时间：')
                break
        runs = ps[0].findall('w:r', ns)
        for ri, r in enumerate(runs):
            t = r.find('w:t', ns)
            if t is not None and '保全法院：' in (t.text or ''):
                self._add_bm_after_run(ps[0], ri, 'pb1_court')
            elif t is not None and '保全时间：' in (t.text or ''):
                self._add_bm_after_run(ps[0], ri, 'pb1_time')
        self._add_bm_after_run(ps[1], 0, 'pb1_case_no')
        self._add_bm_checkbox(ps[2], 'pb1_no', '否')
        
        # R5: 是否申请鉴定
        cell = t3_rows[5].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        self._add_bm_checkbox(ps[0], 'ap1_yes', '是')
        runs = ps[0].findall('w:r', ns)
        for r in runs:
            t = r.find('w:t', ns)
            if t is not None and '否□' in (t.text or ''):
                self._split_run_at_text(ps[0], r, '否')
                break
        runs = ps[0].findall('w:r', ns)
        for ri, r in enumerate(runs):
            t = r.find('w:t', ns)
            if t is not None and '鉴定事项：' in (t.text or ''):
                self._add_bm_after_run(ps[0], ri, 'ap1_detail')
        self._add_bm_checkbox(ps[0], 'ap1_no', '否')
        
        # R8-R12: 事实与理由（C1空段落）
        for ri, bm in [(8, 'fr1_accident'), (9, 'fr1_responsibility'),
                       (10, 'fr1_insurance'), (11, 'fr1_basis'), (12, 'fr1_evidence')]:
            cell = t3_rows[ri].findall('w:tc', ns)[1]
            ps = cell.findall('.//w:p', ns)
            self._add_bm_empty_para(ps[0], bm)
        
        # R14C1: 了解调解
        cell = t3_rows[14].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        # P2: "了解□    不了解□"
        self._add_bm_checkbox(ps[2], 'dr1_know_yes', '了解')
        self._add_bm_checkbox(ps[2], 'dr1_know_no', '不了解')
        
        # R15C1: 先行调解了解
        cell = t3_rows[15].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        for pi, suffix in [(1, '1'), (3, '2'), (5, '3')]:
            self._add_bm_checkbox(ps[pi], f'dr1_prior_yes{suffix}', '了解')
            self._add_bm_checkbox(ps[pi], f'dr1_prior_no{suffix}', '不了解')
        
        # ====== 表4 ======
        t4_rows = tables[4].findall('w:tr', ns)
        
        # R0C1: 4,5点
        cell = t4_rows[0].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        self._add_bm_checkbox(ps[1], 'dr1_prior_yes4', '了解')
        self._add_bm_checkbox(ps[1], 'dr1_prior_no4', '不了解')
        self._add_bm_checkbox(ps[3], 'dr1_prior_yes5', '了解')
        self._add_bm_checkbox(ps[3], 'dr1_prior_no5', '不了解')
        
        # R1C1: 是否先行调解
        cell = t4_rows[1].findall('w:tc', ns)[1]
        ps = cell.findall('.//w:p', ns)
        self._add_bm_checkbox(ps[0], 'dr1_try_yes', '是')
        self._add_bm_checkbox(ps[1], 'dr1_try_no', '否')
        self._add_bm_checkbox(ps[2], 'dr1_try_unsure', '内容')
        
        # ====== 签名区（表外段落） ======
        body = self.root.find('.//w:body', ns)
        for p in body.findall('w:p', ns):
            text = self._get_para_text(p)
            if '具状人' in text:
                runs = p.findall('w:r', ns)
                self._add_bm_after_run(p, len(runs) - 1, 'sig_name')
            elif text.strip().startswith('日期：'):
                self._add_bm_after_run(p, 0, 'sig_date')
        
        # 写回
        doc_path = os.path.join(self.unpacked_dir, 'word/document.xml')
        self._write_xml(doc_path, self.root)
    
    # ================================================================
    # 复制当事人行
    # ================================================================
    def _duplicate_party_rows(self, case):
        import copy, os
        ns = self.NAMESPACES
        W = self.W
        tables = self.root.findall('.//w:tbl', ns)
        
        dup_rules = [
            (0, 2, 'pl', len(case.plaintiffs or [])),
            (0, 3, 'plc', len(case.plaintiffs_company or [])),
            (1, 0, 'ag', len(case.agents or [])),
            (1, 1, 'dp', len(case.defendants_person or [])),
            (1, 2, 'dc', len(case.defendants_company or [])),
            (1, 3, 'tp', len(case.third_parties_person or [])),
            (2, 0, 'tc', len(case.third_parties_company or [])),
        ]
        
        for tbl_idx, row_idx, prefix, count in dup_rules:
            if count <= 0:
                continue
            # 先重命名源行书签为_1
            self._rename_bms_in_row(tables[tbl_idx], row_idx, prefix, 1)
            if count <= 1:
                continue
            rows = tables[tbl_idx].findall('w:tr', ns)
            source_row = rows[row_idx]
            for i in range(2, count + 1):
                new_row = copy.deepcopy(source_row)
                self._rename_bms_in_elem(new_row, prefix, i)
                rows = tables[tbl_idx].findall('w:tr', ns)
                si = rows.index(source_row) if source_row in rows else row_idx
                tables[tbl_idx].insert(si + (i - 1), new_row)
        
        doc_path = os.path.join(self.unpacked_dir, 'word/document.xml')
        self._write_xml(doc_path, self.root)
    
    def _rename_bms_in_row(self, tbl, row_idx, prefix, n):
        row = tbl.findall('w:tr', self.NAMESPACES)[row_idx]
        self._rename_bms_in_elem(row, prefix, n)
    
    def _rename_bms_in_elem(self, elem, prefix, n):
        W = self.W
        old = f'{prefix}1_'
        new = f'{prefix}{n}_'
        for bm in elem.iter(f'{{{W}}}bookmarkStart'):
            name = bm.get('name', '')
            if name.startswith(old):
                bm.set('name', new + name[len(old):])
    
    # ================================================================
    # 构建填充数据
    # ================================================================
    def _build_fill_map(self, case, result):
        fm = {}
        
        # 原告自然人
        for i, p in enumerate(case.plaintiffs or []):
            n = i + 1
            if p.get('name'): fm[f'pl{n}_name'] = ' ' + p['name']
            if p.get('gender'):
                fm[f'pl{n}_gender_m'] = '☑' if p['gender'] == '男' else '□'
                fm[f'pl{n}_gender_f'] = '☑' if p['gender'] == '女' else '□'
            if p.get('birthdate'): fm[f'pl{n}_birthdate'] = ' ' + p['birthdate']
            if p.get('ethnicity'): fm[f'pl{n}_ethnicity'] = ' ' + p['ethnicity']
            if p.get('work'): fm[f'pl{n}_work'] = p['work']
            if p.get('position'): fm[f'pl{n}_position'] = p['position']
            if p.get('phone'): fm[f'pl{n}_phone'] = ' ' + p['phone']
            if p.get('address'): fm[f'pl{n}_address'] = ' ' + p['address']
            if p.get('residence'): fm[f'pl{n}_residence'] = p['residence']
            if p.get('id_type'): fm[f'pl{n}_id_type'] = p['id_type']
            if p.get('id_number'): fm[f'pl{n}_id_number'] = ' ' + p['id_number']
        
        # 原告法人
        for i, c in enumerate(case.plaintiffs_company or []):
            n = i + 1
            if c.get('name'): fm[f'plc{n}_name'] = ' ' + c['name']
            if c.get('address'): fm[f'plc{n}_address'] = ' ' + c['address']
            if c.get('legal_person'): fm[f'plc{n}_legal_person'] = ' ' + c['legal_person']
            if c.get('position'): fm[f'plc{n}_position'] = c['position']
            if c.get('phone'): fm[f'plc{n}_phone'] = ' ' + c['phone']
            if c.get('credit_code'): fm[f'plc{n}_credit_code'] = c['credit_code']
            if c.get('company_type'):
                ct = c['company_type']
                fm[f'plc{n}_type_llc'] = '☑' if ct == '有限责任公司' else '□'
                fm[f'plc{n}_own_private'] = '☑' if ct in ('有限责任公司', '股份有限公司') else '□'
        
        # 被告自然人
        for i, d in enumerate(case.defendants_person or []):
            n = i + 1
            if d.get('name'): fm[f'dp{n}_name'] = ' ' + d['name']
            if d.get('gender'):
                fm[f'dp{n}_gender_m'] = '☑' if d['gender'] == '男' else '□'
                fm[f'dp{n}_gender_f'] = '☑' if d['gender'] == '女' else '□'
            if d.get('birthdate'): fm[f'dp{n}_birthdate'] = ' ' + d['birthdate']
            if d.get('ethnicity'): fm[f'dp{n}_ethnicity'] = ' ' + d['ethnicity']
            if d.get('address'): fm[f'dp{n}_address'] = ' ' + d['address']
            if d.get('id_number'): fm[f'dp{n}_id_number'] = ' ' + d['id_number']
            if d.get('phone'): fm[f'dp{n}_phone'] = ' ' + d['phone']
        
        # 被告法人
        for i, c in enumerate(case.defendants_company or []):
            n = i + 1
            if c.get('name'): fm[f'dc{n}_name'] = ' ' + c['name']
            if c.get('address'): fm[f'dc{n}_address'] = ' ' + c['address']
            if c.get('legal_person'): fm[f'dc{n}_legal_person'] = ' ' + c['legal_person']
            if c.get('position'): fm[f'dc{n}_position'] = c['position']
            if c.get('phone'): fm[f'dc{n}_phone'] = ' ' + c['phone']
            if c.get('credit_code'): fm[f'dc{n}_credit_code'] = c['credit_code']
            if c.get('company_type'):
                ct = c['company_type']
                fm[f'dc{n}_type_llc'] = '☑' if ct == '有限责任公司' else '□'
                fm[f'dc{n}_own_private'] = '☑' if ct in ('有限责任公司', '股份有限公司') else '□'
        
        # 诉讼请求
        if result.get('medical_fee'): fm['sq1_medical_fee'] = f' {result["medical_fee"]:,.2f} 元'
        if result.get('nutrition_fee'): fm['sq1_nutrition'] = f' {result["nutrition_fee"]:,.2f} '
        if result.get('hospital_meal_fee'): fm['sq1_hospital_meal'] = f' {result["hospital_meal_fee"]:,.2f} 元'
        if result.get('nursing_fee'): fm['sq1_nursing_fee'] = f' {result["nursing_fee"]:,.2f} 元'
        if result.get('lost_wage'): fm['sq1_lost_wage'] = f' {result["lost_wage"]:,.2f} 元'
        if result.get('traffic_fee'): fm['sq1_traffic'] = f' {result["traffic_fee"]:,.2f} '
        if result.get('disability_compensation'): fm['sq1_disability'] = f' {result["disability_compensation"]:,.2f} 元'
        if result.get('dependent_living'): fm['sq1_dependent'] = f' {result["dependent_living"]:,.2f} 元'
        if result.get('assistive_device_fee'): fm['sq1_assistive'] = f' {result["assistive_device_fee"]:,.2f} 元'
        if result.get('death_compensation'): fm['sq1_death'] = f' {result["death_compensation"]:,.2f} 元'
        if result.get('funeral_fee'): fm['sq1_funeral'] = f' {result["funeral_fee"]:,.2f} 元'
        if result.get('mental_damage_fee'): fm['sq1_mental'] = f' {result["mental_damage_fee"]:,.2f} 元'
        if result.get('property_damage'): fm['sq1_property_vehicle'] = f' {result["property_damage"]:,.2f} 元'
        if result.get('other_fee'): fm['sq1_other'] = f' {result["other_fee"]:,.2f} 元'
        fm['sq1_total'] = f' {result["total_amount"]:,.2f} 元'
        
        # 诉前保全
        fm['pb1_no'] = '☑'
        fm['ap1_yes'] = '☑' if (case.disability_grade or case.injury_type in (CaseData.INJURY_DISABILITY, CaseData.INJURY_DEATH)) else '□'
        fm['ap1_no'] = '□' if fm['ap1_yes'] == '☑' else '☑'
        if fm['ap1_yes'] == '☑': fm['ap1_detail'] = ' 伤残/死亡鉴定'
        
        # 事实与理由
        if case.accident_detail: fm['fr1_accident'] = case.accident_detail
        if case.responsibility_result: fm['fr1_responsibility'] = case.responsibility_result
        if case.insurance_info: fm['fr1_insurance'] = case.insurance_info
        
        liability_pct = HEBEI_STANDARD['liability_ratio'].get(case.liability_type, 1.0) * 100
        basis = f'被告{case.liability_type}（{liability_pct:.0f}%），应当承担赔偿责任。'
        if case.defendants_company:
            basis += '保险公司应在交强险和商业三者险限额内承担赔偿责任。'
        fm['fr1_basis'] = basis
        
        if case.evidence_list:
            items = [f'{i+1}. {x}' for i, x in enumerate(case.evidence_list)]
            fm['fr1_evidence'] = '；'.join(items) + '。'
        
        # 签名
        if case.plaintiffs: fm['sig_name'] = case.plaintiffs[0].get('name', '')
        fm['sig_date'] = '     年    月    日'
        
        # 纠纷意愿
        fm['dr1_know_yes'] = '☑'
        fm['dr1_know_no'] = '□'
        fm['dr1_try_no'] = '☑'
        fm['dr1_try_yes'] = '□'
        fm['dr1_try_unsure'] = '□'
        
        return fm
    
    # ================================================================
    # 按书签填充
    # ================================================================
    def _fill_by_bookmarks(self, fill_map):
        import xml.etree.ElementTree as ET, copy, os
        ns = self.NAMESPACES
        W = self.W
        
        # 构建书签映射
        all_starts = list(self.root.iter(f'{{{W}}}bookmarkStart'))
        all_ends = list(self.root.iter(f'{{{W}}}bookmarkEnd'))
        
        bm_map = {}
        for bs in all_starts:
            name = bs.get('name', '')
            if name in fill_map:
                bid = bs.get('id')
                for be in all_ends:
                    if be.get('id') == bid:
                        bm_map[name] = (bs, be)
                        break
        
        for name, (bs, be) in bm_map.items():
            value = fill_map[name]
            if value in ('☑', '□'):
                self._fill_cb(bs, be, value)
            else:
                self._fill_text(bs, be, value)
        
        doc_path = os.path.join(self.unpacked_dir, 'word/document.xml')
        self._write_xml(doc_path, self.root)
    
    def _fill_text(self, bs, be, value):
        """填充文本书签"""
        import copy, xml.etree.ElementTree as ET
        W = self.W
        
        # 找书签所在段落
        parent = None
        for p in self.root.iter(f'{{{W}}}p'):
            if bs in list(p) and be in list(p):
                parent = p
                break
        if parent is None:
            return
        
        # 清除bookmarkStart和bookmarkEnd之间的元素
        children = list(parent)
        si = children.index(bs)
        ei = children.index(be)
        for child in children[si + 1:ei]:
            parent.remove(child)
        
        # 在bs之后插入带格式的run
        # 复制前一个run的格式
        new_run = None
        if si > 0:
            prev = children[si - 1]
            ptag = prev.tag.split('}')[-1] if '}' in prev.tag else prev.tag
            if ptag == 'r':
                new_run = copy.deepcopy(prev)
                # 只保留rPr和t
                to_remove = [c for c in new_run if c.tag.split('}')[-1] not in ('rPr', 't')]
                for c in to_remove:
                    new_run.remove(c)
                # 设置文本
                t_elems = new_run.findall(f'{{{W}}}t')
                if t_elems:
                    t_elems[0].text = value
                    t_elems[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    # 删除多余的t
                    for t in t_elems[1:]:
                        new_run.remove(t)
                else:
                    t = ET.Element(f'{{{W}}}t')
                    t.text = value
                    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    new_run.append(t)
        
        if new_run is None:
            new_run = ET.Element(f'{{{W}}}r')
            t = ET.Element(f'{{{W}}}t')
            t.text = value
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            new_run.append(t)
        
        # 重新获取位置（因为可能删除了元素）
        children = list(parent)
        si = children.index(bs)
        parent.insert(si + 1, new_run)
    
    def _fill_cb(self, bs, be, value):
        """填充checkbox书签"""
        W = self.W
        parent = None
        for p in self.root.iter(f'{{{W}}}p'):
            if bs in list(p) and be in list(p):
                parent = p
                break
        if parent is None:
            return
        
        children = list(parent)
        si = children.index(bs)
        ei = children.index(be)
        for child in children[si + 1:ei]:
            for t in child.iter(f'{{{W}}}t'):
                if t.text and '□' in t.text:
                    t.text = t.text.replace('□', value)
                    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
