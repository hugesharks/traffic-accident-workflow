# 交通事故常见损伤三期评定数据（GA/T 1193-2014）
# 用于伤情评估模块：根据损伤描述自动推导误工期/护理期/营养期

THREE_PERIOD_DATA = {
    # ==================== 4.头部损伤 ====================
    "头皮血肿_皮下": {"lost_work": (7, 15), "nursing": (0, 0), "nutrition": (0, 0)},
    "头皮血肿_帽状腱膜下": {"lost_work": (15, 30), "nursing": (1, 7), "nutrition": (1, 7)},
    "头皮血肿_穿刺抽血": {"lost_work": (30, 60), "nursing": (1, 15), "nutrition": (7, 15)},
    "头皮创_小": {"lost_work": (20, 30), "nursing": (1, 7), "nutrition": (1, 7)},
    "头皮创_大": {"lost_work": (45, 60), "nursing": (1, 7), "nutrition": (7, 15)},
    "颅盖骨骨折_线状": {"lost_work": (30, 60), "nursing": (15, 20), "nutrition": (20, 30)},
    "颅盖骨骨折_凹陷粉碎_非手术": {"lost_work": (90, 120), "nursing": (30, 60), "nutrition": (30, 60)},
    "颅盖骨骨折_凹陷粉碎_手术": {"lost_work": (120, 150), "nursing": (30, 60), "nutrition": (30, 60)},
    "颅底骨折_单纯": {"lost_work": (60, 90), "nursing": (15, 20), "nutrition": (20, 30)},
    "颅底骨折_伴脑脊液漏": {"lost_work": (90, 120), "nursing": (30, 60), "nutrition": (30, 60)},
    "颅脑损伤_轻型": {"lost_work": (30, 45), "nursing": (0, 0), "nutrition": (0, 0)},
    "颅脑损伤_中型": {"lost_work": (90, 180), "nursing": (30, 60), "nutrition": (30, 60)},

    # ==================== 5.面部损伤 ====================
    "眼部损伤_眼睑裂伤": {"lost_work": (15, 30), "nursing": (1, 7), "nutrition": (7, 15)},
    "眼部损伤_泪器损伤": {"lost_work": (30, 45), "nursing": (7, 15), "nutrition": (7, 15)},
    "眼部损伤_眼球损伤": {"lost_work": (45, 90), "nursing": (15, 30), "nutrition": (15, 30)},
    "鼻骨骨折_无移位": {"lost_work": (15, 30), "nursing": (0, 0), "nutrition": (7, 15)},
    "鼻骨骨折_有移位手术": {"lost_work": (30, 60), "nursing": (7, 15), "nutrition": (15, 20)},
    "上颌骨骨折": {"lost_work": (60, 120), "nursing": (15, 30), "nutrition": (30, 60)},
    "下颌骨骨折_保守": {"lost_work": (60, 90), "nursing": (15, 30), "nutrition": (30, 45)},
    "下颌骨骨折_手术": {"lost_work": (90, 120), "nursing": (15, 30), "nutrition": (45, 60)},

    # ==================== 6.颈部损伤 ====================
    "颈部软组织损伤": {"lost_work": (15, 30), "nursing": (1, 7), "nutrition": (7, 15)},
    "颈椎骨折_无神经损伤": {"lost_work": (90, 120), "nursing": (30, 60), "nutrition": (30, 60)},
    "颈椎骨折_伴神经损伤": {"lost_work": (180, 365), "nursing": (60, 180), "nutrition": (60, 90)},
    "颈部气管损伤": {"lost_work": (30, 60), "nursing": (7, 15), "nutrition": (15, 20)},

    # ==================== 7.胸部损伤 ====================
    "肋骨骨折_1-3根": {"lost_work": (30, 45), "nursing": (7, 15), "nutrition": (15, 30)},
    "肋骨骨折_4-7根": {"lost_work": (60, 90), "nursing": (15, 30), "nutrition": (30, 45)},
    "肋骨骨折_8根以上": {"lost_work": (90, 120), "nursing": (30, 60), "nutrition": (45, 60)},
    "肋骨骨折_伴血气胸": {"lost_work": (90, 150), "nursing": (30, 60), "nutrition": (45, 60)},
    "胸骨骨折": {"lost_work": (60, 90), "nursing": (15, 30), "nutrition": (30, 45)},
    "肺损伤": {"lost_work": (60, 90), "nursing": (15, 30), "nutrition": (30, 45)},
    "心脏损伤": {"lost_work": (120, 180), "nursing": (30, 60), "nutrition": (45, 60)},

    # ==================== 9.脊柱骨盆损伤 ====================
    "脊柱骨折_无神经损伤": {"lost_work": (90, 150), "nursing": (30, 60), "nutrition": (30, 60)},
    "脊柱骨折_伴神经损伤": {"lost_work": (180, 365), "nursing": (60, 180), "nutrition": (60, 90)},
    "脊柱骨折_手术": {"lost_work": (120, 180), "nursing": (30, 60), "nutrition": (45, 60)},
    "椎间盘突出_保守": {"lost_work": (60, 120), "nursing": (15, 30), "nutrition": (15, 30)},
    "椎间盘突出_手术": {"lost_work": (90, 150), "nursing": (30, 60), "nutrition": (30, 45)},
    "骨盆骨折_稳定型": {"lost_work": (90, 120), "nursing": (30, 45), "nutrition": (30, 45)},
    "骨盆骨折_不稳定型": {"lost_work": (120, 180), "nursing": (45, 60), "nutrition": (45, 60)},

    # ==================== 10.四肢损伤（最常见的交通事故伤） ====================
    # 上肢
    "锁骨骨折_保守": {"lost_work": (60, 90), "nursing": (15, 30), "nutrition": (30, 45)},
    "锁骨骨折_手术": {"lost_work": (90, 120), "nursing": (15, 30), "nutrition": (30, 45)},
    "肱骨骨折_保守": {"lost_work": (60, 120), "nursing": (15, 30), "nutrition": (30, 60)},
    "肱骨骨折_手术": {"lost_work": (90, 150), "nursing": (15, 30), "nutrition": (30, 60)},
    "尺骨鹰嘴骨折": {"lost_work": (60, 90), "nursing": (15, 30), "nutrition": (30, 45)},
    "尺桡骨骨折_保守": {"lost_work": (90, 120), "nursing": (15, 30), "nutrition": (30, 45)},
    "尺桡骨骨折_手术": {"lost_work": (90, 150), "nursing": (15, 30), "nutrition": (30, 45)},
    "桡骨远端骨折_保守": {"lost_work": (60, 90), "nursing": (15, 30), "nutrition": (30, 45)},
    "桡骨远端骨折_手术": {"lost_work": (90, 120), "nursing": (15, 30), "nutrition": (30, 45)},

    # 下肢
    "股骨颈骨折_保守": {"lost_work": (240, 300), "nursing": (90, 180), "nutrition": (90, 120)},
    "股骨颈骨折_手术": {"lost_work": (180, 270), "nursing": (60, 120), "nutrition": (60, 90)},
    "股骨粗隆间骨折_保守": {"lost_work": (180, 240), "nursing": (60, 120), "nutrition": (60, 90)},
    "股骨粗隆间骨折_手术": {"lost_work": (150, 210), "nursing": (45, 90), "nutrition": (45, 60)},
    "股骨干骨折_保守": {"lost_work": (120, 180), "nursing": (30, 90), "nutrition": (60, 90)},
    "股骨干骨折_手术": {"lost_work": (90, 150), "nursing": (30, 60), "nutrition": (45, 60)},
    "股骨远端骨折": {"lost_work": (120, 180), "nursing": (30, 60), "nutrition": (45, 60)},
    "髌骨骨折_保守": {"lost_work": (90, 120), "nursing": (30, 45), "nutrition": (30, 45)},
    "髌骨骨折_手术": {"lost_work": (90, 150), "nursing": (30, 45), "nutrition": (30, 45)},
    "胫骨平台骨折": {"lost_work": (120, 150), "nursing": (30, 60), "nutrition": (60, 90)},
    "胫腓骨骨折_胫骨": {"lost_work": (120, 180), "nursing": (30, 90), "nutrition": (60, 90)},
    "胫腓骨骨折_腓骨": {"lost_work": (60, 90), "nursing": (30, 60), "nutrition": (30, 60)},
    "胫腓骨骨折_双骨折": {"lost_work": (120, 180), "nursing": (30, 90), "nutrition": (60, 90)},
    "胫腓骨骨折_开放性": {"lost_work": (150, 210), "nursing": (45, 90), "nutrition": (60, 90)},
    "胫骨近端粉碎性骨折": {"lost_work": (150, 180), "nursing": (60, 90), "nutrition": (60, 90)},
    "踝部骨折_单踝": {"lost_work": (90, 120), "nursing": (30, 60), "nutrition": (60, 90)},
    "踝部骨折_双踝": {"lost_work": (90, 180), "nursing": (30, 60), "nutrition": (60, 90)},
    "踝部骨折_三踝": {"lost_work": (90, 180), "nursing": (30, 60), "nutrition": (60, 90)},
    "跟骨骨折_保守": {"lost_work": (90, 150), "nursing": (30, 60), "nutrition": (30, 60)},
    "跟骨骨折_手术": {"lost_work": (120, 180), "nursing": (30, 60), "nutrition": (45, 60)},

    # 关节
    "肩关节脱位": {"lost_work": (30, 60), "nursing": (15, 30), "nutrition": (15, 30)},
    "肘关节脱位": {"lost_work": (30, 60), "nursing": (15, 30), "nutrition": (15, 30)},
    "髋关节脱位": {"lost_work": (90, 150), "nursing": (30, 60), "nutrition": (45, 60)},
    "膝关节韧带损伤_保守": {"lost_work": (60, 90), "nursing": (15, 30), "nutrition": (30, 45)},
    "膝关节韧带损伤_手术": {"lost_work": (90, 150), "nursing": (30, 60), "nutrition": (30, 45)},
    "半月板损伤_保守": {"lost_work": (60, 90), "nursing": (15, 30), "nutrition": (30, 45)},
    "半月板损伤_手术": {"lost_work": (90, 120), "nursing": (15, 30), "nutrition": (30, 45)},

    # ==================== 11.手/足损伤 ====================
    "指骨骨折": {"lost_work": (30, 60), "nursing": (7, 15), "nutrition": (15, 30)},
    "掌骨骨折": {"lost_work": (45, 60), "nursing": (15, 30), "nutrition": (15, 30)},
    "腕骨骨折": {"lost_work": (60, 90), "nursing": (15, 30), "nutrition": (30, 45)},
    "趾骨骨折": {"lost_work": (30, 60), "nursing": (7, 15), "nutrition": (15, 30)},
    "跖骨骨折": {"lost_work": (45, 60), "nursing": (15, 30), "nutrition": (15, 30)},

    # ==================== 12.周围神经损伤 ====================
    "臂丛神经损伤": {"lost_work": (90, 180), "nursing": (30, 60), "nutrition": (30, 45)},
    "桡神经损伤": {"lost_work": (90, 150), "nursing": (30, 45), "nutrition": (30, 45)},
    "正中神经损伤": {"lost_work": (90, 150), "nursing": (30, 45), "nutrition": (30, 45)},
    "尺神经损伤": {"lost_work": (90, 150), "nursing": (30, 45), "nutrition": (30, 45)},
    "坐骨神经损伤": {"lost_work": (120, 210), "nursing": (30, 60), "nutrition": (45, 60)},
    "腓总神经损伤": {"lost_work": (90, 180), "nursing": (30, 60), "nutrition": (30, 45)},

    # ==================== 14.烧伤（交通事故起火场景） ====================
    "烧伤_轻度_I度": {"lost_work": (15, 20), "nursing": (0, 0), "nutrition": (7, 15)},
    "烧伤_浅II度_小面积": {"lost_work": (20, 30), "nursing": (7, 15), "nutrition": (15, 20)},
    "烧伤_深II度_小面积": {"lost_work": (30, 45), "nursing": (15, 20), "nutrition": (20, 30)},
    "烧伤_III度_小面积": {"lost_work": (60, 90), "nursing": (30, 45), "nutrition": (30, 45)},
}

# 关键词匹配映射：用户描述→三期数据key
INJURY_KEYWORD_MAP = {
    # 头部
    "头皮血肿": "头皮血肿_皮下",
    "颅骨骨折": "颅盖骨骨折_线状",
    "颅底骨折": "颅底骨折_单纯",
    "脑外伤": "颅脑损伤_轻型",
    "脑挫裂伤": "颅脑损伤_中型",
    "蛛网膜下腔出血": "颅脑损伤_中型",

    # 面部
    "鼻骨骨折": "鼻骨骨折_无移位",
    "上颌骨骨折": "上颌骨骨折",
    "下颌骨骨折": "下颌骨骨折_保守",

    # 颈部
    "颈椎骨折": "颈椎骨折_无神经损伤",
    "颈部软组织": "颈部软组织损伤",
    "挥鞭样损伤": "颈部软组织损伤",
    "颈椎间盘突出": "椎间盘突出_保守",

    # 胸部
    "肋骨骨折": "肋骨骨折_1-3根",
    "胸骨骨折": "胸骨骨折",
    "血气胸": "肋骨骨折_伴血气胸",
    "肺挫伤": "肺损伤",

    # 脊柱骨盆
    "胸椎骨折": "脊柱骨折_无神经损伤",
    "腰椎骨折": "脊柱骨折_无神经损伤",
    "腰椎间盘突出": "椎间盘突出_保守",
    "骨盆骨折": "骨盆骨折_稳定型",
    "截瘫": "脊柱骨折_伴神经损伤",

    # 上肢
    "锁骨骨折": "锁骨骨折_保守",
    "肱骨骨折": "肱骨骨折_保守",
    "尺桡骨骨折": "尺桡骨骨折_保守",
    "桡骨远端骨折": "桡骨远端骨折_保守",
    "肩关节脱位": "肩关节脱位",

    # 下肢（交通事故最常见）
    "股骨颈骨折": "股骨颈骨折_手术",
    "股骨粗隆间骨折": "股骨粗隆间骨折_手术",
    "股骨干骨折": "股骨干骨折_手术",
    "髌骨骨折": "髌骨骨折_保守",
    "胫骨平台骨折": "胫骨平台骨折",
    "胫腓骨骨折": "胫腓骨骨折_双骨折",
    "胫骨骨折": "胫腓骨骨折_胫骨",
    "腓骨骨折": "胫腓骨骨折_腓骨",
    "粉碎性骨折_胫腓": "胫骨近端粉碎性骨折",
    "踝关节骨折": "踝部骨折_单踝",
    "踝部骨折": "踝部骨折_单踝",
    "三踝骨折": "踝部骨折_三踝",
    "双踝骨折": "踝部骨折_双踝",
    "跟骨骨折": "跟骨骨折_保守",
    "髋关节脱位": "髋关节脱位",
    "膝关节韧带": "膝关节韧带损伤_保守",
    "半月板损伤": "半月板损伤_保守",

    # 手足
    "指骨骨折": "指骨骨折",
    "掌骨骨折": "掌骨骨折",
    "跖骨骨折": "跖骨骨折",
    "趾骨骨折": "趾骨骨折",

    # 神经
    "臂丛神经": "臂丛神经损伤",
    "桡神经": "桡神经损伤",
    "腓总神经": "腓总神经损伤",
}


def lookup_three_period(injury_description: str) -> dict:
    """
    根据损伤描述查找三期数据
    
    参数:
        injury_description: 损伤描述文本（如"右胫腓骨粉碎性骨折"）
    
    返回:
        {
            'matched_key': '胫腓骨骨折_双骨折',
            'lost_work': (120, 180),  # 误工期范围（天）
            'nursing': (30, 90),      # 护理期范围（天）
            'nutrition': (60, 90),    # 营养期范围（天）
            'lost_work_mid': 150,     # 误工期中值
            'nursing_mid': 60,        # 护理期中值
            'nutrition_mid': 75,      # 营养期中值
            'source': 'GA/T 1193-2014',
        }
        或 None（未匹配）
    """
    results = lookup_three_period_multi(injury_description)
    if results:
        return results[0]
    return None


def lookup_three_period_multi(injury_description: str) -> list:
    """
    根据损伤描述查找所有匹配的三期数据（支持多处损伤）
    
    返回: [dict, ...] 或 []
    """
    # 清洗修饰词
    clean = injury_description
    modifiers = ['右侧', '左侧', '右', '左', '双', '多发', '粉碎性', '开放性',
                 '闭合性', '横断', '斜行', '螺旋', '压缩性', '爆裂性']
    for mod in modifiers:
        clean = clean.replace(mod, '')

    # 匹配所有关键词
    matched = []
    seen_keys = set()
    for keyword, data_key in INJURY_KEYWORD_MAP.items():
        if keyword in clean and data_key not in seen_keys:
            seen_keys.add(data_key)
            matched.append((len(keyword), keyword, data_key))

    if not matched:
        return []

    # 按关键词长度降序（更具体的优先）
    matched.sort(reverse=True)

    results = []
    for _, keyword, data_key in matched:
        data = THREE_PERIOD_DATA[data_key]
        results.append({
            'matched_key': data_key,
            'keyword': keyword,
            'lost_work': data['lost_work'],
            'nursing': data['nursing'],
            'nutrition': data['nutrition'],
            'lost_work_mid': (data['lost_work'][0] + data['lost_work'][1]) // 2,
            'nursing_mid': (data['nursing'][0] + data['nursing'][1]) // 2,
            'nutrition_mid': (data['nutrition'][0] + data['nutrition'][1]) // 2,
            'source': 'GA/T 1193-2014',
        })

    return results
