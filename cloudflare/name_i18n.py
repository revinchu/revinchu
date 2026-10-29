"""병원·명소·가격 항목·행정구역 이름의 외국어 표기 (import_public_data.py 가 사용).

공공데이터(심평원)에는 외국어 병원명이 없어 직접 정리했다.
- 대형병원·대학병원: 각 병원이 쓰는 공식 영문/중문/일문 명칭을 우선
- 그 밖의 병원: 고유명은 음역, 병원 종류는 번역 (예: 굿윌치과병원 → Goodwill Dental Hospital)
새 병원이 추가되면 여기 없을 수 있다. 그때는 국문 이름을 그대로 쓰고 import 로그에 표시한다.
"""

# 국문 이름(심평원 yadmNm) → {en, zh, ja, vi}
HOSPITALS = {
    # 치과병원
    "서울대학교치과병원": ("Seoul National University Dental Hospital", "首尔大学齿科医院", "ソウル大学校歯科病院", "Bệnh viện Nha khoa Đại học Quốc gia Seoul"),
    "연세대학교치과대학치과병원": ("Yonsei University Dental Hospital", "延世大学齿科医院", "延世大学校歯科病院", "Bệnh viện Nha khoa Đại học Yonsei"),
    "경희대학교치과병원": ("Kyung Hee University Dental Hospital", "庆熙大学齿科医院", "慶熙大学校歯科病院", "Bệnh viện Nha khoa Đại học Kyung Hee"),
    "강동경희대학교치과병원": ("Kyung Hee University Dental Hospital at Gangdong", "江东庆熙大学齿科医院", "江東慶熙大学校歯科病院", "Bệnh viện Nha khoa Đại học Kyung Hee Gangdong"),
    "석플란트치과병원": ("Seokplant Dental Hospital", "石普兰特齿科医院", "ソクプラント歯科病院", "Bệnh viện Nha khoa Seokplant"),
    "관악서울대학교치과병원": ("Seoul National University Gwanak Dental Hospital", "首尔大学冠岳齿科医院", "ソウル大学校冠岳歯科病院", "Bệnh viện Nha khoa Gwanak Đại học Quốc gia Seoul"),
    "목동예치과병원": ("Mokdong Ye Dental Hospital", "木洞艺齿科医院", "木洞イェ歯科病院", "Bệnh viện Nha khoa Ye Mokdong"),
    "더와이즈치과병원": ("The Wise Dental Hospital", "The Wise齿科医院", "ザ・ワイズ歯科病院", "Bệnh viện Nha khoa The Wise"),
    "서울특별시장애인치과병원": ("Seoul Dental Hospital for People with Disabilities", "首尔市残疾人齿科医院", "ソウル市障害者歯科病院", "Bệnh viện Nha khoa cho Người khuyết tật Seoul"),
    "신촌다인치과병원": ("Sinchon Dain Dental Hospital", "新村多仁齿科医院", "新村ダイン歯科病院", "Bệnh viện Nha khoa Dain Sinchon"),
    "굿윌치과병원": ("Goodwill Dental Hospital", "古德威尔齿科医院", "グッドウィル歯科病院", "Bệnh viện Nha khoa Goodwill"),
    "이앤미치과병원": ("E&Me Dental Hospital", "E&Me齿科医院", "イーアンドミー歯科病院", "Bệnh viện Nha khoa E&Me"),
    "플란치과병원": ("Plan Dental Hospital", "普兰齿科医院", "プラン歯科病院", "Bệnh viện Nha khoa Plan"),
    "365드림치과병원": ("365 Dream Dental Hospital", "365梦想齿科医院", "365ドリーム歯科病院", "Bệnh viện Nha khoa 365 Dream"),
    "동래로덴치과병원": ("Dongnae Roden Dental Hospital", "东莱罗登齿科医院", "東萊ローデン歯科病院", "Bệnh viện Nha khoa Roden Dongnae"),
    "엘리트치과병원": ("Elite Dental Hospital", "精英齿科医院", "エリート歯科病院", "Bệnh viện Nha khoa Elite"),
    "오오치과병원": ("Oh Oh Dental Hospital", "OhOh齿科医院", "オオ歯科病院", "Bệnh viện Nha khoa Oh Oh"),
    "라인업치과병원": ("Line Up Dental Hospital", "Line Up齿科医院", "ラインアップ歯科病院", "Bệnh viện Nha khoa Line Up"),
    # 안과 병원
    "건양의료재단 김안과병원": ("Kim's Eye Hospital", "金眼科医院", "キム眼科病院", "Bệnh viện Mắt Kim's"),
    "누네안과병원": ("Nune Eye Hospital", "努内眼科医院", "ヌネ眼科病院", "Bệnh viện Mắt Nune"),
    "실로암안과병원": ("Siloam Eye Hospital", "西罗亚眼科医院", "シロアム眼科病院", "Bệnh viện Mắt Siloam"),
    "공안과병원": ("Kong Eye Hospital", "孔眼科医院", "コン眼科病院", "Bệnh viện Mắt Kong"),
    "혜민안과병원": ("Hyemin Eye Hospital", "惠民眼科医院", "ヘミン眼科病院", "Bệnh viện Mắt Hyemin"),
    "부산성모안과병원": ("Busan Sungmo Eye Hospital", "釜山圣母眼科医院", "釜山聖母眼科病院", "Bệnh viện Mắt Sungmo Busan"),
    "밝은눈안과병원": ("Balgeunnun Eye Hospital", "明眼眼科医院", "パルグンヌン眼科病院", "Bệnh viện Mắt Balgeunnun"),
    "정근안과병원": ("Jeonggeun Eye Hospital", "正根眼科医院", "チョングン眼科病院", "Bệnh viện Mắt Jeonggeun"),
    # 상급종합병원
    "재단법인아산사회복지재단 서울아산병원": ("Asan Medical Center", "首尔峨山医院", "ソウル峨山病院", "Trung tâm Y tế Asan"),
    "서울대학교병원": ("Seoul National University Hospital", "首尔大学医院", "ソウル大学校病院", "Bệnh viện Đại học Quốc gia Seoul"),
    "연세대학교의과대학세브란스병원": ("Severance Hospital, Yonsei University", "延世大学塞弗伦斯医院", "延世大学校セブランス病院", "Bệnh viện Severance, Đại học Yonsei"),
    "삼성서울병원": ("Samsung Medical Center", "三星首尔医院", "サムスンソウル病院", "Trung tâm Y tế Samsung"),
    "학교법인가톨릭학원가톨릭대학교서울성모병원": ("Seoul St. Mary's Hospital", "首尔圣母医院", "ソウル聖母病院", "Bệnh viện Seoul St. Mary's"),
    "학교법인 고려중앙학원 고려대학교의과대학부속병원(안암병원)": ("Korea University Anam Hospital", "高丽大学安岩医院", "高麗大学校安岩病院", "Bệnh viện Anam, Đại học Korea"),
    "연세대학교의과대학 강남세브란스병원": ("Gangnam Severance Hospital", "江南塞弗伦斯医院", "江南セブランス病院", "Bệnh viện Gangnam Severance"),
    "고려대학교의과대학부속구로병원": ("Korea University Guro Hospital", "高丽大学九老医院", "高麗大学校九老病院", "Bệnh viện Guro, Đại học Korea"),
    "경희대학교병원": ("Kyung Hee University Hospital", "庆熙大学医院", "慶熙大学校病院", "Bệnh viện Đại học Kyung Hee"),
    "한양대학교병원": ("Hanyang University Seoul Hospital", "汉阳大学医院", "漢陽大学校病院", "Bệnh viện Đại học Hanyang"),
    "이화여자대학교의과대학부속목동병원": ("Ewha Womans University Mokdong Hospital", "梨花女子大学木洞医院", "梨花女子大学校木洞病院", "Bệnh viện Mokdong, Đại học Nữ Ewha"),
    "건국대학교병원": ("Konkuk University Medical Center", "建国大学医院", "建国大学校病院", "Bệnh viện Đại học Konkuk"),
    "강북삼성병원": ("Kangbuk Samsung Hospital", "江北三星医院", "江北サムスン病院", "Bệnh viện Kangbuk Samsung"),
    "중앙대학교병원": ("Chung-Ang University Hospital", "中央大学医院", "中央大学校病院", "Bệnh viện Đại học Chung-Ang"),
    "부산대학교병원": ("Pusan National University Hospital", "釜山大学医院", "釜山大学校病院", "Bệnh viện Đại học Quốc gia Pusan"),
    "인제대학교부산백병원": ("Inje University Busan Paik Hospital", "仁济大学釜山白医院", "仁済大学校釜山白病院", "Bệnh viện Paik Busan, Đại học Inje"),
    "동아대학교병원": ("Dong-A University Hospital", "东亚大学医院", "東亜大学校病院", "Bệnh viện Đại học Dong-A"),
    "고신대학교복음병원": ("Kosin University Gospel Hospital", "高神大学福音医院", "高神大学校福音病院", "Bệnh viện Phúc Âm, Đại học Kosin"),
    # 한방병원
    "경희대학교한방병원": ("Kyung Hee University Korean Medicine Hospital", "庆熙大学韩方医院", "慶熙大学校韓方病院", "Bệnh viện Y học Hàn Quốc Đại học Kyung Hee"),
    "자생한방병원": ("Jaseng Hospital of Korean Medicine", "自生韩方医院", "自生韓方病院", "Bệnh viện Y học Hàn Quốc Jaseng"),
    "장덕한방병원": ("Jangdeok Korean Medicine Hospital", "长德韩方医院", "長徳韓方病院", "Bệnh viện Y học Hàn Quốc Jangdeok"),
    "강동경희대학교한방병원": ("Kyung Hee University Korean Medicine Hospital at Gangdong", "江东庆熙大学韩方医院", "江東慶熙大学校韓方病院", "Bệnh viện Y học Hàn Quốc Đại học Kyung Hee Gangdong"),
    "의료법인해인의료재단 동신한방병원": ("Dongshin Korean Medicine Hospital", "东信韩方医院", "東信韓方病院", "Bệnh viện Y học Hàn Quốc Dongshin"),
    "잠실자생한방병원": ("Jamsil Jaseng Hospital of Korean Medicine", "蚕室自生韩方医院", "蚕室自生韓方病院", "Bệnh viện Y học Hàn Quốc Jaseng Jamsil"),
    "목동자생한방병원": ("Mokdong Jaseng Hospital of Korean Medicine", "木洞自生韩方医院", "木洞自生韓方病院", "Bệnh viện Y học Hàn Quốc Jaseng Mokdong"),
    "보라매자생한방병원": ("Boramae Jaseng Hospital of Korean Medicine", "波拉梅自生韩方医院", "ボラメ自生韓方病院", "Bệnh viện Y học Hàn Quốc Jaseng Boramae"),
    "모커리한방병원": ("Mokhuri Korean Medicine Hospital", "木厚里韩方医院", "モコリ韓方病院", "Bệnh viện Y học Hàn Quốc Mokhuri"),
    "대전대학교서울한방병원": ("Daejeon University Seoul Korean Medicine Hospital", "大田大学首尔韩方医院", "大田大学校ソウル韓方病院", "Bệnh viện Y học Hàn Quốc Seoul, Đại học Daejeon"),
    "해운대자생한방병원": ("Haeundae Jaseng Hospital of Korean Medicine", "海云台自生韩方医院", "海雲台自生韓方病院", "Bệnh viện Y học Hàn Quốc Jaseng Haeundae"),
    "동의대부속 한방병원": ("Dong-eui University Korean Medicine Hospital", "东义大学附属韩方医院", "東義大学校附属韓方病院", "Bệnh viện Y học Hàn Quốc Đại học Dong-eui"),
    "부산자생한방병원": ("Busan Jaseng Hospital of Korean Medicine", "釜山自生韩方医院", "釜山自生韓方病院", "Bệnh viện Y học Hàn Quốc Jaseng Busan"),
    "서면더존한방병원": ("Seomyeon The Zone Korean Medicine Hospital", "西面The Zone韩方医院", "西面ザ・ゾーン韓方病院", "Bệnh viện Y học Hàn Quốc The Zone Seomyeon"),
    "연산당당한방병원": ("Yeonsan Dangdang Korean Medicine Hospital", "莲山堂堂韩方医院", "蓮山タンダン韓方病院", "Bệnh viện Y học Hàn Quốc Dangdang Yeonsan"),
    "부산태흥당한방병원": ("Busan Taeheungdang Korean Medicine Hospital", "釜山泰兴堂韩方医院", "釜山泰興堂韓方病院", "Bệnh viện Y học Hàn Quốc Taeheungdang Busan"),
    "태흥당한방병원": ("Taeheungdang Korean Medicine Hospital", "泰兴堂韩方医院", "泰興堂韓方病院", "Bệnh viện Y học Hàn Quốc Taeheungdang"),
    "메타한방병원": ("Meta Korean Medicine Hospital", "美塔韩方医院", "メタ韓方病院", "Bệnh viện Y học Hàn Quốc Meta"),
    "부산더존한방병원": ("Busan The Zone Korean Medicine Hospital", "釜山The Zone韩方医院", "釜山ザ・ゾーン韓方病院", "Bệnh viện Y học Hàn Quốc The Zone Busan"),
    "개금진한방병원": ("Gaegeumjin Korean Medicine Hospital", "开金真韩方医院", "ケグムジン韓方病院", "Bệnh viện Y học Hàn Quốc Gaegeumjin"),
}

# 명소 (TourAPI 국문 제목) → {en, zh, ja, vi}
SPOTS = {
    "명동": ("Myeong-dong", "明洞", "明洞", "Myeong-dong"),
    "경복궁": ("Gyeongbokgung Palace", "景福宫", "景福宮", "Cung Gyeongbokgung"),
    "남산서울타워": ("N Seoul Tower", "N首尔塔", "Nソウルタワー", "Tháp N Seoul"),
    "동대문디자인플라자(DDP)": ("Dongdaemun Design Plaza (DDP)", "东大门设计广场（DDP）", "東大門デザインプラザ（DDP）", "Dongdaemun Design Plaza (DDP)"),
    "북촌한옥마을": ("Bukchon Hanok Village", "北村韩屋村", "北村韓屋村", "Làng Hanok Bukchon"),
    "인사동": ("Insa-dong", "仁寺洞", "仁寺洞", "Insa-dong"),
    "롯데월드 어드벤처": ("Lotte World Adventure", "乐天世界探险", "ロッテワールド・アドベンチャー", "Lotte World Adventure"),
    "남대문시장": ("Namdaemun Market", "南大门市场", "南大門市場", "Chợ Namdaemun"),
    "창덕궁과 후원 [유네스코 세계유산]": ("Changdeokgung Palace (UNESCO World Heritage)", "昌德宫（世界文化遗产）", "昌徳宮（ユネスコ世界遺産）", "Cung Changdeokgung (Di sản Thế giới UNESCO)"),
    "청계천": ("Cheonggyecheon Stream", "清溪川", "清渓川", "Suối Cheonggyecheon"),
    "해운대해수욕장": ("Haeundae Beach", "海云台海水浴场", "海雲台海水浴場", "Bãi biển Haeundae"),
    "광안리해수욕장": ("Gwangalli Beach", "广安里海水浴场", "広安里海水浴場", "Bãi biển Gwangalli"),
    "부산 감천문화마을": ("Gamcheon Culture Village", "甘川文化村", "甘川文化村", "Làng văn hóa Gamcheon"),
    "부산 자갈치시장": ("Jagalchi Market", "札嘎其市场", "チャガルチ市場", "Chợ Jagalchi"),
    "해동용궁사": ("Haedong Yonggungsa Temple", "海东龙宫寺", "海東龍宮寺", "Chùa Haedong Yonggungsa"),
    "BIFF 광장": ("BIFF Square", "BIFF广场", "BIFF広場", "Quảng trường BIFF"),
    "태종대": ("Taejongdae", "太宗台", "太宗台", "Taejongdae"),
    "흰여울문화마을": ("Huinnyeoul Culture Village", "白浅滩文化村", "ヒンヨウル文化村", "Làng văn hóa Huinnyeoul"),
}

# 가격 항목 조각 → {zh, ja, vi} (영문은 import_public_data.ITEM_EN)
ITEM_SEGMENTS = {
    "치과임플란트(1치당)": ("种植牙（每颗）", "インプラント（1本あたり）", "Cấy ghép implant (mỗi răng)"),
    "크라운": ("牙冠", "クラウン", "Mão răng"),
    "치석제거": ("洗牙", "歯石除去", "Cạo vôi răng"),
    "전악": ("全口", "全顎", "toàn hàm"),
    "올세라믹": ("全瓷", "オールセラミック", "sứ toàn phần"),
    "기타": ("其他", "その他", "khác"),
    "레이저각막상피절삭성형술(라섹)": ("LASEK 激光上皮下角膜磨镶术", "LASEK（ラゼック）", "LASEK"),
    "레이저각막절삭성형술(라식)": ("LASIK 激光原位角膜磨镶术", "LASIK（レーシック）", "LASIK"),
    "조절성 인공수정체": ("高端人工晶体（多焦点/老花）", "多焦点・老眼用眼内レンズ", "Thủy tinh thể nhân tạo cao cấp (đa tiêu cự/lão thị)"),
    "약침술": ("药针", "薬鍼", "Dược châm"),
    "경혈": ("穴位", "経穴", "huyệt"),
    "추나요법": ("推拿疗法", "推拿療法（チュナ）", "Liệu pháp Chuna"),
    "단순추나": ("简单", "単純", "đơn giản"),
    "복잡추나": ("复杂", "複雑", "phức tạp"),
    "특수(탈구)추나": ("特殊（脱位）", "特殊（脱臼）", "đặc biệt (trật khớp)"),
    "특수(내장기, 두개천골)추나": ("特殊（内脏、颅骶）", "特殊（内臓・頭蓋仙骨）", "đặc biệt (nội tạng, sọ-cùng)"),
    "뇌 MRI": ("脑部MRI", "脳MRI", "MRI não"),
    "뇌혈관 MRA": ("脑血管MRA", "脳血管MRA", "MRA mạch máu não"),
    "갑상선 초음파": ("甲状腺超声", "甲状腺エコー", "Siêu âm tuyến giáp"),
    "유방 초음파": ("乳腺超声", "乳腺エコー", "Siêu âm vú"),
    "상복부 초음파 (간·담낭·췌장 등)": ("上腹部超声（肝、胆囊、胰腺等）", "上腹部エコー（肝臓・胆のう・膵臓など）",
                                   "Siêu âm bụng trên (gan, túi mật, tụy)"),
    "경동맥 초음파": ("颈动脉超声", "頸動脈エコー", "Siêu âm động mạch cảnh"),
}

# 병원 종류 → {zh, ja, vi}
HOSPITAL_TYPE = {
    "tertiary": ("上级综合医院", "上級総合病院", "Bệnh viện tuyến trên"),
    "general": ("综合医院", "総合病院", "Bệnh viện đa khoa"),
    "hospital": ("医院", "病院", "Bệnh viện"),
    "dental_hospital": ("齿科医院", "歯科病院", "Bệnh viện nha khoa"),
    "korean_medicine_hospital": ("韩方医院", "韓方病院", "Bệnh viện y học Hàn Quốc"),
}


def names(table, ko):
    """국문 → {ko, en, zh, ja, vi}. 표에 없으면 국문만."""
    row = table.get(ko)
    return {"ko": ko, **dict(zip(("en", "zh", "ja", "vi"), row))} if row else {"ko": ko}


def item_names(ko):
    """가격 항목 국문 표시명 '크라운 · Zirconia' → {zh, ja, vi} (조각 단위 치환, 모르는 조각은 그대로)."""
    out = {}
    for idx, lang in enumerate(("zh", "ja", "vi")):
        parts = []
        for part in ko.split(" · "):
            seg = ITEM_SEGMENTS.get(part)
            parts.append(seg[idx] if seg else part)
        out[lang] = " · ".join(parts)
    return out


# 외국인환자 유치기관 등록 의원: 국문 이름 → 영문 브랜드 (진료과 표기는 CLINIC_TYPE 으로 붙인다)
CLINIC_BRANDS = {
    # 피부과
    "닥터스피부과의원": "Doctors", "밴스피부과의원": "Vands", "역삼밴스피부과의원": "Vands Yeoksam",
    "임이석테마피부과의원": "Lim Yi-seok Theme", "갤러리아피부과의원": "Galleria", "더셀피부과의원": "The Cell",
    "더힐피부과의원": "The Heal", "리프톤피부과의원": "Lifton", "플로레스피부과의원": "Flores",
    "바노바기피부과의원": "Banobagi", "오가나셀 피부과의원 동대문점": "Organacell Dongdaemun",
    "고운세상김양제장봉석피부과의원": "Gowoon Sesang", "더퍼스트피부과의원 부산본점": "The First Busan",
    "뮬피부과의원": "Mul", "부산맥스웰피부과의원": "Busan Maxwell", "이현재피부과의원": "Lee Hyun-jae",
    "그대고운피부과의원": "Geudae Goun",
    # 성형외과
    "디에이성형외과의원": "DA", "에이비성형외과의원": "AB", "기린성형외과의원": "Girin", "나나성형외과의원": "Nana",
    "뷰성형외과의원": "View", "모제림성형외과의원": "Mojelim", "바노바기성형외과의원": "Banobagi", "팝성형외과의원": "Pop",
    "리엔장성형외과의원": "Lien Jang", "신상성형외과의원": "Shinsang", "우리성형외과의원": "Woori", "원진성형외과의원": "Wonjin",
    "하늘성형외과의원": "Haneul", "더페이스성형외과의원": "The Face", "리팅성형외과의원": "Liting", "뉴라인성형외과의원": "New Line",
    "케이성형외과의원": "K", "디자인성형외과의원": "Design",
    # 안과
    "비앤빛안과의원": "B&VIIT", "하늘안과의원": "Haneul", "서울신세계안과의원": "Seoul Shinsegae",
    "센트럴서울안과의원": "Central Seoul", "밝은눈안과의원": "Balgeunnun", "밝은성모안과의원": "Balgeun Sungmo",
    "서울밝은세상안과의원": "Seoul Balgeun Sesang", "지에스안과의원": "GS", "수정안과의원": "Sujeong", "누네빛안과의원": "Nunebit",
    "굿모닝백이안과의원": "Good Morning Baek-i", "굿모닝성모안과의원": "Good Morning Sungmo",
    # 정형외과
    "세종스포츠정형외과의원": "Sejong Sports", "서울선정형외과의원": "Seoul Sun", "더바름정형외과의원": "The Bareum",
    "강남본정형외과의원": "Gangnam Bon", "마포공덕튼튼본정형외과의원": "Mapo Gongdeok Teunteun Bon", "반듯한정형외과의원": "Bandeutan",
    "제애정형외과의원": "Jeae", "뉴본정형외과의원": "New Bone", "참좋은정형외과신경과의원": "Chamjoeun", "뉴젠정형외과의원": "Newgen",
    # 산부인과·여성의원
    "(의)성광의료재단 차여성의원": "CHA", "의)성광의료재단 마곡차여성의원": "CHA Magok", "포유문산부인과의원": "Poyumun",
    "고은빛산부인과의원": "Goeunbit", "최상산부인과의원": "Choesang", "청담산부인과의원": "Cheongdam", "나를위한산부인과의원": "For Me",
    "미래연여성의원": "Miraeyeon", "메디우먼여성의원": "Mediwoman", "로앤산부인과의원": "Roaen", "애플산부인과의원": "Apple",
    "센텀이룸여성의원": "Centum Iroom",
}

# 의원 종류 → (국문, 영문, 중문, 일문, 베트남어)
CLINIC_TYPE = {
    "dermatology": ("피부과 의원", "Dermatology Clinic", "皮肤科诊所", "皮膚科クリニック", "Phòng khám Da liễu"),
    "plastic_surgery": ("성형외과 의원", "Plastic Surgery Clinic", "整形外科诊所", "美容外科クリニック", "Phòng khám Phẫu thuật Thẩm mỹ"),
    "ophthalmology": ("안과 의원", "Eye Clinic", "眼科诊所", "眼科クリニック", "Phòng khám Mắt"),
    "orthopedics": ("정형외과 의원", "Orthopedic Clinic", "骨科诊所", "整形外科クリニック", "Phòng khám Chỉnh hình"),
    "fertility": ("산부인과·여성 의원", "Women's Clinic (OB/GYN)", "妇产科诊所", "産婦人科クリニック", "Phòng khám Sản phụ khoa"),
}


def clinic_names(ko, spec):
    """의원 국문 이름 → {ko, en, zh, ja, vi}. 영문 브랜드가 없으면 국문만."""
    brand = CLINIC_BRANDS.get(ko)
    if not brand:
        return {"ko": ko}
    _, en, zh, ja, vi = CLINIC_TYPE[spec]
    return {"ko": ko, "en": f"{brand} {en}", "zh": f"{brand}{zh}", "ja": f"{brand}{ja}", "vi": f"{vi} {brand}"}
