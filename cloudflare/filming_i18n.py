"""K-팝 뮤직비디오·드라마·영화 촬영지 (import_public_data.py filming 이 사용).

원자료: 한국문화정보원 「미디어콘텐츠 영상 촬영지 데이터」 (.import/raw/kcisa_filming_locations.csv, 2022-11 작성).
서울·부산의 공공장소(공원·궁·거리·다리·문화시설 등) 중 해외 팬에게 알려진 작품의 촬영지만 골랐다.
식당·카페·사유지는 폐업·사생활 문제로 제외. 장면 설명 국문은 원자료, 영문은 원자료를 옮긴 것이다.
"""

# 작품: 키 → (종류, 국문, 영문, 중문, 일문, 베트남어)
WORKS = {
    "bts_run": ("mv", "방탄소년단 'RUN'", "BTS — 'RUN'", "防弹少年团《RUN》", "BTS「RUN」", "BTS — 'RUN'"),
    "psy_gentleman": ("mv", "싸이 '젠틀맨'", "PSY — 'Gentleman'", "PSY《Gentleman》", "PSY「Gentleman」", "PSY — 'Gentleman'"),
    "psy_daddy": ("mv", "싸이 '대디'", "PSY — 'Daddy'", "PSY《Daddy》", "PSY「Daddy」", "PSY — 'Daddy'"),
    "svt_call": ("mv", "세븐틴 'CALL CALL CALL!'", "SEVENTEEN — 'CALL CALL CALL!'", "SEVENTEEN《CALL CALL CALL!》",
                 "SEVENTEEN「CALL CALL CALL!」", "SEVENTEEN — 'CALL CALL CALL!'"),
    "gfriend_megustas": ("mv", "여자친구 '오늘부터 우리는'", "GFRIEND — 'Me Gustas Tu'", "GFRIEND《今天开始我们是》",
                         "GFRIEND「今日から私たちは」", "GFRIEND — 'Me Gustas Tu'"),
    "gfriend_summerrain": ("mv", "여자친구 '여름비'", "GFRIEND — 'Summer Rain'", "GFRIEND《夏雨》", "GFRIEND「夏の雨」",
                           "GFRIEND — 'Summer Rain'"),
    "apink_remember": ("mv", "에이핑크 'Remember'", "Apink — 'Remember'", "Apink《Remember》", "Apink「Remember」", "Apink — 'Remember'"),
    "iu_night": ("mv", "아이유 '밤편지'", "IU — 'Through the Night'", "IU《夜信》", "IU「夜の手紙」", "IU — 'Through the Night'"),
    "akmu_lastgoodbye": ("mv", "악뮤 '오랜 날 오랜 밤'", "AKMU — 'Last Goodbye'", "AKMU《许久的日子许久的夜晚》",
                         "AKMU「長い日、長い夜」", "AKMU — 'Last Goodbye'"),
    "nct_cure": ("mv", "유영진×태용 '함께 (Cure)'", "Yoo Young-jin × TAEYONG — 'Cure'", "刘英镇×泰容《Cure》",
                 "ユ・ヨンジン×テヨン「Cure」", "Yoo Young-jin × TAEYONG — 'Cure'"),
    "baekhyun_park": ("mv", "백현 '놀이공원'", "BAEKHYUN — 'Amusement Park'", "伯贤《游乐园》", "ベッキョン「Amusement Park」",
                      "BAEKHYUN — 'Amusement Park'"),
    "jaypark_gang": ("mv", "박재범 '깡' 리믹스", "Jay Park — 'GANG' remix", "朴宰范《GANG》Remix", "パク・ジェボム「GANG」Remix",
                     "Jay Park — 'GANG' remix"),
    "lovelyz_candy": ("mv", "러블리즈 'Candy Jelly Love'", "Lovelyz — 'Candy Jelly Love'", "Lovelyz《Candy Jelly Love》",
                      "Lovelyz「Candy Jelly Love」", "Lovelyz — 'Candy Jelly Love'"),
    "exo_lmr": ("mv", "엑소 'Love Me Right'", "EXO — 'Love Me Right'", "EXO《Love Me Right》", "EXO「Love Me Right」", "EXO — 'Love Me Right'"),
    "twice_ooh": ("mv", "트와이스 '우아하게'", "TWICE — 'Like OOH-AHH'", "TWICE《Like OOH-AHH》", "TWICE「Like OOH-AHH」",
                  "TWICE — 'Like OOH-AHH'"),
    "goblin": ("drama", "도깨비", "Goblin (Guardian: The Lonely and Great God)", "鬼怪", "トッケビ", "Yêu tinh (Goblin)"),
    "woo": ("drama", "이상한 변호사 우영우", "Extraordinary Attorney Woo", "非常律师禹英雨", "ウ・ヨンウ弁護士は天才肌",
            "Nữ luật sư kỳ lạ Woo Young Woo"),
    "star": ("drama", "별에서 온 그대", "My Love from the Star", "来自星星的你", "星から来たあなた", "Vì sao đưa anh tới"),
    "itaewon": ("drama", "이태원 클라쓰", "Itaewon Class", "梨泰院Class", "梨泰院クラス", "Tầng lớp Itaewon"),
    "vincenzo": ("drama", "빈센조", "Vincenzo", "黑道律师文森佐", "ヴィンチェンツォ", "Vincenzo"),
    "delluna": ("drama", "호텔 델루나", "Hotel Del Luna", "德鲁纳酒店", "ホテルデルーナ", "Khách sạn Ánh Trăng (Hotel Del Luna)"),
    "theking": ("drama", "더 킹: 영원의 군주", "The King: Eternal Monarch", "国王：永远的君主", "ザ・キング：永遠の君主", "Quân vương bất diệt"),
    "startup": ("drama", "스타트업", "Start-Up", "启动了", "スタートアップ：夢の扉", "Khởi nghiệp (Start-Up)"),
    "cloy": ("drama", "사랑의 불시착", "Crash Landing on You", "爱的迫降", "愛の不時着", "Hạ cánh nơi anh"),
    "2521": ("drama", "스물다섯 스물하나", "Twenty-Five Twenty-One", "二十五，二十一", "二十五、二十一", "Hai mươi lăm, hai mươi mốt"),
    "mymister": ("drama", "나의 아저씨", "My Mister", "我的大叔", "マイ・ディア・ミスター", "Ông chú của tôi"),
    "truebeauty": ("drama", "여신강림", "True Beauty", "女神降临", "女神降臨", "Vẻ đẹp đích thực"),
    "kingdom": ("drama", "킹덤", "Kingdom", "王国", "キングダム", "Vương triều xác sống"),
    "dots": ("drama", "태양의 후예", "Descendants of the Sun", "太阳的后裔", "太陽の末裔", "Hậu duệ mặt trời"),
    "m1987": ("movie", "1987", "1987: When the Day Comes", "1987：黎明到来的那一天", "1987、ある闘いの真実", "1987: When the Day Comes"),
    "priests": ("movie", "검은 사제들", "The Priests", "黑祭司们", "プリースト 悪魔を葬る者", "The Priests"),
    "confidential": ("movie", "공조", "Confidential Assignment", "共助", "コンフィデンシャル/共助", "Confidential Assignment"),
    "masquerade": ("movie", "광해, 왕이 된 남자", "Masquerade", "双面君王", "王になった男", "Masquerade"),
    "host": ("movie", "괴물", "The Host", "汉江怪物", "グエムル -漢江の怪物-", "Quái vật sông Hàn (The Host)"),
    "burning": ("movie", "버닝", "Burning", "燃烧", "バーニング 劇場版", "Burning"),
    "handmaiden": ("movie", "아가씨", "The Handmaiden", "小姐", "お嬢さん", "Người hầu gái (The Handmaiden)"),
    "assassination": ("movie", "암살", "Assassination", "暗杀", "暗殺", "Ám sát (Assassination)"),
    "attorney": ("movie", "변호인", "The Attorney", "辩护人", "弁護人", "Luật sư (The Attorney)"),
    "nowhere": ("movie", "아저씨", "The Man from Nowhere", "大叔", "アジョシ", "Người vô danh (The Man from Nowhere)"),
    "sunny": ("movie", "써니", "Sunny", "阳光姐妹淘", "サニー 永遠の仲間たち", "Sunny"),
    "wedding": ("movie", "너의 결혼식", "On Your Wedding Day", "你的婚礼", "君の結婚式", "On Your Wedding Day"),
    "real": ("movie", "리얼", "Real", "Real", "リアル", "Real"),
    "willyou": ("movie", "당신, 거기 있어줄래요", "Will You Be There?", "你会在那里吗", "あなた、そこにいてくれますか", "Will You Be There?"),
}

# 장소: (원자료 장소명, 지역, 영문, 중문, 일문, 베트남어, [(원자료 연번, 작품 키, 영문 장면 설명)])
PLACES = [
    ("반포한강공원", "seoul", "Banpo Hangang Park", "盘浦汉江公园", "盤浦漢江公園", "Công viên sông Hàn Banpo", [
        (5327, "bts_run", "Scenes from BTS's 'RUN' music video were filmed here."),
        (13896, "delluna", "Where the young Chan-seong and his father meet the goddess Mago.")]),
    ("서울도서관", "seoul", "Seoul Library", "首尔图书馆", "ソウル図書館", "Thư viện Seoul", [
        (9846, "psy_gentleman", "PSY dances in front of this building in 'Gentleman'.")]),
    ("수영만 요트 경기장", "busan", "Busan Yachting Center (Suyeong Bay)", "水营湾游艇竞技场", "水営湾ヨット競技場",
     "Bến du thuyền vịnh Suyeong", [(9842, "psy_daddy", "A filming location of PSY's 'Daddy' music video.")]),
    ("동작대교", "seoul", "Dongjak Bridge", "铜雀大桥", "銅雀大橋", "Cầu Dongjak", [
        (8214, "svt_call", "Seen in SEVENTEEN's 'CALL CALL CALL!' music video."),
        (6600, "vincenzo", "Jang Han-seok talks to Choi Myung-hee on this bridge in the finale."),
        (1235, "host", "Where the monster appears and the family fights it.")]),
    ("올림픽공원 나홀로나무", "seoul", "Olympic Park 'Lone Tree'", "奥林匹克公园孤独树", "オリンピック公園「ひとりぼっちの木」",
     "Cây cô đơn, Công viên Olympic", [(10691, "gfriend_megustas", "Filming location of GFRIEND's 'Me Gustas Tu'.")]),
    ("북촌 한옥마을", "seoul", "Bukchon Hanok Village", "北村韩屋村", "北村韓屋村", "Làng Hanok Bukchon", [
        (10693, "gfriend_summerrain", "Filming location of GFRIEND's 'Summer Rain'.")]),
    ("광안대교", "busan", "Gwangan Bridge", "广安大桥", "広安大橋", "Cầu Gwangan", [
        (10529, "apink_remember", "The bridge in Apink's 'Remember' music video.")]),
    ("송정해수욕장", "busan", "Songjeong Beach", "松亭海水浴场", "松亭海水浴場", "Bãi biển Songjeong", [
        (2892, "willyou", "Where the young Su-hyeon and Tae-ho talk by the sea, 30 years earlier.")]),
    ("문화공감수정", "busan", "Munhwa Gonggam Sujeong (historic house)", "文化共感水晶（历史建筑）", "文化共感スジョン（近代建築）",
     "Nhà cổ Munhwa Gonggam Sujeong", [
        (10053, "iu_night", "The café-house where IU filmed 'Through the Night'."),
        (10163, "akmu_lastgoodbye", "Filming location of AKMU's 'Last Goodbye'."),
        (63, "m1987", "Appears in '1987' as a café near Busan Station.")]),
    ("망원한강공원", "seoul", "Mangwon Hangang Park", "望远汉江公园", "望遠漢江公園", "Công viên sông Hàn Mangwon", [
        (14594, "nct_cure", "Filming location of Yoo Young-jin × TAEYONG's 'Cure'.")]),
    ("롯데월드", "seoul", "Lotte World", "乐天世界", "ロッテワールド", "Lotte World", [
        (5969, "baekhyun_park", "BAEKHYUN filmed 'Amusement Park' in front of the carousel here."),
        (10627, "truebeauty", "Han Seo-jun takes Lim Ju-kyung ice skating here.")]),
    ("문래 창작촌", "seoul", "Mullae Art Village", "文来创作村", "文来創作村", "Làng nghệ thuật Mullae", [
        (5145, "jaypark_gang", "Jay Park filmed his 'GANG' remix video here."),
        (10108, "nowhere", "Where Won Bin calls Kim Sae-ron to come out.")]),
    ("배재학당역사박물관", "seoul", "Appenzeller/Pai Chai Hakdang History Museum", "培材学堂历史博物馆", "培材学堂歴史博物館",
     "Bảo tàng lịch sử Pai Chai Hakdang", [(3453, "lovelyz_candy", "Filming location of Lovelyz's 'Candy Jelly Love'.")]),
    ("용마랜드", "seoul", "Yongma Land", "龙马乐园", "龍馬ランド", "Yongma Land", [
        (10581, "exo_lmr", "Filming location of EXO's 'Love Me Right' video and album."),
        (13113, "twice_ooh", "Filming location of TWICE's 'Like OOH-AHH'.")]),
    ("덕수궁 돌담길", "seoul", "Deoksugung Stone Wall Road", "德寿宫石墙路", "徳寿宮の石垣道", "Đường tường đá Deoksugung", [
        (3113, "goblin", "The Grim Reaper's office sits behind this palace wall."),
        (11631, "woo", "Woo Young-woo and Lee Jun-ho's evening date in episode 10."),
        (6070, "star", "Do Min-joon walks along the wall in episode 7.")]),
    ("광화문", "seoul", "Gwanghwamun Gate", "光化门", "光化門", "Cổng Gwanghwamun", [
        (3112, "goblin", "The gate Kim Shin and Ji Eun-tak walk through after returning from Canada.")]),
    ("감고당길", "seoul", "Gamgodang-gil", "感古堂路", "感古堂通り", "Phố Gamgodang", [
        (3106, "goblin", "The stone-wall lane where Gong Yoo and Kim Go-eun first meet.")]),
    ("녹사평 육교", "seoul", "Noksapyeong Overpass", "绿莎坪天桥", "緑莎坪歩道橋", "Cầu vượt Noksapyeong", [
        (11685, "itaewon", "The main backdrop of Itaewon Class.")]),
    ("이태원세계음식거리", "seoul", "Itaewon World Food Street", "梨泰院世界美食街", "梨泰院世界グルメ通り",
     "Phố ẩm thực thế giới Itaewon", [
        (11686, "itaewon", "Where Park Sae-ro-yi runs into Oh Soo-ah."),
        (1189, "confidential", "Chase scenes were filmed on this food street.")]),
    ("청계상가", "seoul", "Cheonggye Sangga (Sewoon Arcade)", "清溪商街", "清渓商街", "Khu thương mại Cheonggye", [
        (6613, "vincenzo", "The building appears as 'Geumga Plaza'."),
        (10647, "truebeauty", "Lim Ju-kyung goes up to this rooftop in episode 1.")]),
    ("서울로7017", "seoul", "Seoullo 7017", "首尔路7017", "ソウル路7017", "Seoullo 7017", [
        (6614, "vincenzo", "Vincenzo meets lawyer Hong Yu-chan here in episode 3.")]),
    ("한국의 집", "seoul", "Korea House", "韩国之家", "韓国の家", "Korea House", [
        (6616, "vincenzo", "The Babel Group leaders meet here in episode 9.")]),
    ("정독도서관", "seoul", "Jeongdok Public Library", "正读图书馆", "正読図書館", "Thư viện Jeongdok", [
        (13917, "delluna", "Where the eyeless ghost had her accident.")]),
    ("동백섬", "busan", "Dongbaekseom Island", "冬柏岛", "冬柏島", "Đảo Dongbaek", [
        (3022, "theking", "Home of the Corea Empire's royal palace in the drama.")]),
    ("영화의전당", "busan", "Busan Cinema Center", "电影殿堂", "映画の殿堂", "Trung tâm Điện ảnh Busan", [
        (3023, "theking", "King Lee Gon rides past the building with his mounted guard.")]),
    ("다대포 해수욕장", "busan", "Dadaepo Beach", "多大浦海水浴场", "多大浦海水浴場", "Bãi biển Dadaepo", [
        (3017, "theking", "Where Lee Gon rides his horse Maximus.")]),
    ("광화문 광장", "seoul", "Gwanghwamun Square", "光化门广场", "光化門広場", "Quảng trường Gwanghwamun", [
        (3033, "theking", "Lee Gon and Jeong Tae-eul meet and suddenly embrace here.")]),
    ("노들섬", "seoul", "Nodeul Island", "鹭得岛", "ノドゥル島", "Đảo Nodeul", [
        (8682, "startup", "Home of the start-up hub 'Sandbox', the drama's main set."),
        (6599, "vincenzo", "Han Seung-hyuk meets Jang Han-seo here in episode 13.")]),
    ("코모도호텔 부산", "busan", "Commodore Hotel Busan", "釜山科莫多酒店", "コモドホテル釜山", "Khách sạn Commodore Busan", [
        (6715, "cloy", "Appears as an upscale hotel in Pyongyang."),
        (3765, "real", "A main backdrop of the film 'Real'.")]),
    ("마로니에 공원", "seoul", "Marronnier Park", "马罗尼埃公园", "マロニエ公園", "Công viên Marronnier", [
        (8620, "2521", "Na Hee-do's chat friend waits here with a yellow rose in episode 8.")]),
    ("N서울타워", "seoul", "N Seoul Tower", "N首尔塔", "Nソウルタワー", "Tháp N Seoul", [
        (6064, "star", "Kim Soo-hyun and Jun Ji-hyun hang a love lock here."),
        (10629, "truebeauty", "Lee Su-ho and Han Seo-jun come to the foot of the tower."),
        (5974, "burning", "The landmark always seen from Jong-su's room window.")]),
    ("국립극장", "seoul", "National Theater of Korea", "国立剧场", "国立劇場", "Nhà hát Quốc gia Hàn Quốc", [
        (6072, "star", "Where Do Min-joon returns and kisses Cheon Song-yi.")]),
    ("남산 케이블카", "seoul", "Namsan Cable Car", "南山缆车", "南山ケーブルカー", "Cáp treo Namsan", [
        (10645, "truebeauty", "Lim Ju-kyung rides the cable car to see the first snow.")]),
    ("낙산공원", "seoul", "Naksan Park", "骆山公园", "駱山公園", "Công viên Naksan", [
        (1980, "mymister", "Where the grandmother is shown the moon from a cart."),
        (10280, "assassination", "One of Seoul's best night views, seen in 'Assassination'."),
        (1190, "confidential", "Where Yoo Hae-jin's lie is uncovered.")]),
    ("창덕궁", "seoul", "Changdeokgung Palace", "昌德宫", "昌徳宮", "Cung Changdeokgung", [
        (12940, "kingdom", "Crown Prince Lee Chang stands in a pavilion here."),
        (2381, "wedding", "Kim Young-kwang and Park Bo-young's date.")]),
    ("창경궁", "seoul", "Changgyeonggung Palace", "昌庆宫", "昌慶宮", "Cung Changgyeonggung", [
        (12941, "kingdom", "Appears as the queen's residence.")]),
    ("경복궁", "seoul", "Gyeongbokgung Palace", "景福宫", "景福宮", "Cung Gyeongbokgung", [
        (12939, "kingdom", "Used as royal hideaways in the series."),
        (1212, "masquerade", "Lee Byung-hun escapes with the queen holding her hand.")]),
    ("가톨릭대학교 서울성모병원", "seoul", "Seoul St. Mary's Hospital", "首尔圣母医院", "ソウル聖母病院", "Bệnh viện Seoul St. Mary's", [
        (12991, "dots", "The hospital where Dr. Kang Mo-yeon works.")]),
    ("천주교 서울대교구 주교좌명동대성당", "seoul", "Myeongdong Cathedral", "明洞圣堂", "明洞聖堂", "Nhà thờ chính tòa Myeongdong", [
        (66, "m1987", "Where Father Kim Seung-hoon reads the statement."),
        (1002, "priests", "Gang Dong-won fetches the sacred objects here.")]),
    ("문화역서울284", "seoul", "Culture Station Seoul 284", "文化站首尔284", "文化駅ソウル284", "Ga văn hóa Seoul 284", [
        (10283, "assassination", "The former station building, now an exhibition space."),
        (9915, "handmaiden", "The door through which Sook-hee and Hideko escape.")]),
    ("서대문형무소역사관", "seoul", "Seodaemun Prison History Hall", "西大门刑务所历史馆", "西大門刑務所歴史館",
     "Bảo tàng nhà tù Seodaemun", [
        (9914, "handmaiden", "Appears as the mental hospital."),
        (6590, "vincenzo", "Kang Ho-cheol is released from this building in episode 16.")]),
    ("흰여울문화마을", "busan", "Huinnyeoul Culture Village", "白浅滩文化村", "ヒンヨウル文化村", "Làng văn hóa Huinnyeoul", [
        (6015, "attorney", "The 1980s backdrop of 'The Attorney'.")]),
    ("초량 이바구길", "busan", "Choryang Ibagu-gil", "草梁故事路", "草梁イバグキル", "Phố Choryang Ibagu", [
        (10106, "nowhere", "The alley Won Bin walks to his pawnshop.")]),
    ("정동길", "seoul", "Jeongdong-gil", "贞洞路", "貞洞通り", "Phố Jeongdong", [
        (9894, "sunny", "Where the two Na-mis comfort each other and meet.")]),
    ("백인제 가옥", "seoul", "Baek In-je House", "白麟济家屋", "白麟済家屋", "Nhà cổ Baek In-je", [
        (10281, "assassination", "The house of Kang In-guk and Mitsuko.")]),
    ("종묘", "seoul", "Jongmyo Shrine", "宗庙", "宗廟", "Đền Jongmyo", [
        (1213, "masquerade", "The opening scene of 'Masquerade'.")]),
]


# 관광공사(TourAPI)에 공공누리 사진이 없는 곳: 위키미디어 공용의 자유 이용 사진 (url, 저작자·라이선스, 원본 페이지)
PHOTO_OVERRIDES = {
    "동작대교": ("https://upload.wikimedia.org/wikipedia/commons/thumb/8/8d/Dongjak_Bridge.jpg/1280px-Dongjak_Bridge.jpg",
             "Bart0278 · Public domain", "https://commons.wikimedia.org/wiki/File:Dongjak_Bridge.jpg"),
    "용마랜드": ("https://upload.wikimedia.org/wikipedia/commons/thumb/e/ee/Yongma_Land.jpg/1280px-Yongma_Land.jpg",
             "Christian Bolz · CC BY-SA 4.0", "https://commons.wikimedia.org/wiki/File:Yongma_Land.jpg"),
    "가톨릭대학교 서울성모병원": ("https://upload.wikimedia.org/wikipedia/commons/thumb/3/33/%EA%B0%80%ED%86%A8%EB%A6%AD%EB%8C%80%ED%95%99%EA%B5%90_%EC%84%9C%EC%9A%B8%EC%84%B1%EB%AA%A8%EB%B3%91%EC%9B%90.jpg/1280px-%EA%B0%80%ED%86%A8%EB%A6%AD%EB%8C%80%ED%95%99%EA%B5%90_%EC%84%9C%EC%9A%B8%EC%84%B1%EB%AA%A8%EB%B3%91%EC%9B%90.jpg",
                       "Aspere · CC0", "https://commons.wikimedia.org/wiki/File:%EA%B0%80%ED%86%A8%EB%A6%AD%EB%8C%80%ED%95%99%EA%B5%90_%EC%84%9C%EC%9A%B8%EC%84%B1%EB%AA%A8%EB%B3%91%EC%9B%90.jpg"),
    "문화역서울284": ("https://upload.wikimedia.org/wikipedia/commons/thumb/f/f4/Culture_Station_Seoul_284_%28_2014_08_05_%29.jpg/1280px-Culture_Station_Seoul_284_%28_2014_08_05_%29.jpg",
                 "HaulS · CC0", "https://commons.wikimedia.org/wiki/File:Culture_Station_Seoul_284_(_2014_08_05_).jpg"),
    # 육교 자체 사진은 없어 바로 아래 녹사평역 사진을 쓴다
    "녹사평 육교": ("https://upload.wikimedia.org/wikipedia/commons/7/7a/Noksa02.jpg", "Noksapyeong Stn. · Marcopolis · Public domain", "https://commons.wikimedia.org/wiki/File:Noksa02.jpg"),
}
