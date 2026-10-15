// 무료 한글 글꼴 모음 (fonts-archive: 눈누 · 기관 · 기업이 무료로 배포한 글꼴을 웹 글꼴로 정리한 저장소, 글꼴마다 README 에 사용 범위)
// 줄마다 "글꼴 이름(CSS font-family = 저장소 이름)|한글 이름|분류". CSS: https://cdn.jsdelivr.net/gh/fonts-archive/<이름>/<이름>.css
// 상업적 사용이 허용되지 않는 글꼴(애플 산돌고딕 Neo 등)은 넣지 않음. DOM 없음.
export const ARCHIVE_FONTS_TEXT = `
KERISBAEUM|케리스 배움체|디자인
HakgyoansimEohangkkumigiB|학교안심 어항꾸미기|손글씨
HakgyoansimJayeonR|학교안심 자연|손글씨
HakgyoansimJayusiganR|학교안심 자유시간|손글씨
HakgyoansimMondangbunfilR|학교안심 몽당분필|손글씨
HakgyoansimSiganpyoR|학교안심 시간표|손글씨
HakgyoansimSusukkangL|학교안심 수수깡|손글씨
HakgyoansimBandalL|학교안심 반달|손글씨
HakgyoansimYieodalligiL|학교안심 이어달리기|손글씨
HakgyoansimTteokbokkiB|학교안심 떡볶이|손글씨
HakgyoansimSketchbookR|학교안심 스케치북|손글씨
HakgyoansimSangjangR|학교안심 상장|손글씨
HakgyoansimMindeulleholssiR|학교안심 민들레홀씨|손글씨
HakgyoansimGongryongalR|학교안심 공룡알|손글씨
HakgyoansimChaekgalpiR|학교안심 책갈피|손글씨
HakgyoansimKidaripungseonL|학교안심 키다리풍선|손글씨
HakgyoansimOcarinaR|학교안심 오카리나|손글씨
HakgyoansimSaekyeonpilR|학교안심 색연필|손글씨
HakgyoansimManitoR|학교안심 마니또|손글씨
HakgyoansimNalgaeR|학교안심 날개|손글씨
HakgyoansimBoardmarkerR|학교안심 보드마카|손글씨
HakgyoansimPosterB|학교안심 포스터|손글씨
Cafe24PROSlimAir|카페24 PRO Slim Air|디자인
Cafe24PROSlimMax|카페24 PRO Slim Max|디자인
Cafe24PROSlimFit|카페24 PRO Slim Fit|디자인
Cafe24MeongiWhite|카페24 멍이 화이트|디자인
Cafe24NyangiBlack|카페24 냥이 블랙|디자인
Cafe24NyangiWhite|카페24 냥이 화이트|디자인
Cafe24MeongiBlack|카페24 멍이 블랙|디자인
Cafe24Angbanana|카페24 앙바나나|디자인
Cafe24PROUP|카페24 PRO UP|디자인
SacheonUju|사천우주체|디자인
SacheonHangGong|사천항공체|디자인
AstaSans|아스타산스|고딕
770Balsa|770 발사체|디자인
SahmyookUniversity|삼육대체|디자인
BMKkubulim|배민 꾸불림체|디자인
KCC-Hanbit|KCC한빛체|디자인
GasoekOne|가석체|디자인
GrandifloraOne|능소화|디자인
Diphylleia|산하엽|디자인
MoiraiOne|모이라이|디자인
BagelFatOne|베이글팻|디자인
KCCAhnchangho|KCC 안창호체|손글씨
Orbit|Orbit|디자인
Joongnajoche|중나좋체|디자인
YanoljaYache|야놀자 야체|디자인
UNPEOPLEGothic|유앤피플 고딕체|고딕
SolmoeKimDaeGeon|솔뫼 김대건체|디자인
FOUREYES|안경잡이체|디자인
AssacomFreeGothic|아싸컴 무료고딕|고딕
BookendBatang|북엔드 바탕|명조
STUNNING|스터닝 산스|고딕
KCCBangJeonghwan|KCC방정환체|손글씨
KCCHyerim|KCC혜림체|손글씨
NanumHuman|나눔휴먼|디자인
KakaoBigSans|카카오 큰글씨|고딕
KakaoSmallSans|카카오 작은글씨|고딕
DaehwaNanumeuiSori|대화나눔의 소리|디자인
DaehwaNanum|대화나눔 본문체|디자인
GoormSans|구름 산스|고딕
Paperlogy|페이퍼로지|고딕
HakgyoansimByeolbichhaneul|학교안심 별빛하늘|손글씨
HakgyoansimChilpanjiugae|학교안심 칠판지우개|손글씨
HakgyoansimPuzzleOutline|학교안심 퍼즐 Outline|손글씨
HakgyoansimPuzzleBlack|학교안심 퍼즐 Black|손글씨
HakgyoansimDunggeunmiso|학교안심 둥근미소|손글씨
HakgyoansimBadasseugi|학교안심 받아쓰기|손글씨
HakgyoansimByeoljari|학교안심 별자리|손글씨
HakgyoansimAllimjang|학교안심 알림장|손글씨
HakgyoansimChulseokbu|학교안심 출석부|손글씨
HakgyoansimNadeuri|학교안심 나들이|손글씨
HakgyoansimGeurimilgi|학교안심 그림일기|손글씨
HakgyoansimUsan|학교안심 우산|손글씨
HakgyoansimRikodeo|학교안심 리코더|손글씨
HakgyoansimUndongjang|학교안심 운동장|손글씨
HakgyoansimDotbogi|학교안심 돋보기|손글씨
HakgyoansimKkwabaegi|학교안심 꽈배기|손글씨
HakgyoansimSamulham|학교안심 사물함|손글씨
HakgyoansimKossuyeom|학교안심 콧수염|손글씨
HakgyoansimYeohaeng|학교안심 여행|손글씨
HakgyoansimEunhasu|학교안심 은하수|손글씨
HakgyoansimTtwimteul|학교안심 뜀틀|손글씨
HakgyoansimMoheomga|학교안심 모험가|손글씨
HakgyoansimKkokkoma|학교안심 꼬꼬마|손글씨
HakgyoansimMulgyeol|학교안심 물결|손글씨
HakgyoansimJiugae|학교안심 지우개|손글씨
HakgyoansimMonggeulmonggeul|학교안심 몽글몽글|손글씨
JSongMyung|J송명|디자인
JGaegujaengyi|J개구쟁이|디자인
J해바라기|J해바라기|디자인
SLEIGothic|서평원 꺾깍체|고딕
netmarble|넷마블체|디자인
NanumGothicCoding|나눔 고딕 코딩|코딩
Ownglyph_2022_UWY_Da_Gyeong|온글잎 의청수 다경체|손글씨
Ownglyph_2022_UWY_Si_Woo|온글잎 의청수 시우체|손글씨
Ownglyph_2022_UWY_Jae_Seon|온글잎 의청수 재선체|손글씨
Ownglyph_2022_UWY_Yoon_Yeong|온글잎 의청수 윤영체|손글씨
Ownglyph_2022_UWY_Joo_Bin|온글잎 의청수 주빈체|손글씨
NanumGaRamYeonGgoc|나눔손글씨 가람연꽃|손글씨
NanumGarMaesGeur|나눔손글씨 갈맷글|손글씨
NanumGangBuJangNimCe|나눔손글씨 강부장님체|손글씨
NanumGoDigANiGoGoDing|나눔손글씨 고딕 아니고 고딩|손글씨
NanumGoRyeoGeurGgor|나눔손글씨 고려글꼴|손글씨
NanumGomSinCe|나눔손글씨 곰신체|손글씨
NanumGyuRiEuiIrGi|나눔손글씨 규리의 일기|손글씨
NanumGeumEunBoHwa|나눔손글씨 금은보화|손글씨
NanumGiBbeumBarkEum|나눔손글씨 기쁨밝음|손글씨
NanumGimYuICe|나눔손글씨 김유이체|손글씨
NanumGgocNaeEum|나눔손글씨 꽃내음|손글씨
NanumGgeuTeuMeoRiCe|나눔손글씨 끄트머리체|손글씨
NanumNaNeunIGyeoNaenDa|나눔손글씨 나는 이겨낸다|손글씨
NanumNaMuJeongWeon|나눔손글씨 나무정원|손글씨
NanumNaEuiANaeSonGeurSsi|나눔손글씨 나의 아내 손글씨|손글씨
NanumNoRyeogHaNeunDongHeui|나눔손글씨 노력하는 동희|손글씨
NanumNeuRisNeuRisCe|나눔손글씨 느릿느릿체|손글씨
NanumDaSiSiJagHae|나눔손글씨 다시 시작해|손글씨
NanumDaJinCe|나눔손글씨 다진체|손글씨
NanumDaCaeSaRang|나눔손글씨 다채사랑|손글씨
NanumDaHaengCe|나눔손글씨 다행체|손글씨
NanumDarEuiGweDo|나눔손글씨 달의궤도|손글씨
NanumDaeGwangYuRi|나눔손글씨 대광유리|손글씨
NanumDaeHanMinGugYeorSaCe|나눔손글씨 대한민국 열사체|손글씨
NanumDongHwaDdoBag|나눔손글씨 동화또박|손글씨
NanumDungGeunInYeon|나눔손글씨 둥근인연|손글씨
NanumDdaDdeusHanJagByeor|나눔손글씨 따뜻한 작별|손글씨
NanumDdaAgDanDan|나눔손글씨 따악단단|손글씨
NanumDdarEGeEomMaGa|나눔손글씨 딸에게 엄마가|손글씨
NanumDdoBagDdoBag|나눔손글씨 또박또박|손글씨
NanumMaGoCe|나눔손글씨 마고체|손글씨
NanumMasIssNeunCe|나눔손글씨 맛있는체|손글씨
NanumMongDor|나눔손글씨 몽돌|손글씨
NanumMuGungHwa|나눔손글씨 무궁화|손글씨
NanumMuJinJangCe|나눔손글씨 무진장체|손글씨
NanumMiNiSonGeurSsi|나눔손글씨 미니 손글씨|손글씨
NanumMiRaeNaMu|나눔손글씨 미래나무|손글씨
NanumBaReunJeongSin|나눔손글씨 바른정신|손글씨
NanumBaReunHiPi|나눔손글씨 바른히피|손글씨
NanumAGiSaRangCe|나눔손글씨 아기사랑체|손글씨
NanumBanJjagBanJjagByeor|나눔손글씨 반짝반짝 별|손글씨
NanumBaeEunHyeCe|나눔손글씨 배은혜체|손글씨
NanumBaegEuiEuiCeonSa|나눔손글씨 백의의 천사|손글씨
NanumBeoDeuNaMu|나눔손글씨 버드나무|손글씨
NanumBeomSomCe|나눔손글씨 범솜체|손글씨
NanumBuJangNimNunCiCe|나눔손글씨 부장님 눈치체|손글씨
NanumBugGeugSeong|나눔손글씨 북극성|손글씨
NanumBiSangCe|나눔손글씨 비상체|손글씨
NanumBbangGuNiMamSonGeurSsi|나눔손글씨 빵구니맘 손글씨|손글씨
NanumSaRangHaeADeur|나눔손글씨 사랑해 아들|손글씨
NanumSangHaeCanMiCe|나눔손글씨 상해찬미체|손글씨
NanumSeongSirCe|나눔손글씨 성실체|손글씨
NanumSeGyeJeogInHanGeur|나눔손글씨 세계적인 한글|손글씨
NanumSeACe|나눔손글씨 세아체|손글씨
NanumSeHwaCe|나눔손글씨 세화체|손글씨
NanumSoMiCe|나눔손글씨 소미체|손글씨
NanumSoBangGwanEuiGiDo|나눔손글씨 소방관의 기도|손글씨
NanumSonPyeonJiCe|나눔손글씨 손편지체|손글씨
NanumSuJubEunDaeHagSaeng|나눔손글씨 수줍은 대학생|손글씨
NanumSiUGwiYeoWeo|나눔손글씨 시우 귀여워|손글씨
NanumSinHonBuBu|나눔손글씨 신혼부부|손글씨
NanumAReumDeuRiGgocNaMu|나눔손글씨 아름드리 꽃나무|손글씨
NanumABbaGeurSsi|나눔손글씨 아빠글씨|손글씨
NanumABbaEuiYeonAePyeonJi|나눔손글씨 아빠의 연애편지|손글씨
NanumAInMamSonGeurSsi|나눔손글씨 아인맘 손글씨|손글씨
NanumAJumMaJaYu|나눔손글씨 아줌마 자유|손글씨
NanumAnSsangCe|나눔손글씨 안쌍체|손글씨
NanumAmSeuTeReuDam|나눔손글씨 암스테르담|손글씨
NanumYaGeunHaNeunGimJuIm|나눔손글씨 야근하는 김주임|손글씨
NanumYaCaeJangSuBaegGeumRye|나눔손글씨 야채장수 백금례|손글씨
NanumEomMaSaRang|나눔손글씨 엄마사랑|손글씨
NanumEongGeongKwiCe|나눔손글씨 엉겅퀴체|손글씨
NanumYeoReumGeurSsi|나눔손글씨 여름글씨|손글씨
NanumYeonJiCe|나눔손글씨 연지체|손글씨
NanumYeorAHobEuiBanJjagIm|나눔손글씨 열아홉의 반짝임|손글씨
NanumYeorIrCe|나눔손글씨 열일체|손글씨
NanumYeDangCe|나눔손글씨 예당체|손글씨
NanumYeBbeunMinGyeongCe|나눔손글씨 예쁜 민경체|손글씨
NanumOgBiCe|나눔손글씨 옥비체|손글씨
NanumWaIrDeu|나눔손글씨 와일드|손글씨
NanumOeHarMeoNiGeurSsi|나눔손글씨 외할머니글씨|손글씨
NanumOenSonJabIDoYeBbeo|나눔손글씨 왼손잡이도 예뻐|손글씨
NanumURiDdarSonGeurSsi|나눔손글씨 우리딸 손글씨|손글씨
NanumYuNiDdingDdangDdingDdang|나눔손글씨 유니 띵땅띵땅|손글씨
NanumEuiMiIssNeunHanGeur|나눔손글씨 의미있는 한글|손글씨
NanumJaBuSimJiU|나눔손글씨 자부심지우|손글씨
NanumJarHaGoIssEo|나눔손글씨 잘하고 있어|손글씨
NanumJangMiCe|나눔손글씨 장미체|손글씨
NanumJeomGgorCe|나눔손글씨 점꼴체|손글씨
NanumJeongEunCe|나눔손글씨 정은체|손글씨
NanumJungHagSaeng|나눔손글씨 중학생|손글씨
NanumJinJuBagGyeongACe|나눔손글씨 진주 박경아체|손글씨
NanumCeorPirGeurSsi|나눔손글씨 철필글씨|손글씨
NanumCoDingHeuiMang|나눔손글씨 초딩희망|손글씨
NanumKarGugSu|나눔손글씨 칼국수|손글씨
NanumKoKoCe|나눔손글씨 코코체|손글씨
NanumHaNaDoeEoSonGeurSsi|나눔손글씨 하나되어 손글씨|손글씨
NanumHaNaSonGeurSsi|나눔손글씨 하나손글씨|손글씨
NanumHaRamCe|나눔손글씨 하람체|손글씨
NanumHanYunCe|나눔손글씨 한윤체|손글씨
NanumHarABeoJiEuiNaNum|나눔손글씨 할아버지의나눔|손글씨
NanumHaengBogHanDoBi|나눔손글씨 행복한 도비|손글씨
NanumHyeogICe|나눔손글씨 혁이체|손글씨
NanumHyeJunCe|나눔손글씨 혜준체|손글씨
NanumHyoNamNeurHwaITing|나눔손글씨 효남 늘 화이팅|손글씨
NanumHeuiMangNuRi|나눔손글씨 희망누리|손글씨
NanumHeuinGgoRiSuRi|나눔손글씨 흰꼬리수리|손글씨
NanumHimNaeRaNeunMarBoDan|나눔손글씨 힘내라는 말보단|손글씨
YUniverse|Y유니버스체|디자인
YSpotlight|Y 너만을 비춤체|디자인
YOONY|유니폰트|손글씨
YEONGJUPunggiGinseng|영주풍기인삼체|디자인
Yeongdo|영도체|디자인
YeongdeokSnowCrab|영덕대게체|디자인
YeongdeokSea|영덕바다체|디자인
YeongdeokHaeparang|영덕해파랑체|디자인
YeongdeokBlueroad|영덕블루로드체|디자인
Ydestreet|Y이드스트릿체|디자인
Ycomputer|Y콤퓨타체|픽셀
YClover|Y클로버체|디자인
yangjin|양진체|디자인
Wooridaum|우리다움체|디자인
wedrive|위드라이브체|디자인
WatermelonSalad|수박화체|디자인
Warhaven|넥슨 워헤이븐체|디자인
Wanjunuriche|완주누리체|디자인
Wanjudaedunsanche|완주대둔산체|디자인
WAKEUP|정신차렷체|디자인
WagleWagle|와글와글체|디자인
YeongamWolchulsan|영암월출산체|디자인
UlsanJunggu|울산중구체|디자인
UhBeeSehyun|어비 세현체|손글씨
TwaySky|티웨이 하늘체|디자인
TwayFly|티웨이 날다체|디자인
TwayAir|티웨이 항공체|디자인
TtangsBudaeJjigae|땅스부대찌개|디자인
TmonTium|티몬 티움체|디자인
TmonMonsori|티몬 몬소리체|디자인
TmoneyRoundWind|티머니 둥근바람|디자인
TheJamsil|더잠실체|디자인
TAEBAEKmilkyway|태백 은하수체|디자인
TAEBAEKfont|태백체|디자인
SuseongHyejeong|수성혜정체|디자인
SuseongDotum|수성돋움체|고딕
SuseongBatang|수성바탕체|명조
SUNMOONUNIVERSITY|선문대체|디자인
Suncheon|순천체|디자인
SUITE|스위트|고딕
SUIT|수트|고딕
SpoqaHanSansNeo|스포카 한 산스 네오|고딕
SpoqaHanSans|스포카 한 산스|고딕
SOYOMaple|소요단풍체|디자인
SOGANGUNIVERSITY|서강체|디자인
ShinDongYupHandwriting|신동엽 손글씨|손글씨
SFHambakSnow|함박눈체|손글씨
SeroPro|Sero Pro|디자인
SeoulNamsanCondensed|서울남산 장체|고딕
SeoulNamsan|서울남산체|고딕
SeoulHangangCondensed|서울한강 장체|고딕
SeoulHangang|서울한강체|고딕
SEOLLEIMcool|설레임체|디자인
SEOLEIMcool|시원한 설레임체|디자인
SejongHospital|세종병원체|디자인
SEBANGGothic|세방고딕|고딕
SDSamliphopangcheOutline|산돌 삼립호빵체 Outline|디자인
SDSamliphopangcheBasic|산돌 삼립호빵체 Basic|디자인
SANGJUDajungdagam|상주 다정다감체|디자인
SaenggeoJincheon|생거진천체|디자인
S-CoreDream|에스코어 드림|고딕
RIDIBatang|리디바탕|명조
Recipekorea|레코체|디자인
Rebecca|레베카체|디자인
PyeongChangPeace|평창평화체|디자인
PyeongChang|평창체|디자인
PuradakGentleGothic|푸라닭 젠틀 고딕|고딕
PilseungGothic|필승고딕|고딕
PartialSansKR|파셜산스|고딕
OKGUNG|읏맨 궁서체|디자인
OKDDUNG|읏맨 엉뚱체|디자인
OKCHAN|읏맨체|디자인
OGRenaissanceSecret|OG 르네상스 비밀|디자인
NULY|널리고딕|고딕
NPSfont|국민연금체|고딕
NotoSansKR|본고딕 (Noto Sans KR)|고딕
NMFClassic|국립박물관문화재단 클래식체|디자인
NEXONLv2Gothic|넥슨 Lv.2 고딕|고딕
NEXONLv1Gothic|넥슨 Lv.1 고딕|고딕
NEXONFootballGothic|넥슨 풋볼 고딕|고딕
NeoDunggeunmoPro|Neo 둥근모 Pro|픽셀
NanumSquareRound|나눔스퀘어 라운드|디자인
NanumSquareNeo|나눔스퀘어 네오|디자인
NanumSquare|나눔스퀘어|디자인
NanumPen|나눔손글씨 펜|손글씨
NanumMyeongjoEco|나눔명조 에코|명조
NanumMyeongjo-YetHangul|나눔명조 옛한글|명조
NanumGothicEco|나눔고딕 에코|고딕
NanumGothic|나눔고딕|고딕
NanumMyeongjo|나눔명조|명조
NanumGangInHanWiRo|나눔손글씨 강인한 위로|손글씨
NanumBrush|나눔손글씨 붓|손글씨
NanumBarunpen|나눔바른펜|디자인
NanumBarunGothic-YetHangul|나눔바른고딕 옛한글|고딕
NanumBarunGothic|나눔바른고딕|고딕
MonoplexKRWideNerdItalic|모노플렉스KR Wide Nerd Italic|코딩
MonoplexKRWideNerd|모노플렉스KR Wide Nerd|코딩
MonoplexKRWideItalic|모노플렉스KR Wide Italic|코딩
MonoplexKRWide|모노플렉스KR Wide|코딩
MonoplexKRNerdItalic|모노플렉스KR Nerd Italic|코딩
MonoplexKRNerd|모노플렉스KR Nerd|코딩
MonoplexKRItalic|모노플렉스KR Italic|코딩
MonoplexKR|모노플렉스KR|코딩
MBC1961Gulim|MBC 1961 굴림|디자인
MBC1961|MBC 1961|디자인
MaruBuri|마루 부리|명조
Maplestory|넥슨 메이플스토리|디자인
MACHO|마초체|디자인
MabinogiClassic|마비옛체|디자인
LOTTERIADDAG|롯데리아 딱붙어체|디자인
LOTTERIACHAB|롯데리아 촵땡겨체|디자인
LotteMartHappy|롯데마트 행복체|디자인
LotteMartDream|롯데마트 드림체|고딕
LINESeedKR|LINE Seed|고딕
Ladio|라디오고딕|고딕
KoPubDotum|KoPub 돋움체|고딕
KoPubBatang|KoPub 바탕체|명조
KOHINanum|KOHI 나눔|디자인
KOHIBaeum|KOHI 배움|디자인
KOFIHDrLEEJW|KOFIH이종욱체|디자인
KoddiUDOnGothic|KoddiUD 온고딕|고딕
Freesentation|프리젠테이션|고딕
KNUTRUTH|경북대 진리체|디자인
KNPSKkomi|국립공원 꼬미|디자인
KNPSBandali|국립공원 반달이|디자인
KIMM|한국기계연구원체|고딕
KimjungchulScript|김정철손글씨|손글씨
KimjungchulMyungjo|김정철명조|명조
KimjungchulGothic|김정철고딕|고딕
KCCSonkeechung|KCC 손기정체|손글씨
KCCPakKyongni|KCC 박경리체|손글씨
KCCMurukmuruk|KCC 무럭무럭체|손글씨
KCCKimhwanki|KCC 김환기체|손글씨
KCCKimhoon|KCC 김훈체|손글씨
KCCJeongbeom|KCC 정범체|손글씨
KCCImkwontaek|KCC 임권택체|손글씨
KCCGanpan|KCC 간판체|디자인
KCCEunyoung|KCC 은영체|손글씨
KCCDodamdodam|KCC 도담도담체|손글씨
KCCChassam|KCC 차쌤체|손글씨
KCCAhnjunggeun|KCC 안중근체|손글씨
KBODiaGothic|KBO 다이아고딕체|고딕
KartriderKor|넥슨 카트 고딕|고딕
JeonjuSoon|전주완판본 순체|디자인
JeonjuGak|전주완판본 각체|디자인
JeonjuCraftMj|전주공예명조체|명조
JeonjuCraftGo|전주공예고딕체|고딕
Jeongnimsaji|정림사지체|명조
JejuMyeongjo|제주명조|명조
JejuHallasan|제주한라산|디자인
JejuGothic|제주고딕|고딕
JejuDoldam|제주돌담체|디자인
JalnanGothic|여기어때 잘난체 고딕|고딕
Jalnan|여기어때 잘난체|디자인
Interop|Interop|코딩
IntelOneMono|Intel One Mono|코딩
InsungitCutelivelyJisu|귀염발랄체|손글씨
InkLipquid|잉크립퀴드체|디자인
InfinitySans|인피니티산스|고딕
IncheonEducationSotong|인천교육소통체|디자인
IncheonEducationSimin|인천교육시민체|디자인
IncheonEducationJaram|인천교육자람체|디자인
IncheonEducationHimchan|인천교육힘찬체|디자인
IBMPlexSansKR|IBM Plex Sans KR|고딕
IAMAPLAYER|I AM A PLAYER|디자인
HyunyGothic|혀니고딕|고딕
HyunyDororong|혀니도로롱|디자인
HSYuji|HS유지체|디자인
HSSummer|HS여름물빛체|디자인
HSSantokki|HS산토끼체|디자인
HSSaemaul|HS새마을체|디자인
HSJiptokkiRound|HS집토끼체 라운드|디자인
HSJiptokkiBlack|HS집토끼체 블랙|디자인
HSGyoulnoonkot2.0|HS겨울눈꽃체 2.0|손글씨
HSGyoulnoonkot|HS겨울눈꽃체|손글씨
HSGooltokki|HS굴토끼체|디자인
HSGaeulsenggak2.0|HS가을생각체 2.0|손글씨
HSGaeulsenggak|HS가을생각체|손글씨
HSDookubi|HS두꺼비체|디자인
HSBombaram3.0|HS봄바람체 3.0|손글씨
HSBombaram2.1|HS봄바람체 2.1|손글씨
HSBombaram2.0|HS봄바람체 2.0|손글씨
HoengseongHanu|횡성한우체|디자인
HiMelody|Yoon 하이멜로디|손글씨
HighSummit|High Summit|디자인
HBIOS-SYS|HBIOS-SYS|코딩
Harshita|Harshita|디자인
HappinessSansTitle|해피니스 산스 Title|고딕
HappinessSans|해피니스 산스|고딕
HangeulNuri|한글누리체|디자인
HangeulJaemin3|한글재민체 3.0|디자인
Handon3gyeopsal|한돈 삼겹살체|디자인
HANAMDAUM|하남다움체|디자인
HallymMyeongjo|한림명조체|명조
HallymGothic|한림고딕체|고딕
HakgyoansimWooju|학교안심 우주|손글씨
HakgyoansimTuho|학교안심 투호|손글씨
HakgyoansimSonagi|학교안심 소나기|손글씨
HakgyoansimSantteutdotum|학교안심 산뜻돋움|고딕
HakgyoansimSantteutbatang|학교안심 산뜻바탕|명조
HakgyoansimNamu|학교안심 나무|손글씨
HakgyoansimMalgeunnal|학교안심 맑은날|손글씨
HakgyoansimMabeopsa|학교안심 마법사|손글씨
HakgyoansimJeomsimsigan|학교안심 점심시간|손글씨
HakgyoansimGureum|학교안심 구름|손글씨
HakgyoansimGodeunjemok|학교안심 곧은제목|디자인
HakgyoansimGgooreogi|학교안심 꾸러기|손글씨
HakgyoansimGaeulsopung|학교안심 가을소풍|손글씨
HakgyoansimDoldam|학교안심 돌담|손글씨
HakgyoansimButpen|학교안심 붓펜|손글씨
HakgyoansimBunpil|학교안심 분필|손글씨
HakgyoansimBombanghak|학교안심 봄방학|손글씨
HakgyoansimBareondotum|학교안심 바른돋움|고딕
HakgyoansimBareonbatang|학교안심 바른바탕|명조
Hahmlet|함렛|디자인
Haenam|해남체|디자인
GyeonggiTitle|경기천년제목|디자인
GyeonggiBatang|경기천년바탕|명조
GwangyangTouching|광양감동체|디자인
GwangyangSunshine|광양햇살체|디자인
GOSEONGGEUMGANGNURI|고성금강누리체|디자인
GoryeongStrawberry|고령딸기체|디자인
GoodNeighbors|좋은이웃체|디자인
GmarketSans|G마켓 산스|고딕
GimpoTitle|김포제목체|디자인
GimpoGothic|김포고딕체|고딕
GimpoBatang|김포바탕체|명조
GIMJE|김제시체|디자인
GimhaeGaya|김해가야체|디자인
GiantsInline|자이언츠체 Inline|디자인
Giants|자이언츠체|디자인
Ghanachocolate|가나초콜릿체|디자인
GapyeongWave|가평물결체|디자인
GapyeongHanseokbong|가평한석봉체|디자인
GANGWONSTATE|강원특별자치도체|디자인
GangwonEduSaeeum|강원교육새음체|디자인
GangwonEduPower|강원교육튼튼체|디자인
GangwonEduModu|강원교육모두체|디자인
GangwonEduHyeonok|강원교육현옥샘체|디자인
GamtanRoadTantan|감탄로드 탄탄체|디자인
GamtanRoadGamtan|감탄로드 감탄체|디자인
GamtanRoadDotum|감탄로드 돋움체|고딕
GamtanRoadBatang|감탄로드 바탕체|명조
GamjaFlower|Yoon 감자꽃|손글씨
GabiaSolmee|가비아 솔미체|디자인
GabiaOndam|가비아 온담체|디자인
GabiaNapjakBlock|가비아 납작블럭체|디자인
GabiaMaeumgyeol|가비아 마음결체|디자인
GabiaHeuldot|가비아 흘돋체|디자인
GabiaCheongyeon|가비아 청연체|디자인
GabiaBombaram|가비아 봄바람체|디자인
FlowerSalt|꽃소금체|디자인
Firenze|피렌체|디자인
esamanru|이사만루|디자인
Environment|환경체|디자인
EliceDXNeolli|엘리스 DX널리체|디자인
EliceDigitalCoding|엘리스 디지털코딩체|코딩
EliceDigitalBaeum|엘리스 디지털배움체|디자인
EASTARJET|이스타체|디자인
DovemayoWild|거친둘기마요|디자인
DovemayoGothic|둘기마요고딕|고딕
Donoun|두넌체|디자인
DNFBitBit|던파 비트비트체|픽셀
Diary|다이어리체|손글씨
DeogonPrincessClassic|덕온공주옛체|디자인
DeogonPrincess|덕온공주체|디자인
DeltaUniverse|Delta Universe|디자인
Danjo|단조|디자인
DalseoHealing|달서힐링체|디자인
DalseoDarling|달서달링체|디자인
DAEAMLEETAEJOON|대암 이태준체|디자인
D2Coding|D2Coding|코딩
CucumberSalad|오이냉체|디자인
CROOKED|CROOKED|디자인
CookieRun|쿠키런|디자인
ClimateCrisisKR|기후위기 한글|디자인
ChusaLove|추사 사랑체|손글씨
ChungjuKimSaeng|충주김생체|디자인
chuckchuck|척척고딕|고딕
ChosunSm|조선신명조|명조
ChosunSg|조선가는고딕|고딕
ChosunLo|조선로고체|디자인
ChosunKm|조선굵은명조|명조
ChosunKg|조선굵은고딕|고딕
ChosunGu|조선굴림체|디자인
ChosunGs|조선궁서체|디자인
ChosunCentennial|조선100년체|디자인
ChosunBg|조선견고딕|고딕
Charter|Charter|명조
ChangwonDangamAsac|창원단감아삭체|디자인
CBNUJIKJI|충북대직지체|디자인
CBNU70YEARS|충북대70주년체|디자인
Cafe24Syongsyong|카페24 숑숑|디자인
Cafe24Supermagic|카페24 슈퍼매직|디자인
Cafe24SsurroundAir|카페24 써라운드 에어|디자인
Cafe24Ssurround|카페24 써라운드|디자인
Cafe24Ssukssuk|카페24 쑥쑥|디자인
Cafe24Simplehae|카페24 심플해|디자인
Cafe24Shiningstar|카페24 빛나는 별|디자인
Cafe24Oneprettynight|카페24 고운밤|디자인
Cafe24OhsquareAir|카페24 아네모네 에어|디자인
Cafe24Ohsquare|카페24 아네모네|디자인
Cafe24MoyamoyaFace|카페24 모야모야 Face|디자인
Cafe24Moyamoya|카페24 모야모야|디자인
Cafe24Lovingu|카페24 러빙유|디자인
Cafe24Dongdong|카페24 동동|디자인
Cafe24Decozoo|카페24 데코주|디자인
Cafe24Decoshadow|카페24 데코섀도|디자인
Cafe24Decoschool|카페24 데코스쿨|디자인
Cafe24Decomilk|카페24 데코밀크|디자인
Cafe24Decoline|카페24 데코라인|디자인
Cafe24Decobox|카페24 데코박스|디자인
Cafe24Danjunghae|카페24 단정해|디자인
Cafe24Dangdanghae|카페24 당당해|디자인
Cafe24ClassicType|카페24 클래식타입|디자인
Cafe24Behappy|카페24 비해피|디자인
BookkMyungjo|부크크 명조|명조
BookkGothic|부크크 고딕|고딕
BOKEH|BOKEH|디자인
BMYEONSUNG|배민 연성체|디자인
BMKIRANGHAERANG|배민 기랑해랑체|디자인
BMJUA|배민 주아체|디자인
BMHANNAPro|배민 한나체 Pro|디자인
Sam3|삼국지3 글꼴|디자인
Pretendard|프리텐다드|고딕
BMHANNAAir|배민 한나체 Air|디자인
Miraero|미래로글꼴|디자인
GalmuriMono9|갈무리모노9|코딩
GalmuriMono7|갈무리모노7|코딩
GalmuriMono11|갈무리모노11|코딩
Galmuri9|갈무리9|픽셀
Galmuri7|갈무리7|픽셀
Galmuri14|갈무리14|픽셀
Galmuri11|갈무리11|픽셀
BMHANNA11yrs|배민 한나는 열한살체|디자인
DungGeunMo|둥근모꼴|픽셀
DOSSaemmul|도스샘물|픽셀
DOSPilgi|도스필기|픽셀
DOSMyungjo|도스명조|픽셀
DOSIyagi|도스이야기|픽셀
DOSGothic|도스고딕|픽셀
DNFForgedBlade|던파 연단된 칼날|디자인
DNFBitBitv2|던파 비트비트체 v2|픽셀
dalmoori|달무리|픽셀
BMEuljiroOraeorae|배민 을지로오래오래체|디자인
BMEuljiro10yearslater|배민 을지로10년후체|디자인
BMEULJIRO|배민 을지로체|디자인
BMDOHYEON|배민 도현체|디자인
Bazzi|넥슨 배찌체|디자인
AritaDotum|아리따 돋움|고딕
AritaBuri|아리따 부리|명조
ARCHISCULPTURE|건축조각체|디자인
AONE|A1 전용서체|디자인
Aggro|어그로체|디자인
60sSTRIPE|60's STRIPE|디자인
116Watermelon|116수박화체|디자인
116Angmuburi|116앵무부리|명조
116AngdukHonesty|116앵덕정직|디자인
`;
