/* KV 아틀라스 데이터
 * SOURCES: [이름, URL, 분류, 지역, 비용, 검색 템플릿, KV 활용도(1-3), 설명]
 *   검색 템플릿: '{q}'가 들어간 URL / 'site' = 구글 이미지에서 해당 도메인으로 한정 검색 / null = 검색 미지원
 * KV: 광고·포스터 키비주얼 레퍼런스. 이미지는 저작권 때문에 싣지 않고 원출처 검색으로 연결한다.
 *   col 은 대표 색 인상(근사치)이고, 연도·제작사는 공개 자료 기준 대표 정보다.
 */
window.KV_DATA = (function () {
  const CATEGORIES = [
    ['portfolio', '포트폴리오 플랫폼', '디자이너·스튜디오가 직접 올린 프로젝트. 광고 KV 원본과 제작 과정을 보기 좋다.'],
    ['engine', '이미지 검색엔진', '키워드로 웹 전체 이미지를 훑는다. 지역별 엔진마다 결과 성향이 다르다.'],
    ['reverse', '역이미지 검색', '이미지로 출처·고해상도 원본·유사 이미지를 찾는다.'],
    ['curation', '무드보드·큐레이션', '사람이 고른 이미지 모음. 감도 높은 결과가 많다.'],
    ['ad-archive', '광고 아카이브·어워드', '집행된 캠페인과 수상작. 아이디어와 크레딧까지 확인 가능.'],
    ['ad-library', '집행 광고 라이브러리', '지금 실제로 돌고 있는 광고 소재를 플랫폼별로 조회.'],
    ['magazine', '디자인 매거진·블로그', '편집자가 고른 신작과 해설. 트렌드 파악용.'],
    ['branding', '브랜딩·패키지', '아이덴티티 리뉴얼, 패키지 디자인 사례.'],
    ['type', '타이포그래피', '서체 사용 사례와 타이포 중심 작업.'],
    ['archive', '포스터·뮤지엄 아카이브', '디자인사 원본 포스터와 공공 컬렉션. 고전 레퍼런스.'],
    ['motion', '일러스트·3D·모션', '일러스트, 3D 렌더, 모션 그래픽, 타이틀 시퀀스.'],
    ['photo', '사진·스톡', '사진가 포트폴리오와 스톡 라이브러리. 촬영 레퍼런스.'],
    ['web', '웹·UI 갤러리', '랜딩 페이지와 앱 화면. 디지털 배너·프로모션 페이지 참고.'],
    ['ai', 'AI 이미지', '생성 이미지 갤러리와 프롬프트. 무드 탐색과 시안 실험용.'],
    ['color', '컬러', '팔레트 생성·탐색 도구.'],
    ['tool', '수집·정리 도구', '모은 레퍼런스를 저장·태깅·비교하는 도구.'],
  ];

  const G = 'site';
  const S = [
    // 포트폴리오
    ['Behance', 'https://www.behance.net', 'portfolio', '글로벌', '무료', 'https://www.behance.net/search/projects?search={q}', 3, 'Adobe 운영. 광고·브랜딩·KV 케이스스터디가 가장 많은 곳. "key visual", "campaign" 검색 추천.'],
    ['Dribbble', 'https://dribbble.com', 'portfolio', '글로벌', '부분유료', 'https://dribbble.com/search/{q}', 2, 'UI·일러스트·브랜딩 샷 중심. 배너·그래픽 스타일 탐색에 빠르다.'],
    ['ArtStation', 'https://www.artstation.com', 'portfolio', '글로벌', '무료', 'https://www.artstation.com/search?sort_by=relevance&query={q}', 2, '게임·영화 컨셉아트, 키아트, 3D. 게임 KV 레퍼런스의 중심.'],
    ['Cargo', 'https://cargo.site', 'portfolio', '글로벌', '부분유료', G, 2, '실험적인 그래픽 디자이너·스튜디오 포트폴리오 호스팅.'],
    ['Adobe Portfolio', 'https://portfolio.adobe.com', 'portfolio', '글로벌', '부분유료', null, 1, 'Behance와 연동된 개인 포트폴리오 사이트들.'],
    ['Coroflot', 'https://www.coroflot.com', 'portfolio', '미국', '무료', G, 1, '산업·제품·그래픽 디자이너 포트폴리오와 채용.'],
    ['Carbonmade', 'https://carbonmade.com', 'portfolio', '글로벌', '유료', null, 1, '크리에이티브 포트폴리오 빌더.'],
    ['Format', 'https://www.format.com', 'portfolio', '글로벌', '유료', null, 1, '사진가·디자이너 포트폴리오 빌더.'],
    ['The Dots', 'https://the-dots.com', 'portfolio', '영국', '무료', G, 1, '영국 크리에이티브 업계 네트워크와 프로젝트.'],
    ['Contra', 'https://contra.com', 'portfolio', '글로벌', '무료', null, 1, '프리랜서 디자이너 포트폴리오.'],
    ['Domestika', 'https://www.domestika.org/projects', 'portfolio', '스페인·중남미', '부분유료', G, 1, '스페인어권 크리에이터 프로젝트 갤러리.'],
    ['Readymag', 'https://readymag.com/explore', 'portfolio', '글로벌', '부분유료', null, 1, '에디토리얼 웹 퍼블리싱 작품 모음.'],
    ['Semplice Showcase', 'https://www.semplice.com/showcase', 'portfolio', '글로벌', '무료', null, 1, '디자이너 포트폴리오 사이트 큐레이션.'],
    ['노트폴리오', 'https://notefolio.net', 'portfolio', '한국', '무료', G, 2, '국내 최대 크리에이티브 포트폴리오 플랫폼. 국내 광고·편집·일러스트.'],
    ['ZCOOL 站酷', 'https://www.zcool.com.cn', 'portfolio', '중국', '무료', G, 3, '중국 최대 디자이너 커뮤니티. 이커머스 배너·주시각(主视觉) 사례가 매우 많다.'],
    ['UI中国', 'https://www.ui.cn', 'portfolio', '중국', '무료', G, 2, '중국 UI·운영 디자인(배너) 포트폴리오.'],

    // 검색엔진
    ['Google 이미지', 'https://images.google.com', 'engine', '글로벌', '무료', 'https://www.google.com/search?tbm=isch&q={q}', 3, '가장 넓은 인덱스. "도구 > 크기·색상·유형" 필터와 site: 연산자를 같이 쓴다.'],
    ['Yandex 이미지', 'https://yandex.com/images', 'engine', '러시아', '무료', 'https://yandex.com/images/search?text={q}', 3, '유사 이미지 묶음이 강하다. 동유럽·러시아 광고, 고해상도 원본 찾기에 좋다.'],
    ['Bing 이미지', 'https://www.bing.com/images', 'engine', '글로벌', '무료', 'https://www.bing.com/images/search?q={q}', 2, '레이아웃·색상 필터가 직관적. 구글과 다른 결과가 섞인다.'],
    ['네이버 이미지', 'https://search.naver.com', 'engine', '한국', '무료', 'https://search.naver.com/search.naver?where=image&query={q}', 2, '국내 광고·블로그 이미지. 한국어 키워드 필수.'],
    ['다음 이미지', 'https://search.daum.net', 'engine', '한국', '무료', 'https://search.daum.net/search?w=img&q={q}', 1, '국내 카페·티스토리 이미지가 섞인다.'],
    ['Baidu 이미지', 'https://image.baidu.com', 'engine', '중국', '무료', 'https://image.baidu.com/search/index?tn=baiduimage&word={q}', 2, '중국 광고·이커머스 KV. "主视觉", "海报" 같은 중국어 키워드로.'],
    ['Yahoo! JAPAN 이미지', 'https://search.yahoo.co.jp/image', 'engine', '일본', '무료', 'https://search.yahoo.co.jp/image/search?p={q}', 2, '일본 포스터·광고. "キービジュアル", "ポスター" 키워드 추천.'],
    ['DuckDuckGo 이미지', 'https://duckduckgo.com', 'engine', '글로벌', '무료', 'https://duckduckgo.com/?q={q}&iax=images&ia=images', 1, '개인화 없는 결과. 같은 검색을 다른 시각으로.'],
    ['Wikimedia Commons', 'https://commons.wikimedia.org', 'engine', '글로벌', '무료', 'https://commons.wikimedia.org/w/index.php?search={q}&title=Special:MediaSearch&type=image', 1, '퍼블릭 도메인·CC 이미지. 저작권 만료 고전 포스터.'],

    // 역이미지
    ['Google Lens', 'https://lens.google.com', 'reverse', '글로벌', '무료', null, 3, '이미지 일부를 잘라 유사 이미지·출처 검색.'],
    ['TinEye', 'https://tineye.com', 'reverse', '글로벌', '부분유료', null, 2, '최초 게시 시점과 원본 해상도 추적에 강하다.'],
    ['Yandex 역검색', 'https://yandex.com/images', 'reverse', '러시아', '무료', null, 3, '구도·색이 비슷한 이미지를 잘 찾는다. 레퍼런스 확장용으로 최고.'],
    ['Bing Visual Search', 'https://www.bing.com/visualsearch', 'reverse', '글로벌', '무료', null, 1, '이미지 내 사물 단위 검색.'],
    ['Baidu 识图', 'https://graph.baidu.com', 'reverse', '중국', '무료', null, 1, '중국 웹에서의 출처 찾기.'],

    // 큐레이션
    ['Pinterest', 'https://www.pinterest.com', 'curation', '글로벌', '무료', 'https://www.pinterest.com/search/pins/?q={q}', 3, '무드보드의 기본. 핀 하나를 열면 "더 보기"로 유사 핀이 끝없이 이어진다.'],
    ['Are.na', 'https://www.are.na', 'curation', '글로벌', '부분유료', G, 2, '디자이너·연구자가 주제별로 쌓은 채널. 알고리즘 없는 깊은 큐레이션.'],
    ['Savee', 'https://savee.it', 'curation', '글로벌', '부분유료', G, 3, '아트디렉터들이 쓰는 고감도 저장소. 광고·패션·그래픽 비율이 높다.'],
    ['Cosmos', 'https://www.cosmos.so', 'curation', '글로벌', '무료', G, 3, '광고 없는 미니멀 무드보드. 에디토리얼·그래픽 감도 높음.'],
    ['Designspiration', 'https://www.designspiration.com', 'curation', '글로벌', '무료', G, 2, '색상 다중 선택 검색이 가능한 디자인 이미지 모음.'],
    ['Muzli', 'https://muz.li', 'curation', '글로벌', '무료', G, 1, '디자인 영감 큐레이션 피드.'],
    ['We Heart It', 'https://weheartit.com', 'curation', '글로벌', '무료', G, 1, '감성 이미지 모음. 라이프스타일 무드.'],
    ['The Inspiration Grid', 'https://theinspirationgrid.com', 'curation', '글로벌', '무료', G, 2, '그래픽·일러스트·타이포 프로젝트 큐레이션.'],
    ['Abduzeedo', 'https://abduzeedo.com', 'curation', '글로벌', '무료', G, 2, '브랜딩·그래픽·3D 프로젝트 소개.'],
    ['Booooooom', 'https://www.booooooom.com', 'curation', '캐나다', '무료', G, 1, '아트·사진·일러스트 큐레이션.'],
    ['Colossal', 'https://www.thisiscolossal.com', 'curation', '미국', '무료', G, 1, '아트·공예·설치. 비주얼 아이디어 원천.'],
    ['花瓣 Huaban', 'https://huaban.com', 'curation', '중국', '무료', G, 3, '중국판 핀터레스트. 중국 광고 주시각·이커머스 배너 수집이 압도적.'],
    ['Instagram', 'https://www.instagram.com', 'curation', '글로벌', '무료', null, 2, '스튜디오·아트디렉터 계정 팔로우가 최신 KV를 가장 빨리 본다.'],
    ['Tumblr', 'https://www.tumblr.com', 'curation', '글로벌', '무료', G, 1, '오래된 디자인 블로그 아카이브가 남아 있다.'],

    // 광고 아카이브
    ['Ads of the World', 'https://www.adsoftheworld.com', 'ad-archive', '글로벌', '부분유료', G, 3, '전 세계 인쇄·옥외·디지털 광고 아카이브. 매체·업종·국가별로 크레딧 포함.'],
    ["Lürzer's Archive", 'https://www.luerzersarchive.com', 'ad-archive', '글로벌', '유료', G, 3, '1984년부터 이어진 광고 아카이브 매거진. 선별 품질이 높다.'],
    ['Love the Work (Cannes Lions)', 'https://www.lovethework.com', 'ad-archive', '글로벌', '유료', null, 3, '칸 라이언즈 수상작 공식 아카이브.'],
    ['Cannes Lions', 'https://www.canneslions.com', 'ad-archive', '글로벌', '부분유료', null, 2, '세계 최대 크리에이티비티 페스티벌.'],
    ['D&AD', 'https://www.dandad.org', 'ad-archive', '영국', '부분유료', G, 3, '옐로·블랙 펜슬 수상작 아카이브. 크래프트 기준점.'],
    ['The One Show', 'https://www.oneshow.org', 'ad-archive', '미국', '부분유료', G, 2, 'One Club 주관 광고상.'],
    ['ADC Awards', 'https://www.adcawards.org', 'ad-archive', '미국', '부분유료', G, 3, '아트디렉터스클럽. 아트디렉션·일러스트·타이포 부문이 KV에 직접적.'],
    ['Clio Awards', 'https://clios.com', 'ad-archive', '미국', '부분유료', G, 2, '클리오 광고상 수상작.'],
    ['New York Festivals', 'https://www.newyorkfestivals.com', 'ad-archive', '미국', '부분유료', null, 1, '뉴욕 페스티벌 광고상.'],
    ['Spikes Asia', 'https://www.spikes.asia', 'ad-archive', '아시아', '부분유료', null, 2, '아시아·태평양 광고제.'],
    ['ADFEST', 'https://www.adfest.com', 'ad-archive', '아시아', '부분유료', null, 2, '아시아 태평양 광고제(태국 파타야).'],
    ['Effie Awards', 'https://www.effie.org', 'ad-archive', '글로벌', '부분유료', null, 1, '효과(성과) 중심 광고상. 전략 레퍼런스.'],
    ['Communication Arts', 'https://www.commarts.com', 'ad-archive', '미국', '부분유료', G, 3, '광고·디자인·일러스트 연감. 1959년 창간.'],
    ['Graphis', 'https://www.graphis.com', 'ad-archive', '글로벌', '유료', G, 2, '포스터·광고·디자인 연감.'],
    ['AdForum', 'https://www.adforum.com', 'ad-archive', '글로벌', '부분유료', G, 2, '에이전시 정보와 크리에이티브 라이브러리.'],
    ['Campaign Brief', 'https://campaignbrief.com', 'ad-archive', '호주·아시아', '무료', G, 2, '호주·뉴질랜드·아시아 광고 신작 뉴스.'],
    ['Campaign', 'https://www.campaignlive.com', 'ad-archive', '영국', '유료', G, 1, '광고 업계 뉴스와 캠페인.'],
    ['The Drum', 'https://www.thedrum.com', 'ad-archive', '영국', '부분유료', G, 1, '마케팅·광고 뉴스와 캠페인 소개.'],
    ['Ad Age', 'https://adage.com', 'ad-archive', '미국', '유료', null, 1, '미국 광고 업계지.'],
    ['Adweek', 'https://www.adweek.com', 'ad-archive', '미국', '부분유료', null, 1, '미국 광고 업계지.'],
    ['Muse by Clio', 'https://musebyclios.com', 'ad-archive', '미국', '무료', G, 2, '클리오가 운영하는 크리에이티브 뉴스.'],
    ['LBBOnline', 'https://www.lbbonline.com', 'ad-archive', '글로벌', '무료', G, 2, 'Little Black Book. 프로덕션·크리에이티브 크레딧.'],
    ['shots', 'https://shots.net', 'ad-archive', '영국', '유료', null, 1, '광고 영상·프로덕션 아카이브.'],
    ['iSpot.tv', 'https://www.ispot.tv', 'ad-archive', '미국', '부분유료', null, 1, '미국 TV 광고 데이터베이스.'],
    ['TVCF', 'https://www.tvcf.co.kr', 'ad-archive', '한국', '부분유료', G, 3, '국내 TV·인쇄·디지털 광고 최대 아카이브.'],
    ['광고정보센터', 'https://www.adic.co.kr', 'ad-archive', '한국', '무료', G, 2, '국내 광고 자료·수상작·연구 자료.'],
    ['AdverTimes 宣伝会議', 'https://www.advertimes.com', 'ad-archive', '일본', '무료', G, 2, '일본 광고 업계 미디어. 신작 포스터·캠페인.'],
    ['ACC TOKYO CREATIVITY AWARDS', 'https://www.acc-awards.com', 'ad-archive', '일본', '무료', null, 2, '일본 대표 광고상 수상작.'],
    ['数英 Digitaling', 'https://www.digitaling.com', 'ad-archive', '중국', '무료', G, 2, '중국 광고 캠페인 케이스 아카이브.'],
    ['TOPYS', 'https://www.topys.cn', 'ad-archive', '중국', '무료', G, 1, '중국 크리에이티브 큐레이션.'],

    // 집행 광고 라이브러리
    ['Meta 광고 라이브러리', 'https://www.facebook.com/ads/library', 'ad-library', '글로벌', '무료', 'https://www.facebook.com/ads/library/?active_status=all&ad_type=all&country=ALL&q={q}', 3, '페이스북·인스타그램에서 지금 집행 중인 광고 소재 전체 조회. 경쟁사 배너 분석 필수.'],
    ['Google 광고 투명성 센터', 'https://adstransparency.google.com', 'ad-library', '글로벌', '무료', null, 2, '광고주별 구글 검색·디스플레이·유튜브 광고 소재.'],
    ['TikTok Creative Center', 'https://ads.tiktok.com/business/creativecenter', 'ad-library', '글로벌', '무료', null, 2, '틱톡 인기 광고 소재와 트렌드.'],
    ['LinkedIn 광고 라이브러리', 'https://www.linkedin.com/ad-library', 'ad-library', '글로벌', '무료', null, 1, 'B2B 광고 소재 조회.'],
    ['Really Good Emails', 'https://reallygoodemails.com', 'ad-library', '글로벌', '무료', G, 1, '이메일 프로모션 디자인 아카이브.'],
    ['Milled', 'https://milled.com', 'ad-library', '글로벌', '무료', G, 1, '브랜드 프로모션 이메일 검색.'],
    ['Foreplay', 'https://www.foreplay.co', 'ad-library', '글로벌', '유료', null, 1, '퍼포먼스 광고 스와이프 파일 수집 도구.'],

    // 매거진
    ["It's Nice That", 'https://www.itsnicethat.com', 'magazine', '영국', '무료', G, 2, '그래픽·일러스트·사진 신작. 젊은 스튜디오 발굴.'],
    ['Creative Review', 'https://www.creativereview.co.uk', 'magazine', '영국', '부분유료', G, 2, '광고·디자인 비평. 연간 어워드.'],
    ['AIGA Eye on Design', 'https://eyeondesign.aiga.org', 'magazine', '미국', '무료', G, 2, 'AIGA의 그래픽 디자인 매거진.'],
    ['Print Magazine', 'https://www.printmag.com', 'magazine', '미국', '무료', G, 1, '그래픽 디자인 매거진.'],
    ['designboom', 'https://www.designboom.com', 'magazine', '이탈리아', '무료', G, 1, '건축·디자인·아트.'],
    ['Dezeen', 'https://www.dezeen.com', 'magazine', '영국', '무료', G, 1, '건축·인테리어·그래픽 디자인 뉴스.'],
    ['Wallpaper*', 'https://www.wallpaper.com', 'magazine', '영국', '부분유료', G, 1, '디자인·패션·라이프스타일.'],
    ['Fubiz', 'https://www.fubiz.net', 'magazine', '프랑스', '무료', G, 1, '크리에이티브 컬처 큐레이션.'],
    ['Design Week', 'https://www.designweek.co.uk', 'magazine', '영국', '부분유료', G, 1, '영국 디자인 업계 뉴스.'],
    ['Motionographer', 'https://motionographer.com', 'magazine', '글로벌', '무료', G, 1, '모션 디자인 저널.'],
    ['디자인DB', 'https://www.designdb.com', 'magazine', '한국', '무료', G, 1, '한국디자인진흥원 디자인 정보 포털.'],

    // 브랜딩·패키지
    ['Brand New', 'https://www.underconsideration.com/brandnew', 'branding', '미국', '부분유료', G, 2, '아이덴티티 리뉴얼 리뷰. 비포·애프터 비교.'],
    ['BP&O', 'https://bpando.org', 'branding', '영국', '무료', G, 2, '브랜딩·패키징·의견.'],
    ['Identity Designed', 'https://identitydesigned.com', 'branding', '글로벌', '무료', G, 2, '아이덴티티 시스템 케이스스터디.'],
    ['The Brand Identity', 'https://the-brandidentity.com', 'branding', '영국', '무료', G, 2, '브랜드 아이덴티티 큐레이션.'],
    ['Mindsparkle Mag', 'https://mindsparklemag.com', 'branding', '스위스', '무료', G, 2, '브랜딩·웹·그래픽 프로젝트.'],
    ['The Dieline', 'https://thedieline.com', 'branding', '미국', '무료', G, 2, '패키지 디자인 대표 매체.'],
    ['Packaging of the World', 'https://packagingoftheworld.com', 'branding', '글로벌', '무료', G, 2, '패키지 디자인 아카이브.'],
    ['World Brand Design Society', 'https://worldbranddesign.com', 'branding', '글로벌', '무료', G, 1, '브랜드·패키지 디자인 아카이브.'],
    ['Pentawards', 'https://pentawards.com', 'branding', '글로벌', '부분유료', null, 2, '패키지 디자인 전문 어워드.'],
    ['Lovely Package', 'https://lovelypackage.com', 'branding', '글로벌', '무료', G, 1, '패키지 디자인 큐레이션.'],
    ['Red Dot', 'https://www.red-dot.org', 'branding', '독일', '부분유료', null, 2, '브랜드·커뮤니케이션 디자인 부문 수상작.'],
    ['iF Design', 'https://ifdesign.com', 'branding', '독일', '무료', null, 2, '방대한 수상작 데이터베이스(커뮤니케이션 부문 포함).'],
    ['Good Design Award', 'https://www.g-mark.org', 'branding', '일본', '무료', null, 1, '일본 굿디자인상 데이터베이스.'],

    // 타이포
    ['Fonts In Use', 'https://fontsinuse.com', 'type', '글로벌', '무료', G, 3, '서체별·업종별 실제 사용 사례. 포스터·광고 타이포 레퍼런스의 정석.'],
    ['Typewolf', 'https://www.typewolf.com', 'type', '미국', '무료', G, 1, '웹 타이포 사례와 서체 페어링.'],
    ['Type Directors Club', 'https://www.tdc.org', 'type', '미국', '부분유료', G, 2, 'TDC 타이포그래피 어워드.'],
    ['Typographica', 'https://typographica.org', 'type', '글로벌', '무료', null, 1, '서체 리뷰.'],
    ['I Love Typography', 'https://ilovetypography.com', 'type', '글로벌', '무료', null, 1, '타이포 역사와 서체.'],
    ['Typo/graphic posters', 'https://www.typographicposters.com', 'type', '글로벌', '무료', G, 3, '타이포그래피 포스터 대규모 아카이브.'],

    // 아카이브
    ['Poster House', 'https://posterhouse.org', 'archive', '미국', '무료', G, 2, '미국 최초 포스터 전문 뮤지엄.'],
    ['International Poster Gallery', 'https://www.internationalposter.com', 'archive', '미국', '무료', G, 2, '빈티지 광고 포스터 아카이브(판매 포함).'],
    ['eMuseum Museum für Gestaltung', 'https://www.emuseum.ch', 'archive', '스위스', '무료', G, 3, '취리히 디자인 뮤지엄 포스터 컬렉션. 스위스 스타일 원본.'],
    ['Polish Poster Gallery', 'https://www.polishposter.com', 'archive', '폴란드', '무료', G, 2, '폴란드 포스터 학파 작품.'],
    ['Cooper Hewitt Collection', 'https://collection.cooperhewitt.org', 'archive', '미국', '무료', G, 2, '스미스소니언 디자인 뮤지엄 컬렉션.'],
    ['MoMA Collection', 'https://www.moma.org/collection', 'archive', '미국', '무료', G, 2, '그래픽 디자인 컬렉션 포함.'],
    ['V&A Collections', 'https://collections.vam.ac.uk', 'archive', '영국', '무료', G, 1, '영국 V&A 뮤지엄 포스터·광고물.'],
    ['Letterform Archive', 'https://letterformarchive.org', 'archive', '미국', '무료', G, 2, '레터링·타이포·그래픽 디자인 아카이브.'],
    ['Library of Congress', 'https://www.loc.gov/pictures', 'archive', '미국', '무료', 'https://www.loc.gov/pictures/search/?q={q}', 1, '미국 의회도서관 포스터·인쇄물.'],
    ['Internet Archive', 'https://archive.org', 'archive', '글로벌', '무료', 'https://archive.org/search?query={q}', 1, '잡지 스캔본, 옛 광고 지면.'],
    ['Europeana', 'https://www.europeana.eu', 'archive', '유럽', '무료', 'https://www.europeana.eu/en/search?query={q}', 1, '유럽 문화유산 디지털 컬렉션.'],
    ['Rijksmuseum', 'https://www.rijksmuseum.nl', 'archive', '네덜란드', '무료', G, 1, '고해상도 퍼블릭 도메인 이미지.'],
    ['The Met Open Access', 'https://www.metmuseum.org/art/collection', 'archive', '미국', '무료', G, 1, '퍼블릭 도메인 미술품. 색·구도 레퍼런스.'],
    ['ginza graphic gallery', 'https://www.dnp.co.jp/gallery/ggg/', 'archive', '일본', '무료', null, 2, '일본 그래픽 디자인 전문 갤러리 전시 기록.'],
    ['JAGDA', 'https://www.jagda.or.jp', 'archive', '일본', '무료', null, 1, '일본그래픽디자인협회. 연감과 수상작.'],

    // 일러스트·3D·모션
    ['Society of Illustrators', 'https://www.societyillustrators.org', 'motion', '미국', '부분유료', null, 1, '일러스트레이션 연례 전시·수상작.'],
    ['American Illustration', 'https://www.ai-ap.com', 'motion', '미국', '무료', G, 1, '일러스트·사진 연감.'],
    ['3x3 Magazine', 'https://3x3mag.com', 'motion', '미국', '부분유료', null, 1, '현대 일러스트레이션 매거진.'],
    ['DeviantArt', 'https://www.deviantart.com', 'motion', '글로벌', '무료', G, 1, '대형 아트 커뮤니티.'],
    ['Sketchfab', 'https://sketchfab.com', 'motion', '글로벌', '부분유료', G, 1, '3D 모델 뷰어. 3D KV 오브젝트 레퍼런스.'],
    ['Vimeo', 'https://vimeo.com', 'motion', '글로벌', '부분유료', 'https://vimeo.com/search?q={q}', 2, '광고 영상·모션 그래픽 원본. Staff Picks 추천.'],
    ['YouTube', 'https://www.youtube.com', 'motion', '글로벌', '무료', 'https://www.youtube.com/results?search_query={q}', 1, '광고 영상, 메이킹 필름.'],
    ['Art of the Title', 'https://www.artofthetitle.com', 'motion', '글로벌', '무료', G, 2, '영화·드라마 타이틀 시퀀스 아카이브.'],
    ['STASH', 'https://www.stashmedia.tv', 'motion', '글로벌', '유료', null, 1, '모션·VFX 큐레이션.'],
    ['LottieFiles', 'https://lottiefiles.com', 'motion', '글로벌', '부분유료', null, 1, '웹 애니메이션 배너 소재.'],

    // 사진·스톡
    ['Unsplash', 'https://unsplash.com', 'photo', '글로벌', '무료', 'https://unsplash.com/s/photos/{q}', 2, '감도 높은 무료 사진. 시안용 배경.'],
    ['Pexels', 'https://www.pexels.com', 'photo', '글로벌', '무료', 'https://www.pexels.com/search/{q}/', 1, '무료 사진·영상.'],
    ['Pixabay', 'https://pixabay.com', 'photo', '글로벌', '무료', 'https://pixabay.com/images/search/{q}/', 1, '무료 사진·일러스트.'],
    ['Getty Images', 'https://www.gettyimages.com', 'photo', '글로벌', '유료', 'https://www.gettyimages.com/search/2/image?phrase={q}', 2, '프리미엄 스톡과 보도사진.'],
    ['Shutterstock', 'https://www.shutterstock.com', 'photo', '글로벌', '유료', 'https://www.shutterstock.com/search/{q}', 1, '대형 스톡 라이브러리.'],
    ['Adobe Stock', 'https://stock.adobe.com', 'photo', '글로벌', '유료', 'https://stock.adobe.com/search?k={q}', 1, '사진·벡터·3D 스톡.'],
    ['iStock', 'https://www.istockphoto.com', 'photo', '글로벌', '유료', null, 1, 'Getty 계열 스톡.'],
    ['Stocksy', 'https://www.stocksy.com', 'photo', '캐나다', '유료', G, 2, '작가 협동조합 스톡. 광고 촬영 무드 레퍼런스로 좋다.'],
    ['Freepik', 'https://www.freepik.com', 'photo', '스페인', '부분유료', 'https://www.freepik.com/search?query={q}', 1, '벡터·PSD 목업·사진.'],
    ['Flickr', 'https://www.flickr.com', 'photo', '글로벌', '무료', 'https://www.flickr.com/search/?text={q}', 1, '빈티지 광고 스캔 그룹이 많다.'],
    ['500px', 'https://500px.com', 'photo', '글로벌', '부분유료', G, 1, '사진가 커뮤니티.'],
    ['VSCO', 'https://www.vsco.co', 'photo', '글로벌', '무료', null, 1, '필름 톤 사진 커뮤니티.'],
    ['Magnum Photos', 'https://www.magnumphotos.com', 'photo', '글로벌', '유료', G, 1, '다큐멘터리 사진 에이전시.'],
    ['LensCulture', 'https://www.lensculture.com', 'photo', '글로벌', '부분유료', G, 1, '현대 사진 큐레이션과 어워드.'],

    // 웹·UI
    ['Awwwards', 'https://www.awwwards.com', 'web', '글로벌', '부분유료', G, 2, '웹사이트 디자인 어워드. 프로모션 사이트 KV.'],
    ['CSS Design Awards', 'https://www.cssdesignawards.com', 'web', '글로벌', '무료', null, 1, '웹 디자인 어워드.'],
    ['The FWA', 'https://thefwa.com', 'web', '글로벌', '무료', null, 1, '인터랙티브 웹 어워드.'],
    ['The Webby Awards', 'https://www.webbyawards.com', 'web', '미국', '부분유료', null, 1, '인터넷 분야 어워드.'],
    ['SiteInspire', 'https://www.siteinspire.com', 'web', '글로벌', '무료', null, 1, '웹 디자인 큐레이션.'],
    ['Land-book', 'https://land-book.com', 'web', '글로벌', '부분유료', null, 2, '랜딩 페이지 갤러리.'],
    ['Lapa Ninja', 'https://www.lapa.ninja', 'web', '글로벌', '무료', null, 1, '랜딩 페이지 모음.'],
    ['One Page Love', 'https://onepagelove.com', 'web', '글로벌', '무료', null, 1, '원페이지 사이트 모음.'],
    ['Godly', 'https://godly.website', 'web', '글로벌', '무료', null, 2, '감도 높은 웹 디자인 큐레이션.'],
    ['Minimal Gallery', 'https://minimal.gallery', 'web', '글로벌', '무료', null, 1, '미니멀 웹 디자인.'],
    ['Httpster', 'https://httpster.net', 'web', '글로벌', '무료', null, 1, '웹 디자인 쇼케이스.'],
    ['Mobbin', 'https://mobbin.com', 'web', '글로벌', '부분유료', null, 1, '앱 화면 라이브러리. 인앱 배너 참고.'],
    ['Refero', 'https://refero.design', 'web', '글로벌', '부분유료', null, 1, '웹·앱 UI 레퍼런스 검색.'],
    ['Page Flows', 'https://pageflows.com', 'web', '글로벌', '유료', null, 1, '사용자 흐름 녹화 라이브러리.'],

    // AI
    ['Midjourney Explore', 'https://www.midjourney.com/explore', 'ai', '글로벌', '유료', null, 2, '생성 이미지와 프롬프트. 무드 탐색용.'],
    ['Lexica', 'https://lexica.art', 'ai', '글로벌', '부분유료', 'https://lexica.art/?q={q}', 1, '생성 이미지 프롬프트 검색.'],
    ['Krea', 'https://www.krea.ai', 'ai', '글로벌', '부분유료', null, 1, '생성 이미지 도구와 갤러리.'],
    ['Civitai', 'https://civitai.com', 'ai', '글로벌', '무료', null, 1, '오픈 모델·생성 이미지 커뮤니티.'],
    ['PromptHero', 'https://prompthero.com', 'ai', '글로벌', '부분유료', null, 1, '프롬프트 검색.'],
    ['Adobe Firefly', 'https://firefly.adobe.com', 'ai', '글로벌', '부분유료', null, 1, '상업 사용을 고려한 생성 이미지.'],

    // 컬러
    ['Coolors', 'https://coolors.co', 'color', '글로벌', '부분유료', null, 1, '팔레트 생성·탐색.'],
    ['Adobe Color', 'https://color.adobe.com', 'color', '글로벌', '무료', null, 2, '이미지에서 팔레트 추출, 트렌드 팔레트.'],
    ['Color Hunt', 'https://colorhunt.co', 'color', '글로벌', '무료', null, 1, '4색 팔레트 모음.'],
    ['Khroma', 'https://www.khroma.co', 'color', '글로벌', '무료', null, 1, '취향 학습 팔레트 생성.'],
    ['Huemint', 'https://huemint.com', 'color', '글로벌', '무료', null, 1, '실제 레이아웃에 입혀 보는 팔레트 생성.'],

    // 도구
    ['PureRef', 'https://www.pureref.com', 'tool', '글로벌', '무료', null, 2, '레퍼런스 이미지를 한 캔버스에 띄워 두는 데스크톱 도구.'],
    ['Eagle', 'https://eagle.cool', 'tool', '글로벌', '유료', null, 2, '로컬 이미지 레퍼런스 관리·태깅.'],
    ['Milanote', 'https://milanote.com', 'tool', '글로벌', '부분유료', null, 1, '무드보드·기획 보드.'],
    ['mymind', 'https://mymind.com', 'tool', '글로벌', '유료', null, 1, '자동 태깅 개인 저장소.'],
    ['Raindrop.io', 'https://raindrop.io', 'tool', '글로벌', '부분유료', null, 1, '북마크 수집·태깅.'],
  ];

  const SOURCES = S.map(function (r, i) {
    return { id: 's' + i, name: r[0], url: r[1], cat: r[2], region: r[3], cost: r[4], search: r[5], kv: r[6], desc: r[7] };
  });

  const TECHNIQUES = [
    ['metaphor', '시각적 은유', '제품의 효익을 다른 사물로 바꿔 보여준다. 한 장으로 설명이 끝난다.'],
    ['negative', '여백·네거티브 스페이스', '빈 공간이 주인공. 작은 요소 하나로 시선을 모은다.'],
    ['type', '타이포 주도', '글자 자체가 이미지. 크기·굵기·배치로 위계를 만든다.'],
    ['copy', '헤드라인 주도', '카피 한 줄이 아이디어의 전부. 비주얼은 받쳐 준다.'],
    ['hero', '제품 히어로', '제품을 정면으로, 크게, 완벽하게.'],
    ['crop', '크롭·부분 노출', '일부만 보여주고 나머지는 보는 사람이 완성하게 한다.'],
    ['silhouette', '실루엣', '형태만 남기고 디테일을 지운다.'],
    ['photo', '인물·다큐 사진', '사람의 얼굴과 진짜 순간이 설득한다.'],
    ['illust', '일러스트', '사진으로 못 하는 세계를 그린다.'],
    ['collage', '콜라주·포토몽타주', '서로 다른 이미지를 잘라 붙여 새 의미를 만든다.'],
    ['geometric', '기하·구성', '원·사각·사선의 구성으로 리듬을 만든다.'],
    ['grid', '스위스 그리드', '엄격한 그리드와 산세리프. 정보의 질서.'],
    ['color', '컬러 블로킹', '면 색 하나가 브랜드를 기억하게 만든다.'],
    ['logo', '로고·심볼 변형', '로고를 비틀거나 쪼개 메시지를 만든다.'],
    ['sequence', '시리즈 변주', '같은 틀 안에서 한 요소만 바꿔 연작을 만든다.'],
    ['data', '데이터·개인화', '사용자 데이터나 이름을 소재로 쓴다.'],
    ['ugc', '참여·UGC', '사람들이 만든 결과물이 곧 광고가 된다.'],
    ['shock', '역발상·도발', '금기나 약점을 정면으로 드러내 주목을 얻는다.'],
    ['humor', '유머', '웃음으로 경계를 푼다.'],
    ['handmade', '손글씨·수작업', '사람의 손맛, 불완전함의 친근함.'],
  ];

  const MOODS = ['미니멀', '대담', '유머러스', '감성', '도발', '고급', '레트로', '실험적', '진지함', '따뜻함'];
  const INDUSTRIES = ['자동차·모빌리티', 'IT·전자', '통신', '게임', '식음료', '주류', '패션·뷰티', '화장품', '스포츠', '금융', '보험', '제약·헬스케어', '유통·이커머스', '리테일·생활', '가전', '부동산·건설', '여행·항공', '교육', '미디어·엔터', '공공·캠페인', '문화·예술', '브랜드 아이덴티티', '기타'];

  // 광고 분야(목적)
  const FIELDS = ['브랜드', '제품·서비스', '프로모션·세일', '공익·사회', '문화·공연', '기업 PR', '리크루트', '아이덴티티'];

  // [ISO 코드, 이름, 권역]  XX = 글로벌·다국가
  const COUNTRIES = [
    ['KR', '한국', '아시아'], ['JP', '일본', '아시아'], ['CN', '중국', '아시아'], ['TW', '대만', '아시아'], ['HK', '홍콩', '아시아'], ['SG', '싱가포르', '아시아'], ['TH', '태국', '아시아'], ['VN', '베트남', '아시아'], ['ID', '인도네시아', '아시아'], ['PH', '필리핀', '아시아'], ['IN', '인도', '아시아'], ['AE', 'UAE', '중동·아프리카'], ['ZA', '남아공', '중동·아프리카'],
    ['US', '미국', '북미'], ['CA', '캐나다', '북미'], ['MX', '멕시코', '중남미'], ['BR', '브라질', '중남미'], ['AR', '아르헨티나', '중남미'], ['CO', '콜롬비아', '중남미'],
    ['GB', '영국', '유럽'], ['FR', '프랑스', '유럽'], ['DE', '독일', '유럽'], ['NL', '네덜란드', '유럽'], ['SE', '스웨덴', '유럽'], ['NO', '노르웨이', '유럽'], ['DK', '덴마크', '유럽'], ['CH', '스위스', '유럽'], ['IT', '이탈리아', '유럽'], ['ES', '스페인', '유럽'], ['PT', '포르투갈', '유럽'], ['PL', '폴란드', '유럽'], ['BE', '벨기에', '유럽'], ['AT', '오스트리아', '유럽'], ['RU', '러시아', '유럽'],
    ['AU', '호주', '오세아니아'], ['NZ', '뉴질랜드', '오세아니아'],
    ['XX', '글로벌·다국가', '글로벌'],
  ];

  // 클래식 KV의 집행·제작 국가 (다국가 집행은 XX)
  const KV_COUNTRY = {
    'vw-think-small': 'US', 'vw-lemon': 'US', 'avis-try-harder': 'US', 'apple-1984': 'US', 'apple-think-different': 'US', 'apple-ipod-silhouette': 'US', 'apple-shot-on-iphone': 'XX',
    'absolut-bottle': 'US', 'economist-red': 'GB', 'nike-jdi': 'US', 'nike-dream-crazy': 'US', 'nike-londoner': 'GB', 'adidas-impossible': 'XX', 'heinz-draw-ketchup': 'CA', 'heinz-pass': 'US',
    'mcd-follow-arches': 'CA', 'bk-moldy': 'XX', 'bk-burning-stores': 'BR', 'kitkat-break': 'GB', 'coke-share': 'AU', 'dove-real-beauty': 'XX', 'dove-sketches': 'XX', 'spotify-weird': 'XX',
    'spotify-wrapped': 'XX', 'old-spice': 'US', 'snickers-hungry': 'US', 'always-likeagirl': 'US', 'volvo-epic-split': 'SE', 'marmite': 'GB', 'got-milk': 'US', 'milk-mustache': 'US',
    'benetton-colors': 'IT', 'wonderbra': 'GB', 'lacoste-species': 'FR', 'mastercard-priceless': 'US', 'ikea-lamp': 'US', 'ikea-bookbook': 'SG', 'vw-force': 'US', 'dumb-ways': 'AU', 'kfc-fck': 'GB',
    'patagonia-dont-buy': 'US', 'oatly-copy': 'SE', 'sony-balls': 'GB', 'guinness-surfer': 'GB', 'uniqlock': 'JP', 'muji-horizon': 'JP', 'bass-anatomy': 'US', 'bass-vertigo': 'US', 'glaser-iny': 'US',
    'glaser-dylan': 'US', 'obama-hope': 'US', 'mb-der-film': 'CH', 'mb-beethoven': 'CH', 'cassandre-nord': 'FR', 'cassandre-dubonnet': 'FR', 'rodchenko-books': 'RU', 'tanaka-nihon-buyo': 'JP',
    'scher-public-theater': 'US', 'sagmeister-aiga': 'US', 'polish-school': 'PL', 'bauhaus-exhibition': 'DE', 'baemin-type': 'KR', 'bk-whopper-detour': 'US', 'guardian-pigs': 'GB', 'airbnb-belo': 'US',
    'mailchimp-rebrand': 'US', 'tide-ad': 'US',
  };

  // 데이터 출처
  const ORIGINS = [['classic', '클래식 레퍼런스'], ['tvcf', 'TVCF 우수 인쇄광고'], ['trend', '트렌드 수집'], ['user', '직접 분석·등록']];

  // 역프롬프트 대상 생성 도구
  const PROMPT_TARGETS = [['midjourney', 'Midjourney'], ['flux', 'FLUX / SDXL'], ['gptimage', 'GPT Image · DALL·E'], ['imagen', 'Imagen · Gemini'], ['firefly', 'Adobe Firefly'], ['video', '영상(Sora·Veo·Kling)']];
  const MEDIA = [['print', '인쇄'], ['ooh', '옥외'], ['poster', '포스터'], ['digital', '디지털·소셜'], ['film', '영상'], ['packaging', '패키지'], ['identity', '아이덴티티']];

  // [id, 제목, 브랜드, 연도, 제작, 업종, 매체[], 기법[], 무드[], 색[], 포인트, 검색어]
  const K = [
    ['vw-think-small', 'Think Small', 'Volkswagen', 1959, 'DDB', '자동차·모빌리티', ['print'], ['negative', 'copy', 'hero'], ['미니멀', '진지함'], ['#f4f4f2', '#1a1a1a', '#9a9a96'], '지면 대부분을 비우고 작은 비틀 한 대만 둔다. 작다는 약점을 장점으로 뒤집은 역발상 카피와 여백의 교과서.', 'Volkswagen Think Small ad 1959'],
    ['vw-lemon', 'Lemon.', 'Volkswagen', 1960, 'DDB', '자동차·모빌리티', ['print'], ['copy', 'shock', 'hero'], ['미니멀', '유머러스'], ['#f2f2ee', '#151515', '#c9c3a8'], '멀쩡해 보이는 차에 "불량품" 한 단어. 깐깐한 검수를 자랑하기 위해 스스로 흠을 말한다.', 'Volkswagen Lemon ad'],
    ['avis-try-harder', 'We Try Harder', 'Avis', 1962, 'DDB', '자동차·모빌리티', ['print'], ['copy', 'shock'], ['진지함', '대담'], ['#ffffff', '#d4002a', '#111111'], '업계 2등임을 먼저 인정하고, 그래서 더 열심히 한다고 말한다. 언더독 포지셔닝의 원형.', 'Avis We Try Harder ad'],
    ['apple-1984', '1984', 'Apple Macintosh', 1984, 'Chiat\\Day', 'IT·전자', ['film'], ['shock', 'metaphor'], ['도발', '진지함'], ['#5b6d73', '#c9d2d4', '#e85d1f'], '회색 디스토피아 속 붉은 반바지의 질주와 해머. 단 한 번의 방영으로 신화가 된 영상의 색 대비.', 'Apple 1984 commercial'],
    ['apple-think-different', 'Think Different', 'Apple', 1997, 'TBWA\\Chiat\\Day', 'IT·전자', ['print', 'ooh'], ['photo', 'copy', 'sequence'], ['진지함', '고급'], ['#111111', '#ffffff', '#8c8c8c'], '흑백 인물 사진, 작은 로고, 두 단어. 브랜드 철학을 인물로 대신 말하게 한다.', 'Apple Think Different poster'],
    ['apple-ipod-silhouette', 'iPod Silhouette', 'Apple', 2003, 'TBWA\\Chiat\\Day', 'IT·전자', ['ooh', 'print', 'film'], ['silhouette', 'color', 'sequence'], ['대담', '실험적'], ['#ff3ea5', '#b6e51d', '#00a3e0', '#111111', '#ffffff'], '형광 단색 배경, 검은 실루엣, 흰 이어폰 줄 하나. 제품을 가장 작은 디테일로 각인시킨다.', 'iPod silhouette ad'],
    ['apple-shot-on-iphone', 'Shot on iPhone', 'Apple', 2015, 'TBWA\\Media Arts Lab', 'IT·전자', ['ooh', 'digital'], ['ugc', 'photo'], ['감성', '미니멀'], ['#f5f5f7', '#1d1d1f', '#6e6e73'], '사용자가 찍은 사진을 그대로 빌보드에. 결과물이 곧 제품 성능의 증거가 된다.', 'Shot on iPhone billboard'],
    ['absolut-bottle', 'Absolut 병 시리즈', 'Absolut Vodka', 1980, 'TBWA', '주류', ['print'], ['hero', 'sequence', 'metaphor', 'crop'], ['고급', '유머러스'], ['#0b3c8c', '#e8eef5', '#b8a77a'], '모든 광고가 병 실루엣을 변주하고 "Absolut ___" 두 단어로 끝난다. 수십 년 이어진 시리즈 변주의 교본.', 'Absolut Vodka bottle ad campaign'],
    ['economist-red', '레드 포스터 시리즈', 'The Economist', 1988, 'Abbott Mead Vickers BBDO', '미디어·엔터', ['ooh', 'poster'], ['type', 'copy', 'color'], ['미니멀', '유머러스'], ['#e3120b', '#ffffff', '#111111'], '빨간 바탕에 흰 글자 한 줄. 브랜드 컬러와 위트 있는 카피만으로 수십 년간 연작을 쌓았다.', 'The Economist red poster ad'],
    ['nike-jdi', 'Just Do It', 'Nike', 1988, 'Wieden+Kennedy', '스포츠', ['print', 'film'], ['copy', 'photo'], ['대담', '진지함'], ['#111111', '#ffffff', '#f05a28'], '세 단어 태그라인과 운동하는 사람의 사진. 제품보다 태도를 판다.', 'Nike Just Do It ad 1988'],
    ['nike-dream-crazy', 'Dream Crazy', 'Nike', 2018, 'Wieden+Kennedy', '스포츠', ['print', 'ooh', 'film'], ['photo', 'copy', 'shock'], ['도발', '진지함'], ['#141414', '#f2f2f2', '#7a7a7a'], '흑백 클로즈업 얼굴 위에 짧은 문장. 사회적 논쟁을 정면으로 껴안은 브랜드 스탠스.', 'Nike Dream Crazy Kaepernick billboard'],
    ['nike-londoner', 'Nothing Beats a Londoner', 'Nike', 2018, 'Wieden+Kennedy London', '스포츠', ['film', 'ooh'], ['photo', 'sequence', 'humor'], ['대담', '유머러스'], ['#1b1b1b', '#e9e4dc', '#c6262e'], '런던 동네 청소년들의 실제 얼굴과 사투리. 로컬 디테일이 글로벌 브랜드를 친근하게 만든다.', 'Nike Nothing Beats a Londoner'],
    ['adidas-impossible', 'Impossible is Nothing', 'adidas', 2004, '180 Amsterdam', '스포츠', ['print', 'film'], ['copy', 'photo'], ['대담', '감성'], ['#000000', '#ffffff', '#1d4f91'], '선수의 서사와 역설적인 태그라인. 스포츠 KV의 감정 곡선.', 'adidas Impossible is Nothing ad'],
    ['heinz-draw-ketchup', 'Draw Ketchup', 'Heinz', 2021, 'Rethink', '식음료', ['digital', 'ooh'], ['ugc', 'illust', 'logo'], ['유머러스', '따뜻함'], ['#c8102e', '#ffffff', '#2f6f3e'], '사람들에게 케첩을 그려 달라 했더니 대부분 하인즈 병을 그렸다. 그 낙서들이 그대로 광고가 됐다.', 'Heinz Draw Ketchup campaign'],
    ['heinz-pass', 'Pass the Heinz', 'Heinz', 2017, 'David Miami', '식음료', ['ooh', 'print'], ['hero', 'crop'], ['레트로', '유머러스'], ['#c8102e', '#f2e6d0', '#3a2a1e'], '드라마 속 가상의 광고 시안을 실제로 집행. 케첩이 빠진 음식 사진만으로 제품을 떠올리게 한다.', 'Pass the Heinz Mad Men ad'],
    ['mcd-follow-arches', 'Follow the Arches', "McDonald's", 2018, 'Cossette', '식음료', ['ooh'], ['crop', 'logo', 'negative', 'color'], ['미니멀', '유머러스'], ['#da291c', '#ffbc0d', '#ffffff'], '골든 아치 로고의 일부만 크게 잘라 길 안내 화살표로. 로고 인지도를 길찾기로 바꾼 옥외 광고.', "McDonald's Follow the Arches billboard"],
    ['bk-moldy', 'Moldy Whopper', 'Burger King', 2020, 'INGO, David Miami, Publicis', '식음료', ['print', 'ooh', 'film'], ['shock', 'hero', 'photo'], ['도발', '진지함'], ['#3f5d2a', '#d9c7a0', '#8a5a2b', '#f2f2ee'], '곰팡이 핀 버거를 정면 히어로 컷으로. "방부제 없음"을 가장 불편한 이미지로 증명한다.', 'Burger King Moldy Whopper'],
    ['bk-burning-stores', 'Burning Stores', 'Burger King', 2017, 'DAVID São Paulo', '식음료', ['print'], ['shock', 'photo', 'humor'], ['도발', '유머러스'], ['#e86a12', '#2b1a12', '#f2c14e'], '실제 매장 화재 보도사진에 "Flame grilled since 1954". 약점을 직화의 증거로 뒤집었다.', 'Burger King Burning Stores print ad'],
    ['kitkat-break', 'Have a break, have a KitKat', 'KitKat', 1957, 'JWT', '식음료', ['print', 'ooh'], ['copy', 'hero', 'sequence'], ['유머러스', '따뜻함'], ['#d6001c', '#ffffff', '#5b2c1a'], '"쉬는 시간"이라는 상황 자체를 브랜드 자산으로. 수십 년간 변주되는 시각 장치들.', 'KitKat have a break print ad'],
    ['coke-share', 'Share a Coke', 'Coca-Cola', 2011, 'Ogilvy Sydney', '식음료', ['packaging', 'ooh', 'digital'], ['data', 'logo', 'ugc'], ['따뜻함', '대담'], ['#f40009', '#ffffff', '#1a1a1a'], '로고 자리에 사람 이름을. 패키지가 곧 매체이자 선물이 된다.', 'Share a Coke campaign'],
    ['dove-real-beauty', 'Campaign for Real Beauty', 'Dove', 2004, 'Ogilvy & Mather', '패션·뷰티', ['ooh', 'print'], ['photo', 'copy', 'ugc'], ['진지함', '따뜻함'], ['#ffffff', '#1a3d7c', '#e8d5c4'], '모델 대신 평범한 여성의 사진과 체크박스 질문. 보는 사람에게 투표를 요구하는 옥외 광고.', 'Dove Campaign for Real Beauty billboard 2004'],
    ['dove-sketches', 'Real Beauty Sketches', 'Dove', 2013, 'Ogilvy Brazil', '패션·뷰티', ['film', 'digital'], ['photo', 'metaphor'], ['감성', '진지함'], ['#ece6dd', '#6b625b', '#2b2b2b'], '스스로 묘사한 얼굴과 타인이 묘사한 얼굴의 몽타주 비교. 실험 자체가 KV가 된다.', 'Dove Real Beauty Sketches'],
    ['spotify-weird', "Thanks 2016, It's Been Weird", 'Spotify', 2016, 'Spotify 인하우스', '미디어·엔터', ['ooh'], ['data', 'copy', 'type', 'humor'], ['유머러스', '대담'], ['#f037a5', '#cdf564', '#1e3264', '#ffffff'], '실제 청취 데이터에서 뽑은 짧은 농담을 원색 면에 큰 글자로. 데이터 카피 옥외 광고의 시작.', "Spotify Thanks 2016 it's been weird billboard"],
    ['spotify-wrapped', 'Wrapped', 'Spotify', 2016, 'Spotify 인하우스', '미디어·엔터', ['digital'], ['data', 'color', 'type', 'ugc'], ['대담', '실험적'], ['#1ed760', '#ff4632', '#2d46b9', '#f573a0'], '개인 연말 결산을 공유하고 싶은 카드로. 해마다 그래픽 시스템이 바뀌는 데이터 KV.', 'Spotify Wrapped design'],
    ['old-spice', 'The Man Your Man Could Smell Like', 'Old Spice', 2010, 'Wieden+Kennedy Portland', '패션·뷰티', ['film', 'digital'], ['humor', 'hero'], ['유머러스', '대담'], ['#c8102e', '#1d5aa8', '#f2e6d0'], '원테이크로 장면이 계속 바뀌는 과장된 독백. 캐릭터 하나가 브랜드를 갈아엎었다.', 'Old Spice The Man Your Man Could Smell Like'],
    ['snickers-hungry', "You're Not You When You're Hungry", 'Snickers', 2010, 'BBDO New York', '식음료', ['film', 'print', 'ooh'], ['humor', 'copy', 'sequence'], ['유머러스'], ['#5b3a1e', '#2a62b8', '#ffffff'], '배고프면 딴사람이 된다는 한 문장이 끝없는 변주를 만든다. 인쇄물에서도 같은 틀로 반복.', "Snickers You're not you when you're hungry print"],
    ['always-likeagirl', '#LikeAGirl', 'Always', 2014, 'Leo Burnett', '공공·캠페인', ['film', 'digital'], ['photo', 'copy'], ['진지함', '감성'], ['#f2f2f2', '#6a2c91', '#e6007e'], '"여자애처럼"이라는 말의 의미를 실험으로 뒤집었다. 해시태그가 KV의 일부.', 'Always LikeAGirl campaign'],
    ['volvo-epic-split', 'The Epic Split', 'Volvo Trucks', 2013, 'Forsman & Bodenfors', '자동차·모빌리티', ['film'], ['hero', 'shock'], ['대담', '진지함'], ['#e7b35a', '#3a4a5c', '#d9d4c7'], '후진하는 두 트럭 사이에서 다리를 찢는 한 컷. 조향 정밀도를 하나의 장면으로 증명한다.', 'Volvo Trucks Epic Split'],
    ['marmite', 'Love it or Hate it', 'Marmite', 1996, 'DDB London', '식음료', ['film', 'print'], ['copy', 'humor'], ['유머러스'], ['#5a2d0c', '#f6c400', '#c8102e'], '호불호가 갈린다는 약점을 그대로 브랜드 슬로건으로.', 'Marmite love it or hate it ad'],
    ['got-milk', 'got milk?', 'California Milk Processor Board', 1993, 'Goodby, Silverstein & Partners', '식음료', ['film', 'print', 'ooh'], ['copy', 'humor'], ['유머러스'], ['#ffffff', '#111111', '#2a6ebb'], '우유가 없을 때의 곤란함을 보여주는 역발상. 두 단어 질문형 카피.', 'got milk ad'],
    ['milk-mustache', 'Milk Mustache', 'National Fluid Milk Processor', 1995, 'Bozell', '식음료', ['print'], ['photo', 'sequence', 'humor'], ['유머러스', '따뜻함'], ['#ffffff', '#e8e2d8', '#1a1a1a'], '유명인 인물 사진에 하얀 우유 콧수염 하나. 같은 장치로 수백 편을 만든 시리즈.', 'milk mustache ad'],
    ['benetton-colors', 'United Colors of Benetton', 'Benetton', 1989, 'Oliviero Toscani', '패션·뷰티', ['print', 'ooh'], ['photo', 'shock'], ['도발', '진지함'], ['#0f8f3c', '#ffffff', '#111111'], '제품 없이 강렬한 보도사진급 이미지와 초록 로고 박스만. 패션 광고를 사회적 발언으로 바꿨다.', 'United Colors of Benetton Toscani ad'],
    ['wonderbra', 'Hello Boys', 'Wonderbra', 1994, 'TBWA', '패션·뷰티', ['ooh'], ['photo', 'copy', 'humor'], ['대담', '도발'], ['#111111', '#ffffff', '#c9a27e'], '흑백 모델 사진과 두 단어 카피. 옥외 광고 한 장이 사회적 화제가 된 사례.', 'Wonderbra Hello Boys billboard'],
    ['lacoste-species', 'Save Our Species', 'Lacoste', 2018, 'BETC', '패션·뷰티', ['packaging'], ['logo', 'sequence'], ['진지함', '고급'], ['#ffffff', '#00573f', '#1a1a1a'], '악어 로고 자리에 멸종위기종 10종을. 생산 수량을 남은 개체 수와 맞췄다.', 'Lacoste Save Our Species polo'],
    ['mastercard-priceless', 'Priceless', 'Mastercard', 1997, 'McCann', '금융', ['film', 'print'], ['copy', 'sequence'], ['감성', '따뜻함'], ['#eb001b', '#f79e1b', '#ffffff'], '가격이 붙은 물건들 끝에 "값을 매길 수 없는 것". 구조 자체가 무한 변주 장치.', 'Mastercard Priceless ad'],
    ['ikea-lamp', 'Lamp', 'IKEA', 2002, 'Crispin Porter + Bogusky', '리테일·생활', ['film'], ['humor', 'shock'], ['유머러스', '감성'], ['#3b4a5a', '#e7c46a', '#0058a3'], '버려진 램프에 감정이입시키다가 마지막에 그 감정을 비웃는다.', 'IKEA Lamp commercial Spike Jonze'],
    ['ikea-bookbook', 'Bookbook', 'IKEA', 2014, 'BBH Asia Pacific', '리테일·생활', ['film', 'digital'], ['humor', 'metaphor'], ['유머러스', '미니멀'], ['#ffffff', '#0058a3', '#ffda1a'], '종이 카탈로그를 신제품 테크 기기처럼 소개한 패러디.', 'IKEA Bookbook'],
    ['vw-force', 'The Force', 'Volkswagen', 2011, 'Deutsch LA', '자동차·모빌리티', ['film'], ['humor'], ['유머러스', '따뜻함'], ['#1a1a1a', '#e9e9e9', '#3d6ea8'], '다스베이더 복장의 아이와 원격 시동. 기능을 귀여운 순간 하나로.', 'Volkswagen The Force commercial'],
    ['dumb-ways', 'Dumb Ways to Die', 'Metro Trains Melbourne', 2012, 'McCann Melbourne', '공공·캠페인', ['digital', 'film'], ['illust', 'humor'], ['유머러스', '실험적'], ['#f6d32d', '#4fb3e8', '#e8463a', '#76c043'], '귀여운 캐릭터와 잔혹한 가사의 대비. 안전 캠페인을 노래·게임 IP로 키웠다.', 'Dumb Ways to Die'],
    ['kfc-fck', 'FCK', 'KFC', 2018, 'Mother London', '식음료', ['print'], ['logo', 'humor', 'hero'], ['유머러스', '대담'], ['#e4002b', '#ffffff', '#1a1a1a'], '닭이 떨어진 사태에 버킷 로고를 FCK로 바꾼 전면 사과 광고.', 'KFC FCK ad'],
    ['patagonia-dont-buy', "Don't Buy This Jacket", 'Patagonia', 2011, 'Patagonia 인하우스', '패션·뷰티', ['print'], ['copy', 'shock', 'hero'], ['진지함', '도발'], ['#ffffff', '#111111', '#4a6a8a'], '블랙프라이데이 신문 전면에 자사 재킷을 사지 말라는 헤드라인.', "Patagonia Don't Buy This Jacket"],
    ['oatly-copy', '손글씨 카피 광고', 'Oatly', 2014, 'Oatly 인하우스', '식음료', ['ooh', 'packaging', 'print'], ['handmade', 'copy', 'humor'], ['유머러스', '실험적'], ['#ffffff', '#1d2b63', '#111111'], '삐뚤빼뚤한 손글씨와 장황한 셀프디스 카피. 패키지부터 빌보드까지 같은 목소리.', 'Oatly billboard ad'],
    ['sony-balls', 'Balls', 'Sony Bravia', 2005, 'Fallon London', 'IT·전자', ['film'], ['color', 'metaphor'], ['감성', '실험적'], ['#ff2d55', '#ffcc00', '#00b2ff', '#7ed957'], '샌프란시스코 언덕을 튀어 내려가는 수만 개의 컬러 볼. "색"이라는 효익을 실제 장면으로.', 'Sony Bravia Balls'],
    ['guinness-surfer', 'Surfer', 'Guinness', 1999, 'Abbott Mead Vickers BBDO', '주류', ['film'], ['metaphor', 'photo'], ['진지함', '고급'], ['#0d0d0d', '#e9e4d8', '#7a6a4f'], '흑백 파도와 백마. 기다림(따르는 시간)을 서사시로.', 'Guinness Surfer ad'],
    ['uniqlock', 'UNIQLOCK', 'UNIQLO', 2007, 'Projector', '패션·뷰티', ['digital'], ['data', 'sequence'], ['미니멀', '실험적'], ['#ff0000', '#ffffff', '#111111'], '음악·댄스·시계가 반복되는 블로그 위젯. 시간 자체를 광고 매체로 썼다.', 'UNIQLOCK'],
    ['muji-horizon', 'Horizon', 'MUJI', 2003, 'Kenya Hara', '리테일·생활', ['poster', 'print'], ['negative', 'photo'], ['미니멀', '고급'], ['#d9dde0', '#a7b2ba', '#5e6a73'], '수평선 하나와 작은 사람. 비어 있음을 브랜드 철학으로 보여준다.', 'MUJI Horizon Kenya Hara poster'],
    ['bass-anatomy', 'Anatomy of a Murder', '영화 포스터', 1959, 'Saul Bass', '미디어·엔터', ['poster'], ['geometric', 'illust', 'silhouette'], ['대담', '레트로'], ['#e07b2a', '#111111', '#f2efe6'], '조각난 시신 실루엣과 손글씨 타이틀. 컷아웃 그래픽 한 장으로 장르를 말한다.', 'Saul Bass Anatomy of a Murder poster'],
    ['bass-vertigo', 'Vertigo', '영화 포스터', 1958, 'Saul Bass', '미디어·엔터', ['poster'], ['geometric', 'illust', 'type'], ['대담', '레트로'], ['#e8572a', '#111111', '#ffffff'], '나선과 추락하는 실루엣. 영화의 심리를 기하 형태로 번역.', 'Saul Bass Vertigo poster'],
    ['glaser-iny', 'I ♥ NY', 'New York State', 1977, 'Milton Glaser', '공공·캠페인', ['identity', 'print'], ['type', 'logo'], ['따뜻함', '대담'], ['#ffffff', '#111111', '#e4002b'], '글자 사이 빨간 하트 하나. 가장 많이 변주된 도시 아이덴티티.', 'I love NY Milton Glaser'],
    ['glaser-dylan', 'Dylan', 'Columbia Records', 1966, 'Milton Glaser', '미디어·엔터', ['poster'], ['silhouette', 'illust', 'color'], ['레트로', '실험적'], ['#111111', '#f7a81b', '#e2365a', '#3aa6d0'], '검은 옆얼굴 실루엣과 사이키델릭한 색 머리카락.', 'Milton Glaser Dylan poster'],
    ['obama-hope', 'HOPE', 'Obama 2008 (비공식)', 2008, 'Shepard Fairey', '공공·캠페인', ['poster'], ['illust', 'color', 'type'], ['대담', '진지함'], ['#d71a28', '#0f3a5d', '#f2e4c0', '#7ea2b3'], '사진을 4색 평면 스텐실로 단순화하고 한 단어를 아래에. 정치 포스터의 시각 문법이 됐다.', 'Shepard Fairey Hope poster'],
    ['mb-der-film', 'Der Film', 'Kunstgewerbemuseum Zürich', 1960, 'Josef Müller-Brockmann', '문화·예술', ['poster'], ['grid', 'type'], ['미니멀', '진지함'], ['#111111', '#ffffff', '#9a9a9a'], '검정 바탕, 산세리프, 엄격한 그리드. 스위스 스타일 포스터의 정수.', 'Muller-Brockmann Der Film poster'],
    ['mb-beethoven', 'Beethoven', 'Tonhalle Zürich', 1955, 'Josef Müller-Brockmann', '문화·예술', ['poster'], ['geometric', 'grid'], ['미니멀', '실험적'], ['#f2f0ea', '#111111', '#7d7d7d'], '동심원 호가 음악의 리듬처럼 퍼진다. 소리를 기하 구성으로 시각화.', 'Muller-Brockmann Beethoven poster 1955'],
    ['cassandre-nord', 'Nord Express', '철도 포스터', 1927, 'A.M. Cassandre', '자동차·모빌리티', ['poster'], ['geometric', 'illust'], ['레트로', '대담'], ['#1d2f4d', '#c9b68c', '#111111', '#e8e2d0'], '극단적 원근의 기관차 바퀴와 레일. 아르데코 속도감의 원형.', 'Cassandre Nord Express poster'],
    ['cassandre-dubonnet', 'Dubo Dubon Dubonnet', 'Dubonnet', 1932, 'A.M. Cassandre', '주류', ['poster'], ['sequence', 'illust', 'type'], ['레트로', '유머러스'], ['#e9dfc8', '#c8102e', '#111111'], '세 칸에 걸쳐 인물이 채워지고 단어도 완성된다. 연속 컷으로 만든 브랜드 네임 각인.', 'Cassandre Dubonnet poster'],
    ['rodchenko-books', 'Books (Lengiz)', 'Lengiz 출판사', 1924, 'Alexander Rodchenko', '문화·예술', ['poster'], ['geometric', 'photo', 'type', 'collage'], ['대담', '레트로'], ['#d52b1e', '#111111', '#efe6d2'], '외치는 여성 사진과 사선 타이포. 구성주의 포스터의 대표작.', 'Rodchenko Books poster 1924'],
    ['tanaka-nihon-buyo', 'Nihon Buyo', 'UCLA Asian Performing Arts', 1981, 'Ikko Tanaka', '문화·예술', ['poster'], ['geometric', 'color'], ['미니멀', '고급'], ['#2c6fb5', '#f4c430', '#d64d3c', '#111111', '#f2efe8'], '사각·원·면 색만으로 게이샤 얼굴을 구성. 일본 모던 그래픽의 상징.', 'Ikko Tanaka Nihon Buyo poster'],
    ['scher-public-theater', 'The Public Theater', 'The Public Theater', 1994, 'Paula Scher (Pentagram)', '문화·예술', ['poster', 'identity'], ['type', 'collage'], ['대담', '실험적'], ['#ffffff', '#111111', '#f2c500', '#c8102e'], '크기와 굵기가 제각각인 나무활자풍 타이포를 사진과 충돌시킨다. 길거리에서 읽히는 극장 포스터.', 'Paula Scher Public Theater poster'],
    ['sagmeister-aiga', 'AIGA Detroit', 'AIGA', 1999, 'Stefan Sagmeister', '문화·예술', ['poster'], ['handmade', 'shock', 'type'], ['도발', '실험적'], ['#e6d3c3', '#7a1f1f', '#111111'], '강연 정보를 피부에 직접 새겨 촬영. 디자인의 고통을 글자 그대로 보여준다.', 'Sagmeister AIGA Detroit poster'],
    ['polish-school', '폴란드 포스터 학파', '영화·공연 포스터', 1955, 'Henryk Tomaszewski, Jan Lenica 등', '문화·예술', ['poster'], ['illust', 'handmade', 'metaphor'], ['실험적', '레트로'], ['#e94b3c', '#f2e3c6', '#1f3a5f', '#111111'], '영화 장면 대신 은유적 그림과 손글씨. 원작을 재해석하는 회화적 포스터.', 'Polish school of posters'],
    ['bauhaus-exhibition', '바우하우스 전시 포스터', 'Bauhaus', 1923, 'Joost Schmidt', '문화·예술', ['poster'], ['geometric', 'type'], ['레트로', '실험적'], ['#c8102e', '#111111', '#e9dcc0', '#1f4e9a'], '원·사각·선의 구성과 산세리프. 모던 그래픽 디자인의 출발점.', 'Bauhaus Ausstellung 1923 poster Joost Schmidt'],
    ['baemin-type', '배민 타이포 광고', '배달의민족', 2014, '우아한형제들', '식음료', ['ooh', 'print', 'digital'], ['type', 'humor', 'copy'], ['유머러스', '대담'], ['#2ac1bc', '#ffffff', '#111111'], '전용 서체로 쓴 말장난 카피를 크게. 서체 자체가 브랜드 자산이 된 국내 대표 사례.', '배달의민족 한나체 광고'],
    ['bk-whopper-detour', 'Whopper Detour', 'Burger King', 2018, 'FCB New York', '식음료', ['digital'], ['data', 'humor', 'shock'], ['유머러스', '도발'], ['#d62300', '#f5ebdc', '#502314'], '경쟁사 매장 근처에서만 1센트 와퍼를 여는 앱 쿠폰. 위치 데이터가 아이디어.', 'Burger King Whopper Detour'],
    ['guardian-pigs', 'Three Little Pigs', 'The Guardian', 2012, 'BBH London', '미디어·엔터', ['film'], ['metaphor'], ['진지함'], ['#052962', '#ffffff', '#c70000'], '동화를 뉴스 보도로 재구성해 "전체 그림"을 보여주는 저널리즘을 설명.', 'Guardian Three Little Pigs ad'],
    ['airbnb-belo', 'Bélo', 'Airbnb', 2014, 'DesignStudio', '브랜드 아이덴티티', ['identity'], ['logo', 'illust'], ['따뜻함', '미니멀'], ['#ff5a5f', '#ffffff', '#484848'], '사람·장소·사랑·A를 합친 하나의 심볼. 누구나 그릴 수 있는 로고를 의도.', 'Airbnb Belo logo DesignStudio'],
    ['mailchimp-rebrand', '리브랜딩', 'Mailchimp', 2018, 'Collins', '브랜드 아이덴티티', ['identity', 'digital'], ['illust', 'color', 'type'], ['유머러스', '따뜻함'], ['#ffe01b', '#241c15', '#ffffff'], '캐번디시 옐로와 거친 손그림 일러스트. B2B 툴을 개성 있는 브랜드로.', 'Mailchimp rebrand Collins'],
    ['tide-ad', "It's a Tide Ad", 'Tide', 2018, 'Saatchi & Saatchi New York', '리테일·생활', ['film'], ['humor', 'shock'], ['유머러스'], ['#f7941d', '#1f3f99', '#ffffff'], '깨끗한 옷이 나오면 모든 광고가 타이드 광고라는 역발상. 다른 광고의 클리셰를 모두 빌려 쓴다.', "It's a Tide Ad Super Bowl"],
  ];

  const KV = K.map(function (r) {
    const field = r[6].indexOf('identity') >= 0 ? '아이덴티티' : r[5] === '공공·캠페인' ? '공익·사회' : r[5] === '문화·예술' || (r[6].indexOf('poster') >= 0 && r[5] === '미디어·엔터') ? '문화·공연' : '브랜드';
    return { id: r[0], title: r[1], brand: r[2], year: r[3], by: r[4], ind: r[5], med: r[6], tech: r[7], mood: r[8], col: r[9], note: r[10], q: r[11], country: KV_COUNTRY[r[0]] || 'XX', field: field, origin: 'classic' };
  });

  // [분류, 이름, 가로, 세로, 단위, 메모]
  const SPECS = [
    ['디스플레이 배너 (IAB)', 'Medium Rectangle', 300, 250, 'px', '가장 흔한 디스플레이 지면'],
    ['디스플레이 배너 (IAB)', 'Large Rectangle', 336, 280, 'px', ''],
    ['디스플레이 배너 (IAB)', 'Leaderboard', 728, 90, 'px', '데스크톱 상단'],
    ['디스플레이 배너 (IAB)', 'Billboard', 970, 250, 'px', '데스크톱 대형'],
    ['디스플레이 배너 (IAB)', 'Half Page', 300, 600, 'px', ''],
    ['디스플레이 배너 (IAB)', 'Wide Skyscraper', 160, 600, 'px', ''],
    ['디스플레이 배너 (IAB)', 'Mobile Leaderboard', 320, 50, 'px', '모바일 하단 고정'],
    ['디스플레이 배너 (IAB)', 'Large Mobile Banner', 320, 100, 'px', ''],
    ['소셜', '정사각 피드', 1080, 1080, 'px', '1:1'],
    ['소셜', '세로 피드', 1080, 1350, 'px', '4:5, 인스타그램 피드 최대 세로'],
    ['소셜', '스토리·릴스·쇼츠', 1080, 1920, 'px', '9:16, 상하 안전 영역 확보'],
    ['소셜', '링크 공유 이미지', 1200, 630, 'px', '약 1.91:1 (OG 이미지)'],
    ['소셜', 'YouTube 썸네일', 1280, 720, 'px', '16:9'],
    ['인쇄', 'A1 포스터', 594, 841, 'mm', '1:√2'],
    ['인쇄', 'A2 포스터', 420, 594, 'mm', ''],
    ['인쇄', 'A3', 297, 420, 'mm', ''],
    ['인쇄', 'A4 지면', 210, 297, 'mm', '잡지·전단'],
    ['인쇄', 'B1 포스터 (ISO)', 707, 1000, 'mm', ''],
    ['옥외', '6-sheet (영국)', 1200, 1800, 'mm', '버스 쉘터·가로 포스터'],
    ['옥외', '48-sheet (영국)', 6096, 3048, 'mm', '대형 빌보드 2:1'],
  ];

  // 다국어 KV 키워드 (검색 확장용)
  const TERMS = [
    ['en', 'key visual'], ['en', 'print ad'], ['en', 'poster design'], ['en', 'billboard'], ['en', 'banner design'], ['en', 'campaign'],
    ['ko', '키비주얼'], ['ko', '광고 포스터'], ['ko', '배너 디자인'],
    ['ja', 'キービジュアル'], ['ja', 'ポスター'], ['ja', '広告'],
    ['zh', '主视觉'], ['zh', '海报'], ['zh', '广告设计'],
    ['ru', 'ключевой визуал'], ['ru', 'рекламный плакат'],
  ];

  return { CATEGORIES: CATEGORIES, SOURCES: SOURCES, TECHNIQUES: TECHNIQUES, MOODS: MOODS, INDUSTRIES: INDUSTRIES, FIELDS: FIELDS, COUNTRIES: COUNTRIES, ORIGINS: ORIGINS, PROMPT_TARGETS: PROMPT_TARGETS, MEDIA: MEDIA, KV: KV, SPECS: SPECS, TERMS: TERMS };
})();
