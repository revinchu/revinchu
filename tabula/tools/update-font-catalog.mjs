/** Build-time font metadata refresh. No font binaries or workbook text. */
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
const root = fileURLToPath(new URL('../', import.meta.url));
const args = process.argv.slice(2);
if (args.includes('--help')) {
  console.log('node tools/update-font-catalog.mjs [--cache-dir DIR] [--output FILE] [--date YYYY-MM-DD] [--offline | --reuse-cache] [--verify-network]');
  process.exit(0);
}
const value = name => { const i = args.indexOf(name); return i < 0 ? null : args[i + 1]; };
const cache = path.resolve(value('--cache-dir') || 'D:/Codex/Temp/wixel-font-catalog');
const output = path.resolve(value('--output') || path.join(root, 'src/web-font-catalog.js'));
const offline = args.includes('--offline');
const updatedAt = value('--date') || new Date().toISOString().slice(0, 10);
const GOOGLE_METADATA = 'https://fonts.google.com/metadata/fonts';
const GOOGLE_TREE = 'https://api.github.com/repos/google/fonts/git/trees/main?recursive=1';
const sha256 = text => createHash('sha256').update(text).digest('hex');
const decode = text => JSON.parse(text.replace(/^\)\]\}'\s*/, ''));
await fs.mkdir(cache, { recursive: true });
async function readMetadata(name, url) {
  const filename = path.join(cache, name);
  if (offline) return fs.readFile(filename, 'utf8');
  if (args.includes('--reuse-cache')) { try { return await fs.readFile(filename, 'utf8'); } catch {} }
  const response = await fetch(url, { signal: AbortSignal.timeout(30000), headers: { Accept: 'application/json' } });
  if (!response.ok) throw new Error(url + ': HTTP ' + response.status);
  const text = await response.text();
  decode(text); // Never cache an error/login HTML page.
  await fs.writeFile(filename, text);
  return text;
}
const aliases = {
  'Asta Sans':['아스타 산스'],'Bagel Fat One':['베이글 팻 원'],
  'Black And White Picture':['흑백사진'],'Black Han Sans':['검은고딕'],
  'Cute Font':['귀여운 글씨'],'Diphylleia':['산하엽'],
  'Do Hyeon':['배민 도현체','배달의민족 도현'],'Dokdo':['독도체'],
  'Dongle':['동글'],'East Sea Dokdo':['동해 독도'],'Gaegu':['개구'],
  'Gamja Flower':['감자꽃'],'Gasoek One':['가석'],'Gothic A1':['고딕 A1'],
  'Gowun Batang':['고운바탕','고운 바탕'],'Gowun Dodum':['고운돋움','고운 돋움'],
  'Grandiflora One':['그랜디플로라'],'Gugi':['구기'],'Hahmlet':['함렛'],
  'Hi Melody':['하이 멜로디'],'IBM Plex Sans KR':['IBM 플렉스 산스 KR'],
  'Jua':['배민 주아체','배달의민족 주아'],'Kirang Haerang':['배민 기랑해랑체'],
  'Moirai One':['모이라이'],'Nanum Brush Script':['나눔손글씨 붓','나눔 붓글씨'],
  'Nanum Gothic':['나눔고딕','나눔 고딕','NanumGothic'],
  'Nanum Gothic Coding':['나눔고딕코딩','NanumGothicCoding'],
  'Nanum Myeongjo':['나눔명조','나눔 명조','NanumMyeongjo'],
  'Nanum Pen Script':['나눔손글씨 펜','나눔 펜글씨'],
  'Noto Sans KR':['본고딕','노토 산스 한국어'],'Noto Serif KR':['본명조','노토 세리프 한국어'],
  'Orbit':['오르빗'],'Poor Story':['푸어스토리'],'Single Day':['싱글데이'],
  'Song Myung':['송명'],'Stylish':['스타일리시'],'Sunflower':['해바라기'],
  'Yeon Sung':['배민 연성체','배달의민족 연성']
};
const categories = {'Sans Serif':'sans-serif',Serif:'serif',Display:'display',Handwriting:'handwriting',Monospace:'monospace'};
const rawMetadata = await readMetadata('google-metadata.json', GOOGLE_METADATA);
const rawTree = await readMetadata('google-tree.json', GOOGLE_TREE);
const metadata = decode(rawMetadata), tree = decode(rawTree);
if (tree.truncated || !/^[a-f\d]{40}$/.test(tree.sha) || !Array.isArray(metadata.familyMetadataList)) throw new Error('Incomplete Google metadata/tree');
const treePaths = new Set(tree.tree.map(item => item.path));
const directories = new Map(tree.tree.filter(item => item.path.endsWith('/METADATA.pb')).map(item => [item.path.split('/')[1],item.path.slice(0,-12)]));
const rows = [], excluded = [], fallbackLicenses = [];
for (const family of metadata.familyMetadataList) {
  if (!family.isOpenSource || /^(?:Noto (?:Color )?Emoji|Material (?:Icons|Symbols))/.test(family.family)) { excluded.push(family.family); continue; }
  const dir = directories.get(family.family.toLowerCase().replace(/[^a-z0-9]/g,''));
  const licensePath = dir && ['OFL.txt','LICENSE.txt','LICENCE.txt','LICENSE'].map(name => dir+'/'+name).find(name => treePaths.has(name));
  let fallback;
  if (!licensePath) {
    // New served families may precede google/fonts. Only read manifest text.
    const url = 'https://fonts.google.com/download/list?family='+encodeURIComponent(family.family);
    const raw = await readMetadata('license-'+sha256(family.family).slice(0,16)+'.json',url);
    const licenseFile = decode(raw).manifest?.files?.find(item => /^(?:OFL|LICENSE|LICENCE)(?:\.txt)?$/i.test(item.filename));
    if (!licenseFile || !/SIL Open Font License|Apache License|UBUNTU FONT LICENCE/i.test(licenseFile.contents || '')) throw new Error('Unverified font license: '+family.family);
    fallback = {url,copyright:licenseFile.contents.split(/\r?\n/)[0],sha256:sha256(licenseFile.contents),text:licenseFile.contents,license:/SIL Open Font License/i.test(licenseFile.contents)?'OFL-1.1':/Apache License/i.test(licenseFile.contents)?'Apache-2.0':'UFL-1.0'};
    const {text:licenseText,...proof}=fallback;
    fallbackLicenses.push({family:family.family,...proof});
  }
  if (typeof family.family!=='string' || !family.family.trim() || /[<>\x00-\x1f]/.test(family.family)) throw new Error('Invalid font family');
  const variants = Object.keys(family.fonts);
  if (!variants.length || variants.some(v => !/^\d+i?$/.test(v) || parseInt(v,10)<1 || parseInt(v,10)>1000)) throw new Error('Unexpected variant metadata: '+family.family);
  rows.push([family.family,categories[family.category]||'display',family.subsets.filter(s=>s!=='menu'),variants,licensePath||'',aliases[family.family]||[],family.designers||[],fallback||null]);
}
// Curated entries require original license evidence and verified HTTPS/CORS delivery.
const curated = [
  {
    "family": "Pretendard",
    "label": "Pretendard",
    "aliases": [
      "프리텐다드"
    ],
    "source": "cdn",
    "category": "sans-serif",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      100,
      200,
      300,
      400,
      500,
      600,
      700,
      800,
      900
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "100",
      "200",
      "300",
      "400",
      "500",
      "600",
      "700",
      "800",
      "900"
    ],
    "license": "OFL-1.1",
    "cssUrl": "https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard-dynamic-subset.min.css",
    "licenseUrl": "https://github.com/orioncactus/pretendard/blob/v1.3.9/LICENSE",
    "rawLicenseUrl": "https://raw.githubusercontent.com/orioncactus/pretendard/v1.3.9/LICENSE",
    "homepage": "https://github.com/orioncactus/pretendard",
    "version": "1.3.9",
    "authors": [
      "Kil Hyung-jin"
    ]
  },
  {
    "family": "SUIT",
    "label": "SUIT",
    "aliases": [
      "수트"
    ],
    "source": "cdn",
    "category": "sans-serif",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      100,
      200,
      300,
      400,
      500,
      600,
      700,
      800,
      900
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "100",
      "200",
      "300",
      "400",
      "500",
      "600",
      "700",
      "800",
      "900"
    ],
    "license": "OFL-1.1",
    "cssUrl": "https://cdn.jsdelivr.net/gh/sun-typeface/SUIT@v2.0.5/fonts/static/woff2/SUIT.css",
    "licenseUrl": "https://github.com/sun-typeface/SUIT/blob/v2.0.5/LICENSE",
    "rawLicenseUrl": "https://raw.githubusercontent.com/sun-typeface/SUIT/v2.0.5/LICENSE",
    "homepage": "https://sun.fo/suit/",
    "version": "2.0.5",
    "authors": [
      "SUNN.US"
    ]
  },
  {
    "family": "Wanted Sans",
    "label": "Wanted Sans",
    "aliases": [
      "원티드 산스",
      "WantedSans"
    ],
    "source": "cdn",
    "category": "sans-serif",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400,
      500,
      600,
      700,
      800,
      900,
      1000
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400",
      "500",
      "600",
      "700",
      "800",
      "900",
      "1000"
    ],
    "license": "OFL-1.1",
    "cssUrl": "https://cdn.jsdelivr.net/gh/wanteddev/wanted-sans@v1.0.3/packages/wanted-sans/fonts/webfonts/static/split/WantedSans.min.css",
    "licenseUrl": "https://github.com/wanteddev/wanted-sans/blob/v1.0.3/OFL.txt",
    "rawLicenseUrl": "https://raw.githubusercontent.com/wanteddev/wanted-sans/v1.0.3/OFL.txt",
    "homepage": "https://github.com/wanteddev/wanted-sans",
    "version": "1.0.3",
    "authors": [
      "Wanted Lab"
    ]
  },
  {
    "family": "NanumSquareNeo",
    "label": "NanumSquareNeo",
    "aliases": [
      "나눔스퀘어 네오",
      "나눔스퀘어네오",
      "Nanum Square Neo"
    ],
    "source": "cdn",
    "category": "sans-serif",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      300,
      400,
      700,
      800,
      900
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "300",
      "400",
      "700",
      "800",
      "900"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://hangeul.pstatic.net/hangeul_static/webfont/NanumSquareNeo/NanumSquareNeoTTF-aLt.woff",
        "weight": 300,
        "style": "normal"
      },
      {
        "url": "https://hangeul.pstatic.net/hangeul_static/webfont/NanumSquareNeo/NanumSquareNeoTTF-bRg.woff",
        "weight": 400,
        "style": "normal"
      },
      {
        "url": "https://hangeul.pstatic.net/hangeul_static/webfont/NanumSquareNeo/NanumSquareNeoTTF-cBd.woff",
        "weight": 700,
        "style": "normal"
      },
      {
        "url": "https://hangeul.pstatic.net/hangeul_static/webfont/NanumSquareNeo/NanumSquareNeoTTF-dEb.woff",
        "weight": 800,
        "style": "normal"
      },
      {
        "url": "https://hangeul.pstatic.net/hangeul_static/webfont/NanumSquareNeo/NanumSquareNeoTTF-eHv.woff",
        "weight": 900,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://help.naver.com/service/30016/contents/18088?osType=PC",
    "homepage": "https://campaign.naver.com/nanumsquare_neo/",
    "authors": [
      "NAVER"
    ],
    "copyright": "Copyright NAVER. Reserved Font Name NanumSquareNeo.",
    "licenseText": "네이버 나눔글꼴의 지적 재산권은 네이버와 네이버 문화재단에 있습니다. \n네이버 나눔글꼴은 개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 글꼴 자체를 유료로 판매하는 것을 제외한 상업적인 사용이 가능합니다.\n네이버 나눔글꼴은 본 저작권 안내와 라이선스 전문을 포함해서 다른 소프트웨어와 번들하거나 재배포 또는 판매가 가능하고 자유롭게 수정, 재배포하실 수 있습니다.\n네이버 나눔글꼴 라이선스 전문을 포함하기 어려울 경우, 나눔글꼴의 출처 표기를 권장합니다. \n예) 이 페이지에는 네이버에서 제공한 나눔글꼴이 적용되어 있습니다.\n네이버 나눔글꼴을 사용한 인쇄물, 광고물(온라인 포함)의 이미지는 나눔글꼴 프로모션을 위해 활용될 수 있습니다.\n이를 원치 않는 사용자는 언제든지 당사에 요청하실 수 있습니다.\n정확한 사용 조건은 아래 네이버 나눔글꼴 라이선스 전문을 참고하시기 바랍니다.\n네이버 마루 부리 글꼴도 나눔 글꼴과 같은 오픈라이센스 폰트로 동일한 저작권 규정을 적용받습니다.\n※ 참고해 주세요!\n- 나눔글꼴 라이선스 전문의 한글은 이용자의 이해를 돕기 위해 영문 원본을 번역해 서비스하고 있으며, 법적 효력은 영문에 한합니다.\n1. 라이선스\nCopyright (c) 2010, NAVER Corporation (https://www.navercorp.com/) with Reserved Font Name Nanum, Naver Nanum, NanumGothic, Naver NanumGothic, NanumMyeongjo, Naver NanumMyeongjo, NanumBrush, Naver NanumBrush, NanumPen, Naver NanumPen, Naver NanumGothicEco, NanumGothicEco, Naver NanumMyeongjoEco, NanumMyeongjoEco, Naver NanumGothicLight, NanumGothicLight, NanumBarunGothic, Naver NanumBarunGothic, NanumSquareRound, NanumBarunPen, MaruBuri, NanumSquareNeo\nThis Font Software is licensed under the SIL Open Font License, Version 1.1.\nThis license is copied below, and is also available with a FAQ at: http://scripts.sil.org/OFL\nSIL OPEN FONT LICENSE\nVersion 1.1 - 26 February 2007 \n‘나눔, 네이버 나눔, 나눔고딕, 네이버 나눔고딕, 나눔명조, 네이버 나눔명조, 나눔손글씨, 네이버 나눔손글씨, 나눔펜, 네이버 나눔펜, 네이버 나눔고딕에코, 나눔고딕에코, 네이버 나눔명조에코, 나눔명조에코, 네이버 나눔고딕라이트, 나눔고딕라이트, 나눔바른고딕, 네이버나눔바른고딕, 나눔스퀘어라운드, 나눔바른펜, 마루 부리, 나눔스퀘어네오’ 폰트명에 대해 NAVER(https://www.navercorp.com/)가 저작권을 소유하고 있습니다.\n본 폰트 소프트웨어는 SIL 오픈 폰트 라이선스 버전 1.1에 따라 라이선스 취득을 하였습니다.\n본 라이선스는 하단에 복사되었고 http://scripts.sil.org/OFL의 FAQ 란에서도 열람 가능합니다. \nSIL 오픈 폰트 라이선스\n버전 1.1 (2007년 2월 26일)    \n[DEFINITIONS (정의)]\n\"\"Font Software\"\" refers to the set of files released by the Copyright Holder(s) under this license and clearly marked as such. This may include source files, build scripts and documentation.\n\"\"Reserved Font Name\"\" refers to any names specified as such after the copyright statement(s).\n\"\"Original Version\"\" refers to the collection of Font Software components as distributed by the Copyright Holder(s).\n\"\"Modified Version\"\" refers to any derivative made by adding to, deleting, or substituting in part or in whole any of the components of the Original Version, by changing formats or by porting the Font Software to a new environment.‘\n\"\"Author\"\" refers to any designer, engineer, programmer, technical writer or other person who contributed to the Font Software.  \n‘폰트 소프트웨어’는 본 라이선스에 입거해 저작권자가 배포하고 명확하게 같은 표시가 된 파일들의 집합을 뜻하며, 여기에는 소스 파일, 빌드 스크립트와 문서가 이에 포함됩니다. \n‘저작권이 있는 폰트명’은 저작권 정책에 따라서 지정된 이름을 말합니다. \n‘원본’은 저작권자가 배포한 폰트 소프트웨어 구성요소를 의미합니다. \n‘수정본’은 포맷의 변경이나 폰트 소프트웨어를 새로운 환경에 포팅 시켜, 원본의 일부 혹은 전체에 추가, 삭제 대체해 만든 파생 저작물을 의미합니다.\n‘저자’는 폰트 소프트웨어에 기여한 디자이너, 엔지니어, 프로그래머, 기술 전문가 등을 의미합니다.\n2. PREAMBLE (전문)\nThe goals of the Open Font License (OFL) are to stimulate worldwide development of collaborative font projects, to support the font creation efforts of academic and linguistic communities, and to provide a free and open framework  in which fonts may be shared and improved in partnership with others.\nThe OFL allows the licensed fonts to be used, studied, modified and redistributed freely as long as they are not sold by themselves. The fonts, including any derivative works, can be bundled, embedded, redistributed and/or sold with any software provided that any reserved names are not used by derivative works.\nThe fonts and derivatives, however, cannot be released under any other type of license.\nThe requirement for fonts to remain under this license does not apply to any document created using the fonts or their derivatives.  \n본 폰트 라이선스를 오픈하는 것은(이하 OFL)는 전 세계 폰트 개발 프로젝트를 지원하고 학계와 언어 관련 학계의 폰트 개발을 위한 연구를 지지하기 위해서인 동시에, 폰트 제휴를 통해 폰트가 공유되고 개선될 수 있는 자유롭게 개방된 환경을 만들기 위해서입니다.  \nOFL은 라이선스를 취득한 폰트가 그 자체로 판매되지 않는 한 자유롭게 사용, 연구, 수정, 재배포 하는 것을 허가합니다.\n수정된 폰트를 포함한 폰트는 저작권 명이 사용되지 않는 한 기타 소프트웨어와 함께 묶이거나 삽입, 재배포 할 수 있습니다.\n단 폰트와 수정된 폰트는 기타 다른 라이선스에 포함되어 배포될 수는 없습니다.\n이 라이선스 하에 있기 위한 폰트에 대한 요구사항은 본 폰트나 수정본을 사용하여 제작된 어떠한 문서에도 적용되지 않습니다. \n3. PERMISSION & CONDITIONS (허가 및 조건)\nPermission is hereby granted, free of charge, to any person obtaining a copy of the Font Software, to use, study, copy, merge, embed, modify, redistribute, and sell modified and unmodified copies of the Font Software, subject to the following conditions:  \n본 폰트 소프트웨어를 사용하도록 허가받은 개인/기업/단체 누구라도 다음 명시된 조건에 따라 폰트 소프트웨어의 수정 혹은 수정되지 않은 복사본을 무료로 사용, 연구, 복사, 통합, 삽입, 수정, 재배포할 수 있도록 허가합니다.\n1) Neither the Font Software nor any of its individual components, in Original or Modified Versions, may be sold by itself.\n원본이나 수정본의 폰트 소프트웨어 혹은 개별 구성요소인 폰트 자체가 판매되어서는 안됩니다.\n2) Original or Modified Versions of the Font Software may be bundled, redistributed and/or sold with any software, provided that each copy contains the above copyright notice and this license.\nThese can be included either as stand-alone text files, human-readable headers or in the appropriate machine-readable metadata fields within text or binary files as long as those fields can be easily viewed by the user.  \n본 폰트 소프트웨어의 원본 혹은 수정본은 상기 저작권 안내와 본 라이선스에 대한 내용을 포함하는 경우에는 다른 소프트웨어와 함께 묶이거나 재배포 혹은 판매가 가능합니다.\n이는 독립 텍스트 파일과 가독성이 있는 헤더 혹은 유저가 용이하게 열람 가능한 이상 텍스트 파일 혹은 이진 파일 내 기계가 읽을 수 있는 메타데이터 형태를 모두 의미합니다.  \n3) No Modified Version of the Font Software may use the Reserved Font Name(s) unless explicit written permission is granted by the corresponding Copyright Holder.\nThis restriction only applies to the primary font name as presented to the users.  \n본 폰트 소프트웨어의 어떠한 수정본도 동일한 저작권자가 명시적 허가서를 부여하지 않는 한 저작권이 있는 폰트명을 사용해서는 안 됩니다.\n본 제한 사항은 유저들에게 제공된 기존 폰트명을 뜻합니다.\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font Software shall not be used to promote, endorse or advertise any Modified Version, except to acknowledge the contribution(s) of the Copyright Holder(s) and the Author(s) or with their explicit written permission.  \n본 폰트 소프트웨어의 저작권자 혹은 저자의 이름은 그들의 명시적 서면 허가가 있거나 또는 그들의 공헌을 인정하기 위한 경우를 제외하고는 수정본에 대한 사용을 유도, 추천 혹은 광고하기 위한 목적으로 사용할 수 없습니다.  \n5) The Font Software, modified or unmodified, in part or in whole, must be distributed entirely under this license, and must not be distributed any other license.\nThe requirement for fonts to remain under this license does not apply to any document created using the Font Software.\n본 폰트 소프트웨어는 전체나 부분, 혹은 수정 여부에 상관없이 본 라이선스 하에 배포가 되어야 하며 기타 다른 라이선스 하에서는 배포를 할 수 없습니다.\n폰트에 대한 요구 조건은 이 라이선스 하에서만 유효하며 이 라이선스 하에 있기 위한 폰트에 대한 요구사항은 본 폰트 소프트웨어를 사용해 제작한 어떠한 문서에도 적용되지 않습니다.\n4. TERMINATION (계약의 종료)\nThis license becomes null and void if any of the above conditions are not met. \n본 라이선스는 상기 조건 중 일부라도 부합되지 않으면 무효가 될 수 있습니다.\n5. DISCLAIMER (면책조항)\nTHE FONT SOFTWARE IS PROVIDED \"\"AS IS\"\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT OF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT.\nIN NO EVENT SHALL THE COPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, INCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL DAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM OTHER DEALINGS IN THE FONT SOFTWARE.  \n본 폰트 소프트웨어는 저작권, 특허권, 상표권 및 기타 권리의 비침해성과 특정 목적에의 적합성 포함한 명시적, 묵시적인 어떠한 종류의 보증 없이 “있는 그대로” 제공됩니다.\n어떠한 경우에도 저작권자는 본 폰트 소프트웨어의 사용 또는 이의 사용 불가, 그밖에 폰트 소프트웨어의 취급과 관련하여 발생하는 모든 계약, 불법행위 혹은 다른 일로 하여금 발생하는 일반적, 특수적, 간접적, 부차적 혹은 필연적 손해를 포함하는 소송, 손해, 혹은 기타 책임에 대한 의무를 가지지 않습니다."
  },
  {
    "family": "Gmarket Sans",
    "label": "Gmarket Sans",
    "aliases": [
      "G마켓 산스",
      "지마켓 산스",
      "GmarketSans"
    ],
    "source": "cdn",
    "category": "sans-serif",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      300,
      500,
      700
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "300",
      "500",
      "700"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://script.gmarket.com/fonts/GmarketSansLight.woff2",
        "weight": 300,
        "style": "normal"
      },
      {
        "url": "https://script.gmarket.com/fonts/GmarketSansMedium.woff2",
        "weight": 500,
        "style": "normal"
      },
      {
        "url": "https://script.gmarket.com/fonts/GmarketSansBold.woff2",
        "weight": 700,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://corp.gmarket.com/fonts/",
    "homepage": "https://corp.gmarket.com/fonts/",
    "authors": [
      "Gmarket"
    ],
    "copyright": "Copyright © 2019 eBay Korea Co., Ltd. All rights reserved.",
    "licenseText": "Copyright © 2019 eBay Korea Co., Ltd. All rights reserved.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://script.gmarket.com/fonts/GmarketSansLight.woff"
  },
  {
    "family": "Cafe24ProSlimMax",
    "label": "카페24 PRO Slim Max",
    "aliases": [
      "Cafe24 PRO Slim Max",
      "카페24 PRO Slim Max"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24PROSlim-Bold.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright ⓒ Cafe24 Corp. All Rights Reserved. Cafe24 PRO Slim is a trademark of Cafe24.",
    "version": "2025-11-04",
    "licenseText": "Copyright ⓒ Cafe24 Corp. All Rights Reserved. Cafe24 PRO Slim is a trademark of Cafe24.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24PROSlim-Bold.woff"
  },
  {
    "family": "Cafe24ProSlimFit",
    "label": "카페24 PRO Slim Fit",
    "aliases": [
      "Cafe24 PRO Slim Fit",
      "카페24 PRO Slim Fit"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24PROSlim-Regular.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright ⓒ Cafe24 Corp. All Rights Reserved. Cafe24 PRO Slim is a trademark of Cafe24.",
    "version": "2025-11-04",
    "licenseText": "Copyright ⓒ Cafe24 Corp. All Rights Reserved. Cafe24 PRO Slim is a trademark of Cafe24.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24PROSlim-Regular.woff"
  },
  {
    "family": "Cafe24ProSlimAir",
    "label": "카페24 PRO Slim Air",
    "aliases": [
      "Cafe24 PRO Slim Air",
      "카페24 PRO Slim Air"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24PROSlim-Light.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright ⓒ Cafe24 Corp. All Rights Reserved. Cafe24 PRO Slim is a trademark of Cafe24.",
    "version": "2025-11-04",
    "licenseText": "Copyright ⓒ Cafe24 Corp. All Rights Reserved. Cafe24 PRO Slim is a trademark of Cafe24.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24PROSlim-Light.woff"
  },
  {
    "family": "cafe24Proup",
    "label": "카페24 PRO UP",
    "aliases": [
      "Cafe24 PRO UP",
      "카페24 PRO UP"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24PROUP.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.",
    "version": "2025-11-04",
    "licenseText": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24PROUP.woff"
  },
  {
    "family": "cafe24MoyamoyaFace",
    "label": "카페24 모야모야 Face",
    "aliases": [
      "Cafe24 Moyamoya Face",
      "카페24 모야모야 Face"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Moyamoya-Face-v1.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.",
    "version": "1.0",
    "licenseText": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Moyamoya-Face-v1.0.woff"
  },
  {
    "family": "cafe24Moyamoya",
    "label": "카페24 모야모야",
    "aliases": [
      "Cafe24 Moyamoya",
      "카페24 모야모야"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Moyamoya-Regular-v1.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.",
    "version": "1.0",
    "licenseText": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Moyamoya-Regular-v1.0.woff"
  },
  {
    "family": "cafe24SupermagicBold",
    "label": "카페24 슈퍼매직 Bold",
    "aliases": [
      "Cafe24 Supermagic Bold",
      "카페24 슈퍼매직 Bold"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Supermagic-Bold-v1.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.",
    "version": "1.0",
    "licenseText": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Supermagic-Bold-v1.0.woff"
  },
  {
    "family": "cafe24SupermagicRegular",
    "label": "카페24 슈퍼매직 Regular",
    "aliases": [
      "Cafe24 Supermagic Regular",
      "카페24 슈퍼매직 Regular"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Supermagic-Regular-v1.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.",
    "version": "1.0",
    "licenseText": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Supermagic-Regular-v1.0.woff"
  },
  {
    "family": "cafe24Classictype",
    "label": "카페24 클래식타입",
    "aliases": [
      "Cafe24 ClassicType",
      "카페24 클래식타입"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Classictype-v1.1.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.",
    "version": "1.1",
    "licenseText": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Classictype-v1.1.woff"
  },
  {
    "family": "cafe24SsurroundAir",
    "label": "카페24 써라운드 에어",
    "aliases": [
      "Cafe24 Ssurround Air",
      "카페24 써라운드 에어"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24SsurroundAir-v1.1.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.",
    "version": "1.1",
    "licenseText": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24SsurroundAir-v1.1.woff"
  },
  {
    "family": "cafe24Ssurround",
    "label": "카페24 써라운드",
    "aliases": [
      "Cafe24 Ssurround",
      "카페24 써라운드"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Ssurround-v2.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright © Cafe24 Corp. All Rights Reserved.",
    "version": "2.0",
    "licenseText": "Copyright © Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Ssurround-v2.0.woff"
  },
  {
    "family": "cafe24OhsquareAir",
    "label": "카페24 아네모네 에어",
    "aliases": [
      "Cafe24 Ohsquare Air",
      "카페24 아네모네 에어"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24OhsquareAir-v2.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright (c) Cafe24 Corp. All Rights Reserved.",
    "version": "2.0",
    "licenseText": "Copyright (c) Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24OhsquareAir-v2.0.woff"
  },
  {
    "family": "cafe24Ohsquare",
    "label": "카페24 아네모네",
    "aliases": [
      "Cafe24 Ohsquare",
      "카페24 아네모네"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Ohsquare-v2.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright (c) 2020by Cafe24. All rights reserved.",
    "version": "2.0",
    "licenseText": "Copyright (c) 2020by Cafe24. All rights reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Ohsquare-v2.0.woff"
  },
  {
    "family": "cafe24Shiningstar",
    "label": "카페24 빛나는별",
    "aliases": [
      "Cafe24 Shiningstar",
      "카페24 빛나는별"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Shiningstar-v2.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright (c) 2019 by Cafe24. All rights reserved.",
    "version": "2.0",
    "licenseText": "Copyright (c) 2019 by Cafe24. All rights reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Shiningstar-v2.0.woff"
  },
  {
    "family": "cafe24Oneprettynight",
    "label": "카페24 고운밤",
    "aliases": [
      "Cafe24 Oneprettynight",
      "카페24 고운밤"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Oneprettynight-v2.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright (c) 2020 by Cafe24. All rights reserved.",
    "version": "2.0",
    "licenseText": "Copyright (c) 2020 by Cafe24. All rights reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Oneprettynight-v2.0.woff"
  },
  {
    "family": "cafe24Dangdanghae",
    "label": "카페24 당당해",
    "aliases": [
      "Cafe24 Dangdanghae",
      "카페24 당당해"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Dangdanghae-v2.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright (c) 2018 by cafe24. All rights reserved.",
    "version": "2.0",
    "licenseText": "Copyright (c) 2018 by cafe24. All rights reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Dangdanghae-v2.0.woff"
  },
  {
    "family": "Cafe24Danjunghae",
    "label": "카페24 단정해",
    "aliases": [
      "Cafe24 Danjunghae",
      "카페24 단정해"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Danjunghae-v2.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright © Cafe24 Corp. All Rights Reserved.",
    "version": "2.0",
    "licenseText": "Copyright © Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Danjunghae-v2.0.woff"
  },
  {
    "family": "cafe24Simplehae",
    "label": "카페24 심플해",
    "aliases": [
      "Cafe24 Simplehae",
      "카페24 심플해"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Simplehae-v2.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright (c) Cafe24 Corp. All Rights Reserved.",
    "version": "2.0",
    "licenseText": "Copyright (c) Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Simplehae-v2.0.woff"
  },
  {
    "family": "cafe24DongdongLight",
    "label": "카페24 동동 Light",
    "aliases": [
      "Cafe24 Dongdong Light",
      "카페24 동동 Light"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24DongdongLight.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright ⓒ Cafe24 Corp. All Rights Reserved. Cafe24 Dongdong is a trademark of Cafe24.",
    "version": "2025-11-04",
    "licenseText": "Copyright ⓒ Cafe24 Corp. All Rights Reserved. Cafe24 Dongdong is a trademark of Cafe24.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24DongdongLight.woff"
  },
  {
    "family": "cafe24Dongdong",
    "label": "카페24 동동",
    "aliases": [
      "Cafe24 Dongdong",
      "카페24 동동"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Dongdong-v2.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright (c) 2019 by Cafe24. All rights reserved.",
    "version": "2.0",
    "licenseText": "Copyright (c) 2019 by Cafe24. All rights reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Dongdong-v2.0.woff"
  },
  {
    "family": "cafe24SsukssukLight",
    "label": "카페24 쑥쑥 Light",
    "aliases": [
      "Cafe24 Ssukssuk Light",
      "카페24 쑥쑥 Light"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24SsukssukLight.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright © Cafe24 Corp. All rights Reserved.",
    "version": "2025-11-04",
    "licenseText": "Copyright © Cafe24 Corp. All rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24SsukssukLight.woff"
  },
  {
    "family": "cafe24Ssukssuk",
    "label": "카페24 쑥쑥",
    "aliases": [
      "Cafe24 Ssukssuk",
      "카페24 쑥쑥"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Ssukssuk-v2.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright © Cafe24 Corp. All rights Reserved.",
    "version": "2.0",
    "licenseText": "Copyright © Cafe24 Corp. All rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Ssukssuk-v2.0.woff"
  },
  {
    "family": "cafe24Syongsyong",
    "label": "카페24 숑숑",
    "aliases": [
      "Cafe24 Syongsyong",
      "카페24 숑숑"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Syongsyong-v2.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright © Cafe24 Corp. All rights Reserved.",
    "version": "2.0",
    "licenseText": "Copyright © Cafe24 Corp. All rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/ko/Cafe24Syongsyong-v2.0.woff"
  },
  {
    "family": "cafe24Angbanana",
    "label": "카페24 앙바나나",
    "aliases": [
      "Cafe24 Angbanana",
      "카페24 앙바나나"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/en/Cafe24Angbanana.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright \\250\\317 Cafe24 Corp. All Rights Reserved.",
    "version": "2024-04-18",
    "licenseText": "Copyright \\250\\317 Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/en/Cafe24Angbanana.woff"
  },
  {
    "family": "cafe24NyangiB",
    "label": "카페24 냥이 Black",
    "aliases": [
      "Cafe24 Nyangi Black",
      "카페24 냥이 Black"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/en/Cafe24Nyangi-B-v1.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright (c) 2023 by Cafe24. All rights reserved.",
    "version": "1.0",
    "licenseText": "Copyright (c) 2023 by Cafe24. All rights reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/en/Cafe24Nyangi-B-v1.0.woff"
  },
  {
    "family": "cafe24NyangiW",
    "label": "카페24 냥이 White",
    "aliases": [
      "Cafe24 Nyangi White",
      "카페24 냥이 White"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/en/Cafe24Nyangi-W-v1.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright (c) 2024 by Cafe24. All rights reserved.",
    "version": "1.0",
    "licenseText": "Copyright (c) 2024 by Cafe24. All rights reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/en/Cafe24Nyangi-W-v1.0.woff"
  },
  {
    "family": "cafe24MeongiB",
    "label": "카페24 먼기 Black",
    "aliases": [
      "Cafe24 Meongi Black",
      "카페24 먼기 Black"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/en/Cafe24Meongi-B-v1.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright (c) 2023 by Cafe24. All rights reserved.",
    "version": "1.0",
    "licenseText": "Copyright (c) 2023 by Cafe24. All rights reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/en/Cafe24Meongi-B-v1.0.woff"
  },
  {
    "family": "cafe24MeongiW",
    "label": "카페24 먼기 White",
    "aliases": [
      "Cafe24 Meongi White",
      "카페24 먼기 White"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/freefonts/fonts/en/Cafe24Meongi-W-v1.0.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright (c) 2023 by Cafe24. All rights reserved.",
    "version": "1.0",
    "licenseText": "Copyright (c) 2023 by Cafe24. All rights reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/freefonts/fonts/en/Cafe24Meongi-W-v1.0.woff"
  },
  {
    "family": "cafe24Behappy",
    "label": "카페24 비해피",
    "aliases": [
      "Cafe24 Behappy",
      "카페24 비해피"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/font/pkg/cafe24/cafe24Behappy-subset.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.",
    "version": "2024-04-18",
    "licenseText": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/font/pkg/cafe24/cafe24Behappy-subset.woff"
  },
  {
    "family": "cafe24Decobox",
    "label": "카페24 데코박스",
    "aliases": [
      "Cafe24 Decobox",
      "카페24 데코박스"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/font/pkg/cafe24/Cafe24Decobox-subset.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright (c) 2018 by cafe24 . All rights reserved.",
    "version": "2024-04-18",
    "licenseText": "Copyright (c) 2018 by cafe24 . All rights reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/font/pkg/cafe24/Cafe24Decobox-subset.woff"
  },
  {
    "family": "cafe24Decoline",
    "label": "카페24 데코라인",
    "aliases": [
      "Cafe24 Decoline",
      "카페24 데코라인"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/font/pkg/cafe24/Cafe24Decoline-subset.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright (c) 2018 by cafe24 . All rights reserved.",
    "version": "2024-04-18",
    "licenseText": "Copyright (c) 2018 by cafe24 . All rights reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/font/pkg/cafe24/Cafe24Decoline-subset.woff"
  },
  {
    "family": "cafe24Decomilk",
    "label": "카페24 데코밀크",
    "aliases": [
      "Cafe24 Decomilk",
      "카페24 데코밀크"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/font/pkg/cafe24/Cafe24Decomilk-subset.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.",
    "version": "2024-04-18",
    "licenseText": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/font/pkg/cafe24/Cafe24Decomilk-subset.woff"
  },
  {
    "family": "cafe24Decoschool",
    "label": "카페24 데코스쿨",
    "aliases": [
      "Cafe24 Decoschool",
      "카페24 데코스쿨"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/font/pkg/cafe24/Cafe24Decoschool-subset.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright (c) 2018 by cafe24. All rights reserved.",
    "version": "2024-04-18",
    "licenseText": "Copyright (c) 2018 by cafe24. All rights reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/font/pkg/cafe24/Cafe24Decoschool-subset.woff"
  },
  {
    "family": "cafe24Decoshadow",
    "label": "카페24 데코섀도우",
    "aliases": [
      "Cafe24 Decoshadow",
      "카페24 데코섀도우"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/font/pkg/cafe24/Cafe24Decoshadow-subset.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.",
    "version": "2024-04-18",
    "licenseText": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/font/pkg/cafe24/Cafe24Decoshadow-subset.woff"
  },
  {
    "family": "cafe24Decozoo",
    "label": "카페24 데코주",
    "aliases": [
      "Cafe24 Decozoo",
      "카페24 데코주"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/font/pkg/cafe24/Cafe24Decozoo-subset.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.",
    "version": "2024-04-18",
    "licenseText": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/font/pkg/cafe24/Cafe24Decozoo-subset.woff"
  },
  {
    "family": "cafe24Lovingu",
    "label": "카페24 러빙 유",
    "aliases": [
      "Cafe24 Loving U",
      "카페24 러빙 유"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "OFL-1.1",
    "faces": [
      {
        "url": "https://img.cafe24.com/csdstatic/font/pkg/cafe24/Cafe24Lovingu-subset.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseUrl": "https://fonts.cafe24.com/",
    "homepage": "https://fonts.cafe24.com/",
    "authors": [
      "Cafe24"
    ],
    "copyright": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.",
    "version": "2024-04-18",
    "licenseText": "Copyright ⓒ Cafe24 Corp. All Rights Reserved.\n\n카페24에서 제작한 모든 글꼴의 지적 재산권은 카페24(주)에 있습니다.\n개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 상업적인 사용이 가능합니다. 단, 글꼴 자체를 유료로 판매하는 행위를 금지합니다.\n웹디자인, 출판, 웹폰트, CI/BI 제작, 영상 제작 및 자막, 소프트웨어의 번들, 특정 프로그램의 임베드 등 사용범위 제한없이 자유롭게 이용할 수 있습니다.\n수정, 재배포가 가능하며 수정한 폰트에도 OFL(SIL Open Font License)을 적용하여야 합니다.\n카페24 폰트를 사용한 결과물은 카페24의 프로모션을 위해 활용될 수 있습니다.\n\nSIL OPEN FONT LICENSE Version 1.1 - 26 February 2007\n-----------------------------------------------------------\n\nPREAMBLE\nThe goals of the Open Font License (OFL) are to stimulate worldwide\ndevelopment of collaborative font projects, to support the font creation\nefforts of academic and linguistic communities, and to provide a free and\nopen framework in which fonts may be shared and improved in partnership\nwith others.\n\nThe OFL allows the licensed fonts to be used, studied, modified and\nredistributed freely as long as they are not sold by themselves. The\nfonts, including any derivative works, can be bundled, embedded, \nredistributed and/or sold with any software provided that any reserved\nnames are not used by derivative works. The fonts and derivatives,\nhowever, cannot be released under any other type of license. The\nrequirement for fonts to remain under this license does not apply\nto any document created using the fonts or their derivatives.\n\nDEFINITIONS\n\"Font Software\" refers to the set of files released by the Copyright\nHolder(s) under this license and clearly marked as such. This may\ninclude source files, build scripts and documentation.\n\n\"Reserved Font Name\" refers to any names specified as such after the\ncopyright statement(s).\n\n\"Original Version\" refers to the collection of Font Software components as\ndistributed by the Copyright Holder(s).\n\n\"Modified Version\" refers to any derivative made by adding to, deleting,\nor substituting -- in part or in whole -- any of the components of the\nOriginal Version, by changing formats or by porting the Font Software to a\nnew environment.\n\n\"Author\" refers to any designer, engineer, programmer, technical\nwriter or other person who contributed to the Font Software.\n\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge, to any person obtaining\na copy of the Font Software, to use, study, copy, merge, embed, modify,\nredistribute, and sell modified and unmodified copies of the Font\nSoftware, subject to the following conditions:\n\n1) Neither the Font Software nor any of its individual components,\nin Original or Modified Versions, may be sold by itself.\n\n2) Original or Modified Versions of the Font Software may be bundled,\nredistributed and/or sold with any software, provided that each copy\ncontains the above copyright notice and this license. These can be\nincluded either as stand-alone text files, human-readable headers or\nin the appropriate machine-readable metadata fields within text or\nbinary files as long as those fields can be easily viewed by the user.\n\n3) No Modified Version of the Font Software may use the Reserved Font\nName(s) unless explicit written permission is granted by the corresponding\nCopyright Holder. This restriction only applies to the primary font name as\npresented to the users.\n\n4) The name(s) of the Copyright Holder(s) or the Author(s) of the Font\nSoftware shall not be used to promote, endorse or advertise any\nModified Version, except to acknowledge the contribution(s) of the\nCopyright Holder(s) and the Author(s) or with their explicit written\npermission.\n\n5) The Font Software, modified or unmodified, in part or in whole,\nmust be distributed entirely under this license, and must not be\ndistributed under any other license. The requirement for fonts to\nremain under this license does not apply to any document created\nusing the Font Software.\n\nTERMINATION\nThis license becomes null and void if any of the above conditions are\nnot met.\n\nDISCLAIMER\nTHE FONT SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND,\nEXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO ANY WARRANTIES OF\nMERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT\nOF COPYRIGHT, PATENT, TRADEMARK, OR OTHER RIGHT. IN NO EVENT SHALL THE\nCOPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,\nINCLUDING ANY GENERAL, SPECIAL, INDIRECT, INCIDENTAL, OR CONSEQUENTIAL\nDAMAGES, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING\nFROM, OUT OF THE USE OR INABILITY TO USE THE FONT SOFTWARE OR FROM\nOTHER DEALINGS IN THE FONT SOFTWARE.",
    "licenseEvidenceUrl": "https://img.cafe24.com/csdstatic/font/pkg/cafe24/Cafe24Lovingu-subset.woff"
  },
  {
    "family": "Jalnan 2",
    "label": "여기어때 잘난체2",
    "aliases": [
      "잘난체2",
      "Jalnan2",
      "Jalnan 2 Regular"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "Jalnan License",
    "faces": [
      {
        "url": "https://framerusercontent.com/assets/tei8CJnHqrxTy9hMp7WwAK6aIA.woff2",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseText": "‘여기어때 잘난체’는 여기어때에서 배포하는 글꼴(잘난체·잘난체2·잘난체 고딕)이 모두 포함된 명칭으로, ‘여기어때 잘난체’의 지식재산권은 (주)여기어때컴퍼니가 소유하고 있습니다.\n‘여기어때 잘난체’는 개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 자유롭게 사용할 수 있습니다.\n‘여기어때 잘난체’ 글꼴 파일(otf/ttf) 자체를 유료로 판매하거나 어떠한 형태로든 임의 수정·개작하여 글꼴을 재배포 하는 것은 금지합니다.\n‘여기어때 잘난체’는 인쇄물, 광고물(온·오프라인), 상품, CI·BI 등 상업적 목적으로 사용할 수 있습니다.\n① 글꼴 그대로 사용을 권장하며, 형태 변형이 필요할 경우 글꼴의 조형성을 해치지 않는 범위 내에서 간단한 변형이 가능합니다.\n② 상품에 사용 시 여기어때 잘난체 라이선스 혹은 출처 표기를 권장합니다.\n③ CI·BI에 사용 시 상표권 등록은 불가합니다.\n‘여기어때 잘난체’는 본 저작권 안내와 라이선스 전문을 포함해서 다른 소프트웨어에 임베딩 또는 번들하여 판매하거나 재배포가 가능합니다. 임베딩 된 폰트는 별도 판매 불가합니다.\n‘여기어때 잘난체’ 사용 시 라이선스 전문을 포함하기 어려울 경우 ‘여기어때 잘난체’의 출처 표기를 권장합니다. 예) 이 상품에는 (주)여기어때컴퍼니가 제공한 여기어때 잘난체·잘난체 고딕이 적용되어 있습니다.\n‘여기어때 잘난체’를 사용한 인쇄물, 광고물(온·오프라인), 상품, CI·BI 등의 이미지는 여기어때의 마케팅을 위해 활용될 수 있습니다. 이를 원치않는 사용자는 언제든지 당사에 요청하실 수 있습니다.",
    "licenseUrl": "https://drive.google.com/file/d/1TFGKe-QLRnx2IHpJO3ZH3pVUYjQwMx7L/view",
    "homepage": "https://gccompany.co.kr/font",
    "authors": [
      "GC Company"
    ],
    "copyright": "Copyright 2023 GC Company corp. All rights reserved.",
    "version": "2023-09-27"
  },
  {
    "family": "Jalnan Gothic",
    "label": "여기어때 잘난체 고딕",
    "aliases": [
      "잘난 고딕",
      "JalnanGothic",
      "Jalnan Gothic Regular"
    ],
    "source": "cdn",
    "category": "display",
    "scripts": [
      "korean",
      "latin"
    ],
    "weights": [
      400
    ],
    "styles": [
      "normal"
    ],
    "variants": [
      "400"
    ],
    "license": "Jalnan License",
    "faces": [
      {
        "url": "https://framerusercontent.com/assets/dqOvmrr5tgZif3fg3n9XPsEbQ.otf",
        "weight": 400,
        "style": "normal"
      }
    ],
    "licenseText": "‘여기어때 잘난체’는 여기어때에서 배포하는 글꼴(잘난체·잘난체2·잘난체 고딕)이 모두 포함된 명칭으로, ‘여기어때 잘난체’의 지식재산권은 (주)여기어때컴퍼니가 소유하고 있습니다.\n‘여기어때 잘난체’는 개인 및 기업 사용자를 포함한 모든 사용자에게 무료로 제공되며 자유롭게 사용할 수 있습니다.\n‘여기어때 잘난체’ 글꼴 파일(otf/ttf) 자체를 유료로 판매하거나 어떠한 형태로든 임의 수정·개작하여 글꼴을 재배포 하는 것은 금지합니다.\n‘여기어때 잘난체’는 인쇄물, 광고물(온·오프라인), 상품, CI·BI 등 상업적 목적으로 사용할 수 있습니다.\n① 글꼴 그대로 사용을 권장하며, 형태 변형이 필요할 경우 글꼴의 조형성을 해치지 않는 범위 내에서 간단한 변형이 가능합니다.\n② 상품에 사용 시 여기어때 잘난체 라이선스 혹은 출처 표기를 권장합니다.\n③ CI·BI에 사용 시 상표권 등록은 불가합니다.\n‘여기어때 잘난체’는 본 저작권 안내와 라이선스 전문을 포함해서 다른 소프트웨어에 임베딩 또는 번들하여 판매하거나 재배포가 가능합니다. 임베딩 된 폰트는 별도 판매 불가합니다.\n‘여기어때 잘난체’ 사용 시 라이선스 전문을 포함하기 어려울 경우 ‘여기어때 잘난체’의 출처 표기를 권장합니다. 예) 이 상품에는 (주)여기어때컴퍼니가 제공한 여기어때 잘난체·잘난체 고딕이 적용되어 있습니다.\n‘여기어때 잘난체’를 사용한 인쇄물, 광고물(온·오프라인), 상품, CI·BI 등의 이미지는 여기어때의 마케팅을 위해 활용될 수 있습니다. 이를 원치않는 사용자는 언제든지 당사에 요청하실 수 있습니다.",
    "licenseUrl": "https://drive.google.com/file/d/1TFGKe-QLRnx2IHpJO3ZH3pVUYjQwMx7L/view",
    "homepage": "https://gccompany.co.kr/font",
    "authors": [
      "GC Company"
    ],
    "copyright": "Copyright 2023 GC Company corp. All rights reserved.",
    "version": "2023-09-27"
  }
];
if (new Set(rows.map(row=>row[0].toLowerCase()).concat(curated.map(font=>font.family.toLowerCase()))).size!==rows.length+curated.length) throw new Error('Duplicate font family');
const meta = {version:1,updatedAt,googleFamilies:metadata.familyMetadataList.length,googleIncluded:rows.length,curatedFamilies:curated.length,total:rows.length+curated.length,excludedFamilies:excluded,googleCommit:tree.sha,googleMetadataUrl:GOOGLE_METADATA,googleMetadataSha256:sha256(rawMetadata),googleTreeSha256:sha256(rawTree),fallbackLicenses,curatedSources:[{"name":"Cafe24 Korean","url":"https://img.cafe24.com/csdstatic/freefonts/data/fonts_ko.json","sha256":"60573b3ea5d717484a3821dbde74eac3b52e0f4e1dfba8b50472b84bb767df37","verifiedAt":"2026-10-05"},{"name":"Cafe24 English","url":"https://img.cafe24.com/csdstatic/freefonts/data/fonts_en.json","sha256":"f4eb54242d826370fbed6736a8ef6ee170eacba0d52ab9180dc70f618e484a26","verifiedAt":"2026-10-05"},{"name":"Cafe24 Korean CSS","url":"https://img.cafe24.com/csdstatic/freefonts/data/1.0/fonts_ko.css","sha256":"9538aab79119e362ed36dc3b6a480a3e85271cd6286b93ecae6688fd42d7b55d","verifiedAt":"2026-10-05"},{"name":"Cafe24 English CSS","url":"https://img.cafe24.com/csdstatic/freefonts/data/1.0/fonts_en.css","sha256":"82a2bbc50c1f08ae47e7af4821db8108e7c6a0a34cfa22719e68d3ddb1e27d22","verifiedAt":"2026-10-05"},{"name":"NanumSquareNeo CSS","url":"https://hangeul.pstatic.net/hangeul_static/css/nanum-square-neo.css","sha256":"7fee8e65b6b6a458f7b5024d048b667efc8e5ce75a810f8b58f7c041585ac1ba","verifiedAt":"2026-10-05"}],runtimeOrigins:["https://fonts.googleapis.com","https://fonts.gstatic.com","https://cdn.jsdelivr.net","https://hangeul.pstatic.net","https://script.gmarket.com","https://img.cafe24.com","https://framerusercontent.com"]};
const source = '// Generated by tools/update-font-catalog.mjs. No font binaries or document text.\nexport const WEB_FONT_CATALOG_META = '+JSON.stringify(meta,null,2)+';\nconst WEB_FONT_GOOGLE_ROWS = [\n'+rows.map(row=>'  '+JSON.stringify(row)).join(',\n')+'\n];\nconst WEB_FONT_CURATED = '+JSON.stringify(curated,null,2)+';\n'+`export const WEB_FONT_CATALOG = WEB_FONT_GOOGLE_ROWS.map(([family,category,scripts,variants,path,aliases,authors,fallback]) => ({
  family,label:family,aliases,source:'google',category,scripts,variants,
  weights:[...new Set(variants.map(v=>parseInt(v,10)))].sort((a,b)=>a-b),styles:[...new Set(variants.map(v=>v.endsWith('i')?'italic':'normal'))],
  license:fallback?.license || (path.startsWith('apache/')?'Apache-2.0':path.startsWith('ufl/')?'UFL-1.0':'OFL-1.1'),
  licenseUrl:path?'https://github.com/google/fonts/blob/'+WEB_FONT_CATALOG_META.googleCommit+'/'+path:fallback.url,
  rawLicenseUrl:path?'https://raw.githubusercontent.com/google/fonts/'+WEB_FONT_CATALOG_META.googleCommit+'/'+path:fallback.url,
  homepage:'https://fonts.google.com/specimen/'+encodeURIComponent(family).replace(/%20/g,'+'),authors,
  ...(fallback ? {copyright:fallback.copyright,licenseText:fallback.text} : {})
})).concat(WEB_FONT_CURATED);
`;
// Maintenance opt-in: verify curated URLs without downloading font binaries.
if (args.includes('--verify-network')) {
  if (offline) throw new Error('--offline and --verify-network cannot be combined');
  const checks = [];
  const tasks = curated.flatMap(font => font.cssUrl
    ? [{ family:font.family, url:font.cssUrl, css:true }]
    : font.faces.map(face => ({ family:font.family, url:face.url, css:false })));
  let next = 0;
  await Promise.all(Array.from({ length:4 }, async () => {
    while (next < tasks.length) {
      const task = tasks[next++];
      const allowed = meta.runtimeOrigins.includes(new URL(task.url).origin);
      if (!allowed) throw new Error('Unregistered font origin: '+task.url);
      let response = await fetch(task.url, { method:task.css?'GET':'HEAD', headers:{ Origin:'https://wixel-3.wizx.workers.dev' }, signal:AbortSignal.timeout(30000) });
      if (!task.css && response.ok && !response.headers.get('access-control-allow-origin')) {
        // Some original distributors add CORS on GET only; inspect a one-byte range.
        response = await fetch(task.url,{ headers:{ Origin:'https://wixel-3.wizx.workers.dev',Range:'bytes=0-0' }, signal:AbortSignal.timeout(30000) });
      }
      const check = { ...task, status:response.status, cors:response.headers.get('access-control-allow-origin') };
      if (task.css) {
        const css = await response.text();
        check.sha256 = sha256(css);
        check.hasFontFaces = css.includes('@font-face') && css.includes(task.family);
      } else await response.body?.cancel();
      checks.push(check);
    }
  }));
  await fs.writeFile(path.join(cache,'catalog-network-report.json'),JSON.stringify(checks,null,2));
  const failures = checks.filter(check => ![200,206].includes(check.status) || check.cors!=='*' || (check.css && !check.hasFontFaces));
  if (failures.length) throw new Error('Font network verification failed: '+JSON.stringify(failures));
  console.log(JSON.stringify({ verifiedCuratedRequests:checks.length, fontBinariesSaved:0 }));
}


await fs.writeFile(output,source);
await fs.writeFile(path.join(cache,'catalog-build-report.json'),JSON.stringify({...meta,output,bytes:Buffer.byteLength(source),sha256:sha256(source)},null,2));
console.log(JSON.stringify({families:meta.total,google:rows.length,curated:curated.length,output,bytes:Buffer.byteLength(source)}));
