import { parseXml, descendants, child, esc } from './xml.js';

export const THEME_FONTS = [
  { name: 'Office', major: '맑은 고딕', minor: '맑은 고딕', majorLatin: 'Aptos Display', minorLatin: 'Aptos' },
  { name: 'Office 2013–2022', major: '맑은 고딕', minor: '맑은 고딕', majorLatin: 'Calibri Light', minorLatin: 'Calibri' },
  { name: '맑은 고딕', major: '맑은 고딕', minor: '맑은 고딕' },
  { name: '나눔고딕', major: '나눔고딕', minor: '나눔고딕' },
  { name: 'Arial', major: 'Arial', minor: 'Arial' },
  { name: 'Georgia · Arial', major: 'Georgia', minor: 'Arial' },
  { name: 'Times New Roman · Arial', major: 'Times New Roman', minor: 'Arial' },
];

export const THEME_EFFECTS = [
  { id: 'none', name: '효과 없음' },
  { id: 'subtle', name: '은은한 그림자', shadow: { dx: 1.5, dy: 2, blur: 3, color: '#000000', opacity: 0.18 } },
  { id: 'moderate', name: '보통 그림자', shadow: { dx: 3, dy: 4, blur: 6, color: '#000000', opacity: 0.26 } },
  { id: 'strong', name: '강한 그림자', shadow: { dx: 5, dy: 6, blur: 9, color: '#000000', opacity: 0.34 } },
];

/** 원본 테마에서 한국어(한글 스크립트)와 라틴 글꼴을 읽는다. 기타 스크립트는 XML로 보존. */
export function readThemeOptions(xml) {
  if (!xml) return {};
  const root = parseXml(xml), scheme = descendants(root, 'fontScheme')[0], format = descendants(root, 'fmtScheme')[0];
  const result = {};
  if (scheme) {
    const font = key => {
      const node = child(scheme, key), latin = child(node, 'latin')?.attrs.typeface || '';
      const hang = node?.children.find(x => x.name === 'font' && x.attrs.script === 'Hang')?.attrs.typeface;
      return { local: hang || child(node, 'ea')?.attrs.typeface || latin, latin };
    };
    const major = font('majorFont'), minor = font('minorFont');
    result.themeFonts = { name: scheme.attrs.name || '사용자 지정', major: major.local || '맑은 고딕', minor: minor.local || '맑은 고딕', majorLatin: major.latin, minorLatin: minor.latin };
  }
  // 알려진 WIXEL 선택만 재구성한다. 다른 효과 스킴은 원본 XML을 그대로 보존한다.
  const effect = THEME_EFFECTS.find(x => format?.attrs.name === `WIXEL ${x.name}`);
  if (effect) result.themeEffects = structuredClone(effect);
  return result;
}

function replaceElement(xml, name, replacement) {
  const re = new RegExp(`<([A-Za-z_][\\w.-]*:)?${name}\\b[^>]*>[\\s\\S]*?<\\/\\1${name}>`);
  return xml.replace(re, replacement);
}

/** 색 이외의 명시적으로 선택한 글꼴/효과만 변경하며 기존 테마의 나머지 요소는 보존한다. */
export function applyThemeOptionsXml(xml, wb) {
  let result = xml;
  const fonts = wb.themeFonts;
  if (fonts?.major && fonts?.minor) {
    const read = readThemeOptions(result).themeFonts;
    // 가져오기만 한 파일은 스킴이 동일하므로 원본 글꼴 메타데이터를 손대지 않는다.
    if (!read || ['name', 'major', 'minor', 'majorLatin', 'minorLatin'].some(k => String(read[k] ?? '') !== String(fonts[k] ?? ''))) {
      result = replaceElement(result, 'fontScheme', original => {
        const prefix = /^<([A-Za-z_][\w.-]*:)?/.exec(original)?.[1] || '';
        let next = original.replace(/\bname="[^"]*"/, `name="${esc(fonts.name || '사용자 지정')}"`);
        for (const [tag, key] of [['majorFont', 'major'], ['minorFont', 'minor']]) next = replaceElement(next, tag, block => {
          let value = block.replace(new RegExp(`<${prefix}latin\\b[^>]*/>`), `<${prefix}latin typeface="${esc(fonts[`${key}Latin`] || fonts[key])}"/>`);
          const hang = new RegExp(`<${prefix}font\\b(?=[^>]*\\bscript="Hang")[^>]*/>`);
          const entry = `<${prefix}font script="Hang" typeface="${esc(fonts[key])}"/>`;
          value = hang.test(value) ? value.replace(hang, entry) : value.replace(`</${prefix}${tag}>`, `${entry}</${prefix}${tag}>`);
          return value;
        });
        return next;
      });
    }
  }
  const effects = THEME_EFFECTS.find(x => x.id === wb.themeEffects?.id);
  if (effects) result = replaceElement(result, 'fmtScheme', original => {
    const prefix = /^<([A-Za-z_][\w.-]*:)?/.exec(original)?.[1] || '';
    const shadow = effects.shadow;
    const effect = shadow ? `<${prefix}outerShdw blurRad="${Math.round(shadow.blur * 9525)}" dist="${Math.round(Math.hypot(shadow.dx, shadow.dy) * 9525)}" dir="${Math.round((Math.atan2(shadow.dy, shadow.dx) * 180 / Math.PI + 360) % 360 * 60000)}" algn="ctr" rotWithShape="0"><${prefix}srgbClr val="000000"><${prefix}alpha val="${Math.round(shadow.opacity * 100000)}"/></${prefix}srgbClr></${prefix}outerShdw>` : '';
    const style = `<${prefix}effectStyle><${prefix}effectLst>${effect}</${prefix}effectLst></${prefix}effectStyle>`;
    return replaceElement(original.replace(/\bname="[^"]*"/, `name="${esc(`WIXEL ${effects.name}`)}"`), 'effectStyleLst', `<${prefix}effectStyleLst>${style.repeat(3)}</${prefix}effectStyleLst>`);
  });
  return result;
}
