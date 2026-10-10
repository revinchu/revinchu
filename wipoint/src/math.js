// 수식 (DOM 없음): LaTeX 일부 → 수식 트리 → MathML (화면) · OMML (PowerPoint 수식 a14:m), OMML → LaTeX (읽기)
//   지원: 분수 \frac, 근호 \sqrt[n]{}, 위·아래 첨자 ^ _, 큰 연산자 \sum \prod \int (첨자 = 범위), 괄호 \left( \right),
//         행렬 \begin{matrix|pmatrix|bmatrix}, 글자 \text{}, 강조 \hat \bar \vec \dot \tilde, 그리스 문자 · 기호
import { esc } from './xml.js';

export const SYMBOLS = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ε', zeta: 'ζ', eta: 'η', theta: 'θ', iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', rho: 'ρ', sigma: 'σ', tau: 'τ', upsilon: 'υ', phi: 'φ', varphi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω',
  Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
  times: '×', cdot: '⋅', div: '÷', pm: '±', mp: '∓', le: '≤', leq: '≤', ge: '≥', geq: '≥', ne: '≠', neq: '≠', approx: '≈', equiv: '≡', sim: '∼', propto: '∝',
  infty: '∞', partial: '∂', nabla: '∇', forall: '∀', exists: '∃', in: '∈', notin: '∉', subset: '⊂', supset: '⊃', subseteq: '⊆', cup: '∪', cap: '∩', emptyset: '∅',
  to: '→', rightarrow: '→', leftarrow: '←', Rightarrow: '⇒', Leftarrow: '⇐', leftrightarrow: '↔', Leftrightarrow: '⇔', mapsto: '↦',
  cdots: '⋯', ldots: '…', dots: '…', vdots: '⋮', ddots: '⋱', degree: '°', circ: '∘', angle: '∠', perp: '⊥', parallel: '∥', therefore: '∴', because: '∵',
  langle: '⟨', rangle: '⟩', lfloor: '⌊', rfloor: '⌋', lceil: '⌈', rceil: '⌉', neg: '¬', land: '∧', lor: '∨', prime: '′', hbar: 'ℏ', ell: 'ℓ',
};
const NARY = { sum: '∑', prod: '∏', coprod: '∐', int: '∫', iint: '∬', iiint: '∭', oint: '∮', bigcup: '⋃', bigcap: '⋂' };
const FUNCS = ['sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'log', 'ln', 'exp', 'lim', 'max', 'min', 'sup', 'inf', 'det', 'arcsin', 'arccos', 'arctan', 'sinh', 'cosh', 'tanh'];
const ACCENTS = { hat: '̂', bar: '̅', vec: '⃗', dot: '̇', ddot: '̈', tilde: '̃', overline: '̅' };
const OPS = '+-=<>±×÷⋅/!,;:|∣≤≥≠≈→←';

/** LaTeX → 수식 트리 */
export function parseLatex(src) {
  let i = 0;
  const s = String(src ?? '');
  const peek = () => s[i];
  const skipWs = () => { while (i < s.length && /\s/.test(s[i])) i++; };
  const readName = () => { let n = ''; while (i < s.length && /[a-zA-Z]/.test(s[i])) n += s[i++]; return n || s[i++] || ''; };
  const group = () => {
    skipWs();
    if (peek() === '{') { i++; const r = row('}'); i++; return r; }
    return atom() ?? { t: 'row', items: [] };
  };
  const rawGroup = () => { skipWs(); if (peek() !== '{') return ''; i++; let d = 1; let out = ''; while (i < s.length) { const c = s[i++]; if (c === '{') d++; else if (c === '}' && !--d) break; out += c; } return out; };
  const delim = () => {
    skipWs();
    if (s[i] === '\\') { i++; const n = readName(); return n === '{' ? '{' : n === '}' ? '}' : n === '|' ? '‖' : SYMBOLS[n] ?? (n === '.' ? '' : n); }
    return s[i++] === '.' ? '' : s[i - 1];
  };
  function atom() {
    skipWs();
    const c = peek();
    if (c === undefined) return null;
    if (c === '{') { i++; const r = row('}'); i++; return r; }
    if (c === '\\') {
      i++;
      const n = readName();
      if (n === 'frac' || n === 'dfrac' || n === 'tfrac') return { t: 'frac', n: group(), d: group() };
      if (n === 'sqrt') {
        skipWs();
        let idx = null;
        if (peek() === '[') { i++; idx = row(']'); i++; }
        return { t: 'sqrt', x: group(), idx };
      }
      if (n === 'left') { const open = delim(); const x = row('\\right'); i += 6; const close = delim(); return { t: 'fenced', open, close, x }; }
      if (n === 'text' || n === 'mathrm' || n === 'textrm' || n === 'operatorname') return { t: 'text', v: rawGroup() };
      if (n === 'begin') {
        const env = rawGroup();
        const body = [];
        let cur = [];
        let cell = [];
        const flushCell = () => { cur.push({ t: 'row', items: cell }); cell = []; };
        while (i < s.length) {
          skipWs();
          if (s.startsWith('\\end', i)) { i += 4; rawGroup(); break; }
          if (s.startsWith('\\\\', i)) { i += 2; flushCell(); body.push(cur); cur = []; continue; }
          if (s[i] === '&') { i++; flushCell(); continue; }
          const a = atom();
          if (a) cell.push(...script(a));
        }
        if (cell.length || cur.length) { flushCell(); body.push(cur); }
        const m = { t: 'matrix', rows: body };
        const fence = { pmatrix: ['(', ')'], bmatrix: ['[', ']'], vmatrix: ['|', '|'], Bmatrix: ['{', '}'], cases: ['{', ''] }[env];
        return fence ? { t: 'fenced', open: fence[0], close: fence[1], x: m } : m;
      }
      if (ACCENTS[n]) return { t: 'acc', ch: ACCENTS[n], x: group() };
      if (NARY[n]) return { t: 'nary', op: NARY[n] };
      if (FUNCS.includes(n)) return { t: 'func', v: n };
      if (n === ',' || n === ';' || n === ' ' || n === 'quad' || n === 'qquad' || n === '!') return { t: 'sp', v: n };
      if (n === '{' || n === '}' || n === '%' || n === '$' || n === '#' || n === '&' || n === '_') return { t: 'op', v: n };
      if (SYMBOLS[n]) return { t: /[α-ωΑ-Ωℏℓ]/.test(SYMBOLS[n]) ? 'id' : 'op', v: SYMBOLS[n] };
      return { t: 'id', v: n };
    }
    i++;
    if (/[0-9.]/.test(c)) { let v = c; while (i < s.length && /[0-9.]/.test(s[i])) v += s[i++]; return { t: 'num', v }; }
    if (OPS.includes(c) || '()[]'.includes(c)) return { t: 'op', v: c === '-' ? '−' : c === '*' ? '⋅' : c };
    if (c === '\'') return { t: 'op', v: '′' };
    return { t: 'id', v: c };
  }
  /** 첨자 붙이기: 앞 원소에 ^ _ */
  function script(base) {
    let sub = null; let sup = null;
    for (;;) {
      skipWs();
      if (peek() === '^') { i++; sup = group(); } else if (peek() === '_') { i++; sub = group(); } else break;
    }
    if (base.t === 'nary') {
      skipWs();
      // 큰 연산자 뒤의 피연산자: 다음 원소 하나 (괄호 묶음이면 그 묶음)
      const body = atom();
      const rest = body ? script(body) : [];
      return [{ ...base, sub, sup, body: { t: 'row', items: rest } }];
    }
    if (sub && sup) return [{ t: 'subsup', base, sub, sup }];
    if (sup) return [{ t: 'sup', base, sup }];
    if (sub) return [{ t: 'sub', base, sub }];
    return [base];
  }
  function row(end) {
    const items = [];
    while (i < s.length) {
      skipWs();
      if (end && s.startsWith(end, i)) break;
      if (s[i] === '}' && end !== '}') { i++; continue; }
      if (s[i] === '&' || s.startsWith('\\\\', i)) break;
      if ((s[i] === '^' || s[i] === '_')) { items.push(...script({ t: 'row', items: [] })); continue; }
      const a = atom();
      if (!a) break;
      items.push(...script(a));
    }
    return { t: 'row', items };
  }
  return row(null);
}

// ───────────── MathML (화면 표시) ─────────────
export function toMathML(node, { display = true } = {}) {
  return `<math xmlns="http://www.w3.org/1998/Math/MathML"${display ? ' display="block"' : ''}>${mml(node)}</math>`;
}
function mrow(n) { return n.t === 'row' && n.items.length !== 1 ? `<mrow>${n.items.map(mml).join('')}</mrow>` : mml(n.t === 'row' ? n.items[0] : n); }
function mml(n) {
  if (!n) return '<mrow/>';
  switch (n.t) {
    case 'row': return n.items.length === 1 ? mml(n.items[0]) : `<mrow>${n.items.map(mml).join('')}</mrow>`;
    case 'num': return `<mn>${esc(n.v)}</mn>`;
    case 'id': return `<mi>${esc(n.v)}</mi>`;
    case 'func': return `<mi mathvariant="normal">${esc(n.v)}</mi><mo>&#x2061;</mo>`;
    case 'op': return `<mo>${esc(n.v)}</mo>`;
    case 'text': return `<mtext>${esc(n.v)}</mtext>`;
    case 'sp': return `<mspace width="${{ ',': '0.17em', ';': '0.28em', ' ': '0.28em', quad: '1em', qquad: '2em', '!': '0' }[n.v] ?? '0.2em'}"/>`;
    case 'frac': return `<mfrac>${mrow(n.n)}${mrow(n.d)}</mfrac>`;
    case 'sqrt': return n.idx ? `<mroot>${mrow(n.x)}${mrow(n.idx)}</mroot>` : `<msqrt>${mml(n.x)}</msqrt>`;
    case 'sup': return `<msup>${mrow(n.base)}${mrow(n.sup)}</msup>`;
    case 'sub': return `<msub>${mrow(n.base)}${mrow(n.sub)}</msub>`;
    case 'subsup': return `<msubsup>${mrow(n.base)}${mrow(n.sub)}${mrow(n.sup)}</msubsup>`;
    case 'nary': {
      const op = `<mo largeop="true" movablelimits="false">${n.op}</mo>`;
      const lim = n.op === '∫' || n.op === '∬' || n.op === '∭' || n.op === '∮';
      const tag = lim ? (n.sub && n.sup ? 'msubsup' : n.sub ? 'msub' : n.sup ? 'msup' : '') : (n.sub && n.sup ? 'munderover' : n.sub ? 'munder' : n.sup ? 'mover' : '');
      const head = tag ? `<${tag}>${op}${n.sub ? mrow(n.sub) : ''}${n.sup ? mrow(n.sup) : ''}</${tag}>` : op;
      return `<mrow>${head}${n.body ? mml(n.body) : ''}</mrow>`;
    }
    case 'fenced': return `<mrow>${n.open ? `<mo fence="true">${esc(n.open)}</mo>` : ''}${mml(n.x)}${n.close ? `<mo fence="true">${esc(n.close)}</mo>` : ''}</mrow>`;
    case 'matrix': return `<mtable>${n.rows.map((r) => `<mtr>${r.map((c) => `<mtd>${mml(c)}</mtd>`).join('')}</mtr>`).join('')}</mtable>`;
    case 'acc': return `<mover accent="true">${mrow(n.x)}<mo>${n.ch === '̅' ? '‾' : n.ch === '⃗' ? '→' : n.ch === '̂' ? '^' : n.ch === '̃' ? '~' : n.ch === '̈' ? '¨' : '˙'}</mo></mover>`;
    default: return '';
  }
}

// ───────────── OMML (PowerPoint 수식) ─────────────
/** 수식 트리 → <m:oMathPara> (a14:m 안에 넣음). rPr = 글자 서식 (a:rPr 속 내용) */
export function toOMML(node, { size = 28, color = null } = {}) {
  const fill = color ? `<a:solidFill><a:srgbClr val="${color.replace('#', '').toUpperCase()}"/></a:solidFill>` : '';
  const rpr = (italic) => `<a:rPr lang="en-US" sz="${Math.round(size * 100)}" i="${italic ? 1 : 0}">${fill}<a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr>`;
  const run = (t, { italic = false, plain = false } = {}) => `<m:r>${plain ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : ''}${rpr(italic)}<m:t>${esc(t)}</m:t></m:r>`;
  const e = (n) => `<m:e>${om(n)}</m:e>`;
  function om(n) {
    if (!n) return '';
    switch (n.t) {
      case 'row': return n.items.map(om).join('');
      case 'num': return run(n.v);
      case 'id': return run(n.v, { italic: true });
      case 'op': case 'sp': return run(n.t === 'sp' ? ' ' : n.v);
      case 'func': return `<m:func><m:fName>${run(n.v, { plain: true })}</m:fName><m:e></m:e></m:func>`;
      case 'text': return run(n.v, { plain: true });
      case 'frac': return `<m:f><m:num>${om(n.n)}</m:num><m:den>${om(n.d)}</m:den></m:f>`;
      case 'sqrt': return n.idx ? `<m:rad><m:deg>${om(n.idx)}</m:deg>${e(n.x)}</m:rad>` : `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>${e(n.x)}</m:rad>`;
      case 'sup': return `<m:sSup>${e(n.base)}<m:sup>${om(n.sup)}</m:sup></m:sSup>`;
      case 'sub': return `<m:sSub>${e(n.base)}<m:sub>${om(n.sub)}</m:sub></m:sSub>`;
      case 'subsup': return `<m:sSubSup>${e(n.base)}<m:sub>${om(n.sub)}</m:sub><m:sup>${om(n.sup)}</m:sup></m:sSubSup>`;
      case 'nary': {
        const integral = n.op === '∫' || n.op === '∬' || n.op === '∭' || n.op === '∮';
        return `<m:nary><m:naryPr><m:chr m:val="${n.op}"/><m:limLoc m:val="${integral ? 'subSup' : 'undOvr'}"/>${n.sub ? '' : '<m:subHide m:val="1"/>'}${n.sup ? '' : '<m:supHide m:val="1"/>'}</m:naryPr><m:sub>${om(n.sub)}</m:sub><m:sup>${om(n.sup)}</m:sup>${e(n.body)}</m:nary>`;
      }
      case 'fenced': return `<m:d><m:dPr><m:begChr m:val="${esc(n.open)}"/><m:endChr m:val="${esc(n.close)}"/></m:dPr>${e(n.x)}</m:d>`;
      case 'matrix': return `<m:m>${n.rows.map((r) => `<m:mr>${r.map((c) => e(c)).join('')}</m:mr>`).join('')}</m:m>`;
      case 'acc': return `<m:acc><m:accPr><m:chr m:val="${n.ch}"/></m:accPr>${e(n.x)}</m:acc>`;
      default: return '';
    }
  }
  return `<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath>${om(node)}</m:oMath></m:oMathPara>`;
}

const REV = Object.fromEntries(Object.entries(SYMBOLS).reverse().map(([k, v]) => [v, k]));
const REV_NARY = Object.fromEntries(Object.entries(NARY).map(([k, v]) => [v, k]));
const REV_ACC = Object.fromEntries(Object.entries(ACCENTS).reverse().map(([k, v]) => [v, k]));
const val = (el, name) => { const c = el?.children.find((x) => x.name === name); return c ? c.attrs['m:val'] ?? c.attrs.val : undefined; };
const kid = (el, name) => el?.children.find((x) => x.name === name) ?? null;

/** OMML 요소 트리 (xml.js parseXml) → LaTeX */
export function ommlToLatex(el) {
  const g = (x) => { const t = conv(x); return t.length === 1 && /[A-Za-z0-9]/.test(t) ? t : `{${t}}`; };
  const b = (x) => `{${conv(x)}}`;
  const textOf = (r) => (r.children.find((x) => x.name === 't')?.text ?? '');
  function convRun(r) {
    const t = textOf(r);
    const plain = val(kid(r, 'rPr'), 'sty') === 'p';
    if (plain && /^[A-Za-z]{2,}$/.test(t) && FUNCS.includes(t)) return `\\${t} `;
    if (plain && /[A-Za-z가-힣 ]/.test(t) && t.length > 1) return `\\text{${t}}`;
    let out = '';
    for (const ch of t) out += ch === '−' ? '-' : ch === '⋅' ? '\\cdot ' : REV[ch] ? `\\${REV[ch]} ` : REV_NARY[ch] ? `\\${REV_NARY[ch]} ` : ch === '{' || ch === '}' ? `\\${ch}` : ch;
    return out;
  }
  function conv(x) {
    if (!x) return '';
    let out = '';
    for (const c of x.children ?? []) {
      switch (c.name) {
        case 'r': out += convRun(c); break;
        case 'f': out += `\\frac${b(kid(c, 'num'))}${b(kid(c, 'den'))}`; break;
        case 'rad': { const deg = conv(kid(c, 'deg')); out += `\\sqrt${deg ? `[${deg}]` : ''}{${conv(kid(c, 'e'))}}`; break; }
        case 'sSup': out += `${g(kid(c, 'e'))}^${g(kid(c, 'sup'))}`; break;
        case 'sSub': out += `${g(kid(c, 'e'))}_${g(kid(c, 'sub'))}`; break;
        case 'sSubSup': out += `${g(kid(c, 'e'))}_${g(kid(c, 'sub'))}^${g(kid(c, 'sup'))}`; break;
        case 'nary': {
          const pr = kid(c, 'naryPr');
          const ch = val(pr, 'chr') ?? '∫';
          const sub = conv(kid(c, 'sub')); const sup = conv(kid(c, 'sup'));
          out += `\\${REV_NARY[ch] ?? 'int'}${sub ? `_{${sub}}` : ''}${sup ? `^{${sup}}` : ''} ${g(kid(c, 'e'))}`;
          break;
        }
        case 'd': {
          const pr = kid(c, 'dPr');
          const o = val(pr, 'begChr') ?? '('; const cl = val(pr, 'endChr') ?? ')';
          const lx = (d) => (d === '' ? '.' : d === '{' || d === '}' ? `\\${d}` : REV[d] ? `\\${REV[d]}` : d);
          out += `\\left${lx(o)} ${kids(c, 'e').map(conv).join(',')} \\right${lx(cl)}`;
          break;
        }
        case 'm': out += `\\begin{matrix}${kids(c, 'mr').map((r) => kids(r, 'e').map(conv).join(' & ')).join(' \\\\ ')}\\end{matrix}`; break;
        case 'acc': out += `\\${REV_ACC[val(kid(c, 'accPr'), 'chr') ?? '̂'] ?? 'hat'}${b(kid(c, 'e'))}`; break;
        case 'bar': out += `\\overline${b(kid(c, 'e'))}`; break;
        case 'func': out += `${conv(kid(c, 'fName'))}${conv(kid(c, 'e'))}`; break;
        case 'limLow': out += `${conv(kid(c, 'e'))}_${g(kid(c, 'lim'))}`; break;
        case 'limUpp': out += `${conv(kid(c, 'e'))}^${g(kid(c, 'lim'))}`; break;
        case 'groupChr': case 'box': case 'borderBox': case 'sPre': case 'eqArr': case 'phant': out += kids(c, 'e').map(conv).join(c.name === 'eqArr' ? ' \\\\ ' : ''); break;
        case 'oMath': case 'oMathPara': case 'e': case 'num': case 'den': out += conv(c); break;
        default: break;
      }
    }
    return out.replace(/\s+$/, '');
  }
  return conv(el).trim();
}
const kids = (el, name) => (el ? el.children.filter((c) => c.name === name) : []);
