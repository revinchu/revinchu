import { normalizedPathBounds } from './shape-path.js';
// Normalized path editing. Workbook commits and pointer lifetimes belong to app.js.
const endOf = (c) => c && c[0] !== 'Z' ? c.slice(-2) : null;
const mix = (a, b, t) => [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t];
const distance = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);
const samePoint = (a, b) => a && b && distance(a, b) < 1e-9;
function pathSections(commands) {
  const sections = []; let start = -1;
  const finish = (last, end, closed) => {
    if (start < 0) return;
    sections.push({ start, last, end, closed, seam: closed && last > start && samePoint(endOf(commands[start]), endOf(commands[last])) });
    start = -1;
  };
  for (let i = 0; i < commands.length; i++) {
    if (commands[i][0] === 'M') { finish(i - 1, i - 1, false); start = i; }
    else if (commands[i][0] === 'Z') finish(i - 1, i, true);
  }
  finish(commands.length - 1, commands.length - 1, false);
  return sections;
}
const incomingSlot = c => c?.[0] === 'C' ? 3 : c?.[0] === 'Q' ? 1 : -1;
const outgoingSlot = c => c?.[0] === 'C' || c?.[0] === 'Q' ? 1 : -1;
export function editableShapePath(shape) {
  if (shape.path?.paths?.length) return structuredClone(shape.path);
  if (shape.customGeometry) return null;
  const paths = {
    line: [['M', 0, 0], ['L', 1, 1]],
    bentConnector3: [['M', 0, 0], ['L', 0.5, 0], ['L', 0.5, 1], ['L', 1, 1]],
    curvedConnector3: [['M', 0, 0], ['C', 0.5, 0, 0.5, 1, 1, 1]],
    rect: [['M', 0, 0], ['L', 1, 0], ['L', 1, 1], ['L', 0, 1], ['Z']],
  };
  return paths[shape.kind] ? { paths: [{ commands: paths[shape.kind], fill: shape.kind === 'rect', stroke: true }] } : null;
}
export function shapePathHandles(shape) {
  const handles = [];
  for (const [p, path] of (shape.path?.paths ?? []).entries()) {
    for (const section of pathSections(path.commands)) {
      let prev = null;
      for (let c = section.start; c <= section.last; c++) {
        const cmd = path.commands[c], endpoint = endOf(cmd);
        for (let slot = 1; slot < cmd.length; slot += 2) {
          const anchor = slot === cmd.length - 2;
          // The explicit closing Bezier endpoint and M are one vertex.
          if (anchor && section.seam && c === section.last) continue;
          const incomingQuadratic = cmd[0] === 'Q' && section.seam && c === section.last;
          const origin = anchor ? null : incomingQuadratic || slot !== 1 ? endpoint : prev;
          handles.push({ id: `${p}:${c}:${slot}`, p, c, slot, anchor, x: cmd[slot], y: cmd[slot + 1], origin });
        }
        prev = endpoint;
      }
    }
  }
  return handles;
}
export function shapeLocalPoint(shape, point) {
  const w = Math.max(1, shape.w), h = Math.max(1, shape.h), a = -(shape.rot ?? 0) * Math.PI / 180;
  const x = point[0] - shape.x - w / 2, y = point[1] - shape.y - h / 2;
  const u = (x * Math.cos(a) - y * Math.sin(a)) / w + 0.5, v = (x * Math.sin(a) + y * Math.cos(a)) / h + 0.5;
  return [shape.flip ? 1 - u : u, shape.flipV ? 1 - v : v];
}
export function moveShapePoint(path, handle, point, smooth = false, width = 1, height = 1) {
  const out = structuredClone(path), commands = out.paths[handle.p]?.commands, cmd = commands?.[handle.c];
  if (!cmd || !point.every(Number.isFinite)) return out;
  const section = pathSections(commands).find(s => handle.c >= s.start && handle.c <= s.last);
  const slot = handle.slot, dx = point[0] - cmd[slot], dy = point[1] - cmd[slot + 1];
  const shift = (c, i) => { if (c) { c[i] += dx; c[i + 1] += dy; } };
  if (handle.anchor) {
    if (section?.seam && (handle.c === section.start || handle.c === section.last)) {
      const first = commands[section.start], last = commands[section.last], next = commands[section.start + 1];
      const incoming = incomingSlot(last), outgoing = outgoingSlot(next);
      if (incoming >= 0) shift(last, incoming);
      if (outgoing >= 0 && (last !== next || incoming !== outgoing)) shift(next, outgoing);
      first[1] = last[last.length - 2] = point[0]; first[2] = last[last.length - 1] = point[1];
    } else {
      const incoming = incomingSlot(cmd); if (incoming >= 0) shift(cmd, incoming);
      const next = commands[handle.c + 1]; if (outgoingSlot(next) >= 0) shift(next, 1);
    }
  } else if (smooth) {
    const incomingQuadratic = cmd[0] === 'Q' && section?.seam && handle.c === section.last;
    const left = slot === 1 && !incomingQuadratic;
    const acrossStart = left && section?.seam && handle.c === section.start + 1;
    const acrossEnd = !left && section?.seam && handle.c === section.last;
    const anchor = left ? endOf(commands[handle.c - 1]) : endOf(cmd);
    const opposite = commands[acrossStart ? section.last : acrossEnd ? section.start + 1 : handle.c + (left ? -1 : 1)];
    const i = left ? incomingSlot(opposite) : outgoingSlot(opposite);
    if (anchor && i >= 0 && (opposite !== cmd || i !== slot)) {
      const metric = (a, b) => Math.hypot((a[0] - b[0]) * width, (a[1] - b[1]) * height);
      const d = metric(point, anchor), len = smooth === 'symmetric' ? d : metric(anchor, [opposite[i], opposite[i + 1]]);
      if (d > 1e-9) { opposite[i] = anchor[0] - (point[0] - anchor[0]) * len / d; opposite[i + 1] = anchor[1] - (point[1] - anchor[1]) * len / d; }
    }
  }
  cmd[slot] = point[0]; cmd[slot + 1] = point[1];
  return out;
}
export function deleteShapePoint(path, handle) {
  if (!handle?.anchor) return null;
  const out = structuredClone(path), commands = out.paths[handle.p]?.commands;
  if (!commands) return null;
  const section = pathSections(commands).find(s => handle.c >= s.start && handle.c <= s.last);
  if (!section) return null;
  const count = section.last - section.start + 1 - Number(section.seam);
  if (count <= (section.closed ? 3 : 2)) return null;
  const index = section.seam && handle.c === section.last ? section.start : handle.c;
  if (index === section.start) {
    const next = endOf(commands[section.start + 1]);
    if (section.seam) {
      const last = commands[section.last], old = endOf(last), control = incomingSlot(last);
      if (control >= 0) { last[control] += next[0] - old[0]; last[control + 1] += next[1] - old[1]; }
      last[last.length - 2] = next[0]; last[last.length - 1] = next[1];
    }
    commands.splice(section.start, 2, ['M', ...next]);
  } else commands.splice(index, 1);
  return out;
}
function splitCommand(start, cmd, t) {
  const end = endOf(cmd);
  if (cmd[0] === 'L') { const p = mix(start, end, t); return [['L', ...p], ['L', ...end]]; }
  if (cmd[0] === 'Q') {
    const a = mix(start, cmd.slice(1, 3), t), b = mix(cmd.slice(1, 3), end, t), p = mix(a, b, t);
    return [['Q', ...a, ...p], ['Q', ...b, ...end]];
  }
  const a = mix(start, cmd.slice(1, 3), t), b = mix(cmd.slice(1, 3), cmd.slice(3, 5), t), c = mix(cmd.slice(3, 5), end, t);
  const d = mix(a, b, t), e = mix(b, c, t), p = mix(d, e, t);
  return [['C', ...a, ...d, ...p], ['C', ...e, ...c, ...end]];
}
export function insertShapePoint(path, point, width = 1, height = 1) {
  let best = null;
  for (const [p, part] of path.paths.entries()) {
    let prev, first;
    for (const [c, cmd] of part.commands.entries()) {
      if (cmd[0] === 'M') { prev = first = endOf(cmd); continue; }
      const closing = cmd[0] === 'Z', segment = closing ? ['L', ...first] : cmd;
      if (!prev || !['L', 'C', 'Q'].includes(segment[0])) continue;
      if (closing && samePoint(prev, first)) continue; // Already closed by C/Q/L; no zero-length seam segment.
      for (let n = 1; n < 64; n++) {
        const t = n / 64, xy = endOf(splitCommand(prev, segment, t)[0]);
        const d = Math.hypot((xy[0] - point[0]) * width, (xy[1] - point[1]) * height);
        if (!best || d < best.d) best = { p, c, t, d, prev, segment, closing };
      }
      prev = endOf(segment);
    }
  }
  if (!best) return null;
  const out = structuredClone(path), parts = splitCommand(best.prev, best.segment, best.t);
  out.paths[best.p].commands.splice(best.c, 1, ...(best.closing ? [parts[0], ['Z']] : parts));
  return { path: out, handle: { p: best.p, c: best.c, slot: parts[0].length - 2, anchor: true } };
}

export function normalizeEditedShape(shape, path) {
  const bounds = normalizedPathBounds(path); if (!bounds) return { path };
  const ow = Math.max(1, shape.w), oh = Math.max(1, shape.h);
  let { minX, minY, maxX, maxY } = bounds;
  if (maxX - minX < 1 / ow) { const center = (minX + maxX) / 2; minX = center - 0.5 / ow; maxX = center + 0.5 / ow; }
  if (maxY - minY < 1 / oh) { const center = (minY + maxY) / 2; minY = center - 0.5 / oh; maxY = center + 0.5 / oh; }
  const spanX = maxX - minX, spanY = maxY - minY;
  const cx = (minX + maxX) / 2, cy = (minY + maxY) / 2;
  const dx = ((shape.flip ? 1 - cx : cx) - 0.5) * ow, dy = ((shape.flipV ? 1 - cy : cy) - 0.5) * oh;
  const a = (shape.rot ?? 0) * Math.PI / 180, w = spanX * ow, h = spanY * oh;
  const x = shape.x + ow / 2 + dx * Math.cos(a) - dy * Math.sin(a) - w / 2;
  const y = shape.y + oh / 2 + dx * Math.sin(a) + dy * Math.cos(a) - h / 2;
  const normalized = structuredClone(path);
  for (const part of normalized.paths) for (const command of part.commands) for (let i = 1; i < command.length; i += 2) { command[i] = (command[i] - minX) / spanX; command[i + 1] = (command[i + 1] - minY) / spanY; }
  return { x, y, w, h, path: normalized };
}
