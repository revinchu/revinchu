# 테스트 글꼴(EOT, MTX 압축) 만들기: pyftsubset 로 DejaVuSerif-Bold 를 영문 부분 집합(sub.ttf)으로 만든 뒤
# python3 make-mtx-eot.py sub.ttf sub.eot 1   (CTF 부호화는 이 스크립트, LZCOMP 압축은 libeot 의 lzcomp.c 를 COMPRESS_ON 으로 빌드한 pack 도구)
# 결과 글꼴은 DejaVu 글꼴 사용권(Bitstream Vera)으로 재배포 가능합니다.
import sys, struct, subprocess, os
from fontTools.ttLib import TTFont
src, out, xor = sys.argv[1], sys.argv[2], sys.argv[3] == '1'
PACK = os.path.join(os.path.dirname(__file__), '..', 'repo', 'pack')
f = TTFont(src)
raw = {t: f.reader[t] for t in f.reader.keys()}
num_glyphs = f['maxp'].numGlyphs
loca_long = f['head'].indexToLocFormat == 1
loca = raw['loca']
if loca_long: offs = [struct.unpack('>I', loca[i*4:i*4+4])[0] for i in range(num_glyphs+1)]
else: offs = [struct.unpack('>H', loca[i*2:i*2+2])[0]*2 for i in range(num_glyphs+1)]
glyf = raw['glyf']

# triplet table (from fonts.js logic)
T = []
for i in range(10): T.append((2,0,8,0,(i>>1)*256,0,1 if i&1 else -1))
for i in range(10): T.append((2,8,0,(i>>1)*256,0,1 if i&1 else -1,0))
d4=[1,17,33,49]
for j in range(64): T.append((2,4,4,d4[j>>4],d4[(j>>2)&3],1 if j&1 else -1,1 if j&2 else -1))
d8=[1,257,513]
for j in range(36): T.append((3,8,8,d8[j//12],d8[(j//4)%3],1 if j&1 else -1,1 if j&2 else -1))
for j in range(4): T.append((4,12,12,0,0,1 if j&1 else -1,1 if j&2 else -1))
for j in range(4): T.append((5,16,16,0,0,1 if j&1 else -1,1 if j&2 else -1))

def fits(v, bits, delta, sign):
    if bits == 0: return v == 0
    if sign == 0: return False
    if v == 0 and sign < 0: return False
    if (v < 0) != (sign < 0) and v != 0: return False
    a = abs(v) - delta
    return 0 <= a < (1 << bits)
def triplet(dx, dy):
    for i,(nb,xb,yb,ddx,ddy,sx,sy) in enumerate(T):
        okx = (xb == 0 and dx == 0) or (xb and fits(dx, xb, ddx, sx))
        oky = (yb == 0 and dy == 0) or (yb and fits(dy, yb, ddy, sy))
        if okx and oky:
            vx = abs(dx) - ddx if xb else 0
            vy = abs(dy) - ddy if yb else 0
            acc = (vx << yb) | vy
            return i, acc.to_bytes(nb-1, 'big')
    raise Exception('no triplet %d %d' % (dx, dy))
def u255(v):
    if v < 253: return bytes([v])
    if v < 506: return bytes([255, v-253])
    if v < 762: return bytes([254, v-506])
    return bytes([253]) + struct.pack('>H', v)
def s255(v):
    if -250 < v < 250 and v >= 0: return bytes([v])
    if 0 < -v < 250: return bytes([250, -v])
    return bytes([253]) + struct.pack('>h', v)
def split_push(ins):
    vals = []; p = 0
    while p < len(ins):
        op = ins[p]
        if 0xb0 <= op <= 0xb7: n = op - 0xaf; vals += list(ins[p+1:p+1+n]); p += 1 + n
        elif 0xb8 <= op <= 0xbf: n = op - 0xb7; vals += [struct.unpack('>h', ins[p+1+2*k:p+3+2*k])[0] for k in range(n)]; p += 1 + 2*n
        elif op == 0x40: n = ins[p+1]; vals += list(ins[p+2:p+2+n]); p += 2 + n
        elif op == 0x41: n = ins[p+1]; vals += [struct.unpack('>h', ins[p+2+2*k:p+4+2*k])[0] for k in range(n)]; p += 2 + 2*n
        else: break
    return vals, ins[p:]
s0g = bytearray(); s1 = bytearray(); s2 = bytearray()
def put_ins(ins):
    vals, rest = split_push(ins)
    s0g.extend(u255(len(vals)))
    for v in vals: s1.extend(s255(v))
    s0g.extend(u255(len(rest)))
    s2.extend(rest)
for gi in range(num_glyphs):
    g = glyf[offs[gi]:offs[gi+1]]
    if not g: s0g.extend(struct.pack('>h', 0)); continue
    nc, xmin, ymin, xmax, ymax = struct.unpack('>hhhhh', g[:10])
    if nc < 0:
        s0g.extend(struct.pack('>hhhhh', -1, xmin, ymin, xmax, ymax))
        p = 10
        while True:
            flags = struct.unpack('>H', g[p:p+2])[0]
            n = 4 + (4 if flags & 1 else 2) + (8 if flags & 0x80 else 4 if flags & 0x40 else 2 if flags & 8 else 0)
            s0g.extend(g[p:p+n]); p += n
            if not flags & 0x20: break
        if flags & 0x100:
            ni = struct.unpack('>H', g[p:p+2])[0]
            put_ins(g[p+2:p+2+ni])
        continue
    # simple glyph: decode
    p = 10
    ends = list(struct.unpack('>%dH' % nc, g[p:p+2*nc])); p += 2*nc
    ni = struct.unpack('>H', g[p:p+2])[0]; ins = g[p+2:p+2+ni]; p += 2 + ni
    npts = ends[-1] + 1
    fl = []
    while len(fl) < npts:
        f0 = g[p]; p += 1; fl.append(f0)
        if f0 & 8:
            r = g[p]; p += 1; fl += [f0] * r
    xs = []; x = 0
    for f0 in fl:
        if f0 & 2: v = g[p]; p += 1; v = v if f0 & 0x10 else -v
        elif f0 & 0x10: v = 0
        else: v = struct.unpack('>h', g[p:p+2])[0]; p += 2
        xs.append(v)
    ys = []
    for f0 in fl:
        if f0 & 4: v = g[p]; p += 1; v = v if f0 & 0x20 else -v
        elif f0 & 0x20: v = 0
        else: v = struct.unpack('>h', g[p:p+2])[0]; p += 2
        ys.append(v)
    if gi % 5 == 0: s0g.extend(struct.pack('>hhhhhh', 0x7fff, nc, xmin, ymin, xmax, ymax))
    else: s0g.extend(struct.pack('>h', nc))
    prev = -1
    for k, e in enumerate(ends):
        s0g.extend(u255(e if k == 0 else e - prev)); prev = e
    codes = []; coord = bytearray()
    for k in range(npts):
        idx, b = triplet(xs[k], ys[k])
        codes.append(idx | (0 if fl[k] & 1 else 0x80)); coord.extend(b)
    s0g.extend(bytes(codes)); s0g.extend(coord)
    put_ins(ins)
# cvt
cvt = None
if 'cvt ' in raw:
    vals = struct.unpack('>%dh' % (len(raw['cvt '])//2), raw['cvt '])
    cvt = bytearray(struct.pack('>H', len(vals))); last = 0
    for v in vals:
        d = v - last; last = v
        if 0 <= d < 238: cvt.append(d)
        elif 238 <= d < 238*9 and False: pass
        elif -2142 < d < 0 and (-d) % 1 == 0 and -d < 238*9:
            k = (-d) // 238; cvt += bytes([239 + k, (-d) - 238*k])
        elif 238 <= d < 238*9 + 238:
            k = d // 238 - 1
            if 0 <= k <= 7: cvt += bytes([248 + k, d - 238*(k+1)])
            else: cvt += bytes([238]) + struct.pack('>h', d)
        else: cvt += bytes([238]) + struct.pack('>h', d)
tags = sorted(t for t in raw if t not in ('loca', 'DSIG'))
body = {}
for t in tags:
    if t == 'glyf': body[t] = bytes(s0g)
    elif t == 'cvt ' and cvt is not None: body[t] = bytes(cvt)
    else: body[t] = raw[t]
hdr = struct.pack('>IHHHH', 0x00010000, len(tags), 0, 0, 0)
off = 12 + 16 * len(tags); dirb = b''; data = b''
for t in tags:
    d = body[t]
    dirb += t.encode('latin1') + struct.pack('>III', 0, off + len(data), len(d))
    data += d + b'\0' * ((4 - len(d) % 4) % 4)
c0 = hdr + dirb + data
blocks = []
for i, s in enumerate([c0, bytes(s1), bytes(s2)]):
    open('/tmp/_c%d' % i, 'wb').write(s)
    subprocess.check_call([PACK, '/tmp/_c%d' % i, '/tmp/_z%d' % i])
    blocks.append(open('/tmp/_z%d' % i, 'rb').read())
o2 = 10 + len(blocks[0]); o3 = o2 + len(blocks[1])
mtx = bytes([3]) + (0).to_bytes(3, 'big') + o2.to_bytes(3, 'big') + o3.to_bytes(3, 'big') + blocks[0] + blocks[1] + blocks[2]
fd = bytes(b ^ 0x50 for b in mtx) if xor else mtx
fam = 'Test Family'.encode('utf-16le')
pre = struct.pack('<III', 0, 0, 0x00020001)  # sizes later
flags = 4 | (0x10000000 if xor else 0)
h = struct.pack('<I', flags) + bytes(10) + bytes([1, 0]) + struct.pack('<I', 400) + struct.pack('<H', 0) + struct.pack('<H', 0x504C) + bytes(16) + bytes(8) + bytes(4) + bytes(16) + struct.pack('<H', 0)
h += struct.pack('<H', len(fam)) + fam + struct.pack('<H', 0) + struct.pack('<H', 0) + struct.pack('<H', 0) + struct.pack('<H', 0) + struct.pack('<H', 0) + struct.pack('<H', 0) + struct.pack('<H', 0) + struct.pack('<H', 0)
total = 12 + len(h) + len(fd)
eot = struct.pack('<III', total, len(fd), 0x00020001) + h + fd
open(out, 'wb').write(eot)
print('glyphs', num_glyphs, 'ctf sizes', len(c0), len(s1), len(s2), 'eot', len(eot))
