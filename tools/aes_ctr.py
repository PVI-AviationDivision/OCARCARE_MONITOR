# -*- coding: utf-8 -*-
"""AES-256-CTR + HMAC-SHA256 thuần Python (không cần cài thêm thư viện).
Tương thích WebCrypto: AES-CTR (counter 16 byte, length=64) + HMAC SHA-256, khoá từ PBKDF2-SHA256."""
import base64, hashlib, hmac, os

_SBOX = [0] * 256
def _init():
    p = q = 1
    while True:
        p = p ^ ((p << 1) & 0xFF) ^ (0x1B if p & 0x80 else 0)
        q ^= q << 1; q ^= q << 2; q ^= q << 4; q &= 0xFF
        if q & 0x80: q ^= 0x09
        x = q ^ (((q << 1) | (q >> 7)) & 0xFF) ^ (((q << 2) | (q >> 6)) & 0xFF) ^ \
            (((q << 3) | (q >> 5)) & 0xFF) ^ (((q << 4) | (q >> 4)) & 0xFF)
        _SBOX[p] = x ^ 0x63
        if p == 1: break
    _SBOX[0] = 0x63
_init()

def _xt(a): return ((a << 1) ^ 0x1B) & 0xFF if a & 0x80 else a << 1
_T0 = []
for _s in _SBOX:
    _s2 = _xt(_s); _s3 = _s2 ^ _s
    _T0.append((_s2 << 24) | (_s << 16) | (_s << 8) | _s3)
_T1 = [((t >> 8) | (t << 24)) & 0xFFFFFFFF for t in _T0]
_T2 = [((t >> 16) | (t << 16)) & 0xFFFFFFFF for t in _T0]
_T3 = [((t >> 24) | (t << 8)) & 0xFFFFFFFF for t in _T0]

def _expand(key):
    nk = len(key) // 4; nr = nk + 6
    w = [int.from_bytes(key[4*i:4*i+4], "big") for i in range(nk)]
    rcon = 1
    for i in range(nk, 4 * (nr + 1)):
        t = w[i-1]
        if i % nk == 0:
            t = ((t << 8) | (t >> 24)) & 0xFFFFFFFF
            t = (_SBOX[t >> 24] << 24) | (_SBOX[(t >> 16) & 255] << 16) | (_SBOX[(t >> 8) & 255] << 8) | _SBOX[t & 255]
            t ^= rcon << 24; rcon = _xt(rcon)
        elif nk > 6 and i % nk == 4:
            t = (_SBOX[t >> 24] << 24) | (_SBOX[(t >> 16) & 255] << 16) | (_SBOX[(t >> 8) & 255] << 8) | _SBOX[t & 255]
        w.append(w[i-nk] ^ t)
    return w, nr

def _enc_block(w, nr, blk):
    s0, s1, s2, s3 = (int.from_bytes(blk[4*i:4*i+4], "big") ^ w[i] for i in range(4))
    T0, T1, T2, T3, S = _T0, _T1, _T2, _T3, _SBOX
    k = 4
    for _ in range(nr - 1):
        t0 = T0[s0 >> 24] ^ T1[(s1 >> 16) & 255] ^ T2[(s2 >> 8) & 255] ^ T3[s3 & 255] ^ w[k]
        t1 = T0[s1 >> 24] ^ T1[(s2 >> 16) & 255] ^ T2[(s3 >> 8) & 255] ^ T3[s0 & 255] ^ w[k+1]
        t2 = T0[s2 >> 24] ^ T1[(s3 >> 16) & 255] ^ T2[(s0 >> 8) & 255] ^ T3[s1 & 255] ^ w[k+2]
        t3 = T0[s3 >> 24] ^ T1[(s0 >> 16) & 255] ^ T2[(s1 >> 8) & 255] ^ T3[s2 & 255] ^ w[k+3]
        s0, s1, s2, s3 = t0, t1, t2, t3; k += 4
    out = b""
    for a, b_, c, d, wk in ((s0, s1, s2, s3, w[k]), (s1, s2, s3, s0, w[k+1]), (s2, s3, s0, s1, w[k+2]), (s3, s0, s1, s2, w[k+3])):
        v = (S[a >> 24] << 24) | (S[(b_ >> 16) & 255] << 16) | (S[(c >> 8) & 255] << 8) | S[d & 255]
        out += (v ^ wk).to_bytes(4, "big")
    return out

def aes_ctr(key, iv, data):
    """iv 16 byte; tăng 64 bit thấp (tương thích WebCrypto length=64)."""
    w, nr = _expand(key)
    hi, lo = iv[:8], int.from_bytes(iv[8:], "big")
    out = bytearray()
    for i in range(0, len(data), 16):
        ks = _enc_block(w, nr, hi + lo.to_bytes(8, "big"))
        chunk = data[i:i+16]
        out += bytes(x ^ y for x, y in zip(chunk, ks))
        lo = (lo + 1) & 0xFFFFFFFFFFFFFFFF
    return bytes(out)

ITER = 150_000

def encrypt(plaintext: bytes, password: str) -> dict:
    salt = os.urandom(16)
    iv = os.urandom(8) + b"\x00" * 8
    km = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, ITER, 64)
    ct = aes_ctr(km[:32], iv, plaintext)
    mac = hmac.new(km[32:], iv + ct, hashlib.sha256).digest()
    b = lambda x: base64.b64encode(x).decode()
    return dict(v=1, iter=ITER, salt=b(salt), iv=b(iv), ct=b(ct), mac=b(mac))

def decrypt(obj: dict, password: str) -> bytes:
    d = lambda k: base64.b64decode(obj[k])
    km = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), d("salt"), obj["iter"], 64)
    if not hmac.compare_digest(hmac.new(km[32:], d("iv") + d("ct"), hashlib.sha256).digest(), d("mac")):
        raise ValueError("Sai mật khẩu")
    return aes_ctr(km[:32], d("iv"), d("ct"))
