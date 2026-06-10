"""
Pure Python Track-Key signing for WMS API.
Reverse-engineered from app.d70ca2a3.js — no Node.js needed.
"""

import time as _time

MASK32 = 0xFFFFFFFF


def _add(a: int, b: int) -> int:
    return (a + b) & MASK32


def _rol(v: int, s: int) -> int:
    return ((v << s) | (v >> (32 - s))) & MASK32


def _ff(a, b, c, d, x, s, ac):
    return _add(_rol(_add(_add(a, (b & c) | ((~b & MASK32) & d)), _add(x, ac)), s), b)


def _gg(a, b, c, d, x, s, ac):
    return _add(_rol(_add(_add(a, (b & d) | (c & (~d & MASK32))), _add(x, ac)), s), b)


def _hh(a, b, c, d, x, s, ac):
    return _add(_rol(_add(_add(a, b ^ c ^ d), _add(x, ac)), s), b)


def _ii(a, b, c, d, x, s, ac):
    return _add(_rol(_add(_add(a, c ^ (b | (~d & MASK32))), _add(x, ac)), s), b)


def _md5_raw(text: str) -> list[int]:
    """MD5 producing 4 x 32-bit words, matching the JS implementation exactly."""
    # JS f(): UTF-8 encode to byte values
    byte_vals: list[int] = []
    for ch in text:
        cp = ord(ch)
        if cp <= 0x7F:
            byte_vals.append(cp)
        elif cp <= 0x7FF:
            byte_vals.extend([0xC0 | (cp >> 6), 0x80 | (cp & 0x3F)])
        elif cp <= 0xFFFF:
            byte_vals.extend([0xE0 | (cp >> 12), 0x80 | ((cp >> 6) & 0x3F), 0x80 | (cp & 0x3F)])
        else:
            byte_vals.extend([0xF0 | (cp >> 18), 0x80 | ((cp >> 12) & 0x3F),
                              0x80 | ((cp >> 6) & 0x3F), 0x80 | (cp & 0x3F)])

    # JS g(): pack bytes into 32-bit words (little-endian, 8 bits per char)
    # JS allocates sparse array; we pre-allocate enough space for padding
    n_words = (len(byte_vals) + 3) // 4
    pad_idx = 14 + (((len(byte_vals) * 8 + 64) >> 9) << 4)
    words = [0] * max(n_words, pad_idx + 1)
    for idx, bv in enumerate(byte_vals):
        words[idx >> 2] |= bv << ((idx % 4) * 8)

    # Pad: set bit after message
    bit_len = len(byte_vals) * 8
    words[bit_len >> 5] |= 0x80 << (bit_len % 32)
    words[14 + (((bit_len + 64) >> 9) << 4)] = bit_len

    # MD5 rounds
    a = 1732584193
    b = 3989678985
    c = 2562383614
    d = 271733878

    for s in range(0, len(words), 16):
        # Ensure block has 16 elements
        block = [words[s + k] if s + k < len(words) and words[s + k] is not None else 0 for k in range(16)]

        oa, ob, oc, od = a, b, c, d

        a = _ff(a, b, c, d, block[0], 7, -680876936)
        d = _ff(d, a, b, c, block[1], 12, -389564586)
        c = _ff(c, d, a, b, block[2], 17, 606105819)
        b = _ff(b, c, d, a, block[3], 22, -1044525330)
        a = _ff(a, b, c, d, block[4], 7, -176418897)
        d = _ff(d, a, b, c, block[5], 12, 1200080426)
        c = _ff(c, d, a, b, block[6], 17, -1473231341)
        b = _ff(b, c, d, a, block[7], 22, -45705983)
        a = _ff(a, b, c, d, block[8], 7, 1770035416)
        d = _ff(d, a, b, c, block[9], 12, -1958414417)
        c = _ff(c, d, a, b, block[10], 17, -42063)
        b = _ff(b, c, d, a, block[11], 22, -1990404162)
        a = _ff(a, b, c, d, block[12], 7, 1804603682)
        d = _ff(d, a, b, c, block[13], 12, -40341101)
        c = _ff(c, d, a, b, block[14], 17, -1502002290)
        b = _ff(b, c, d, a, block[15], 22, 1236535329)

        a = _gg(a, b, c, d, block[1], 5, -165796510)
        d = _gg(d, a, b, c, block[6], 9, -1069501632)
        c = _gg(c, d, a, b, block[11], 14, 643717713)
        b = _gg(b, c, d, a, block[0], 20, -373897302)
        a = _gg(a, b, c, d, block[5], 5, -701558691)
        d = _gg(d, a, b, c, block[10], 9, 38016083)
        c = _gg(c, d, a, b, block[15], 14, -660478335)
        b = _gg(b, c, d, a, block[4], 20, -405537848)
        a = _gg(a, b, c, d, block[9], 5, 568446438)
        d = _gg(d, a, b, c, block[14], 9, -1019803690)
        c = _gg(c, d, a, b, block[3], 14, -187633961)
        b = _gg(b, c, d, a, block[8], 20, 1163531501)
        a = _gg(a, b, c, d, block[13], 5, -1444681467)
        d = _gg(d, a, b, c, block[2], 9, -51403784)
        c = _gg(c, d, a, b, block[7], 14, 1735328473)
        b = _gg(b, c, d, a, block[12], 20, -1926607734)

        a = _hh(a, b, c, d, block[5], 4, -378558)
        d = _hh(d, a, b, c, block[8], 11, -2022574463)
        c = _hh(c, d, a, b, block[11], 16, 1839030562)
        b = _hh(b, c, d, a, block[14], 23, -35309556)
        a = _hh(a, b, c, d, block[1], 4, -1530992060)
        d = _hh(d, a, b, c, block[4], 11, 1272893353)
        c = _hh(c, d, a, b, block[7], 16, -155497632)
        b = _hh(b, c, d, a, block[10], 23, -1094730640)
        a = _hh(a, b, c, d, block[13], 4, 681279174)
        d = _hh(d, a, b, c, block[0], 11, -358537222)
        c = _hh(c, d, a, b, block[3], 16, -722521979)
        b = _hh(b, c, d, a, block[6], 23, 76029189)
        a = _hh(a, b, c, d, block[9], 4, -640364487)
        d = _hh(d, a, b, c, block[12], 11, -421815835)
        c = _hh(c, d, a, b, block[15], 16, 530742520)
        b = _hh(b, c, d, a, block[2], 23, -995338651)

        a = _ii(a, b, c, d, block[0], 6, -198630844)
        d = _ii(d, a, b, c, block[7], 10, 1126891415)
        c = _ii(c, d, a, b, block[14], 15, -1416354905)
        b = _ii(b, c, d, a, block[5], 21, -57434055)
        a = _ii(a, b, c, d, block[12], 6, 1700485571)
        d = _ii(d, a, b, c, block[3], 10, -1894986606)
        c = _ii(c, d, a, b, block[10], 15, -1051523)
        b = _ii(b, c, d, a, block[1], 21, -2054922799)
        a = _ii(a, b, c, d, block[8], 6, 1873313359)
        d = _ii(d, a, b, c, block[15], 10, -30611744)
        c = _ii(c, d, a, b, block[6], 15, -1560198380)
        b = _ii(b, c, d, a, block[13], 21, 1309151649)
        a = _ii(a, b, c, d, block[4], 6, -145523070)
        d = _ii(d, a, b, c, block[11], 10, -1120210379)
        c = _ii(c, d, a, b, block[2], 15, 718787259)
        b = _ii(b, c, d, a, block[9], 21, -343485551)

        a = _add(a, oa)
        b = _add(b, ob)
        c = _add(c, oc)
        d = _add(d, od)

    return [a, b, c, d]


_HEX = '0123456789abcdef'
_B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'


def _md5_hex(text: str) -> str:
    words = _md5_raw(text)
    out = []
    for i in range(4 * len(words)):
        out.append(_HEX[(words[i >> 2] >> ((i % 4) * 8 + 4)) & 15])
        out.append(_HEX[(words[i >> 2] >> ((i % 4) * 8)) & 15])
    return ''.join(out)


def _md5_base64(text: str) -> str:
    words = _md5_raw(text)
    out = []
    total_bits = 32 * len(words)
    for i in range(0, 4 * len(words), 3):
        val = ((words[i >> 2] >> ((i % 4) * 8)) & 0xFF) << 16
        if i + 1 < 4 * len(words):
            val |= ((words[(i + 1) >> 2] >> (((i + 1) % 4) * 8)) & 0xFF) << 8
        if i + 2 < 4 * len(words):
            val |= ((words[(i + 2) >> 2] >> (((i + 2) % 4) * 8)) & 0xFF)
        for s in range(4):
            if 8 * i + 6 * s > total_bits:
                break
            out.append(_B64[(val >> (6 * (3 - s))) & 63])
    return ''.join(out)


def generate_track_key(body_text: str, timestamp_ms: int | None = None) -> str:
    """Generate Track-Key header value for WMS API requests."""
    if timestamp_ms is None:
        timestamp_ms = int(_time.time() * 1000)
    ts_str = str(timestamp_ms)
    return _md5_base64(ts_str) + _md5_hex(body_text) + _md5_hex(ts_str)


# Self-test
if __name__ == '__main__':
    ts = 1718000000000
    body = '{"current":1,"size":100,"whCode":"US02"}'
    key = generate_track_key(body, ts)
    expected = 'LA15u7qylZ4SxnECpXd8ww90b5201a2b87ed408c3198b112931a2a2c0d79bbbab2959e12c67102a5777cc3'
    assert key == expected, f"Mismatch!\n  got:    {key}\n  expect: {expected}"
    print(f"OK  {key}")
