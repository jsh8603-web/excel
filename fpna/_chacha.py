"""
fpna._chacha — RFC 8439 ChaCha20-Poly1305 AEAD 순수 파이썬 이식(stdlib only).

⚠ 자작 암호 '설계'가 아니라 표준 알고리즘 이식이다. 정확성은 RFC 8439 공식
테스트벡터(test_vectors())로 박제한다. 무설치(vendor 불필요) 보장 — struct 만 사용.

한계: 파이썬 정수/바이트 연산은 상수시간이 아니라 타이밍 부채널에 노출된다.
→ 텍스트 페이로드(회사↔집 운반물) 보호 용도. 고위협(능동적 공격자) 모델엔 부적합.

상위 래퍼(fpna.crypto)가 scrypt KDF + secrets nonce + base64 armor 를 얹는다.
"""
from __future__ import annotations

import struct

_MASK32 = 0xffffffff
_CONST = (0x61707865, 0x3320646e, 0x79622d32, 0x6b206574)  # "expand 32-byte k"


def _rotl32(v: int, c: int) -> int:
    v &= _MASK32
    return ((v << c) & _MASK32) | (v >> (32 - c))


def _quarter(s, a, b, c, d):
    s[a] = (s[a] + s[b]) & _MASK32; s[d] = _rotl32(s[d] ^ s[a], 16)
    s[c] = (s[c] + s[d]) & _MASK32; s[b] = _rotl32(s[b] ^ s[c], 12)
    s[a] = (s[a] + s[b]) & _MASK32; s[d] = _rotl32(s[d] ^ s[a], 8)
    s[c] = (s[c] + s[d]) & _MASK32; s[b] = _rotl32(s[b] ^ s[c], 7)


def _chacha20_block(key: bytes, counter: int, nonce: bytes) -> bytes:
    """key=32B, nonce=12B, counter=uint32 → 64B 키스트림 블록."""
    state = (list(_CONST)
             + list(struct.unpack("<8I", key))
             + [counter & _MASK32]
             + list(struct.unpack("<3I", nonce)))
    w = list(state)
    for _ in range(10):  # 20 rounds = 10 × (column + diagonal)
        _quarter(w, 0, 4, 8, 12); _quarter(w, 1, 5, 9, 13)
        _quarter(w, 2, 6, 10, 14); _quarter(w, 3, 7, 11, 15)
        _quarter(w, 0, 5, 10, 15); _quarter(w, 1, 6, 11, 12)
        _quarter(w, 2, 7, 8, 13); _quarter(w, 3, 4, 9, 14)
    out = [(w[i] + state[i]) & _MASK32 for i in range(16)]
    return struct.pack("<16I", *out)


def chacha20_xor(key: bytes, counter: int, nonce: bytes, data: bytes) -> bytes:
    out = bytearray(len(data))
    for i in range(0, len(data), 64):
        ks = _chacha20_block(key, counter + (i // 64), nonce)
        chunk = data[i:i + 64]
        for j, b in enumerate(chunk):
            out[i + j] = b ^ ks[j]
    return bytes(out)


def _poly1305_mac(msg: bytes, otk: bytes) -> bytes:
    """one-time key otk=32B → 16B 태그 (RFC 8439 §2.5)."""
    r = int.from_bytes(otk[:16], "little") & 0x0ffffffc0ffffffc0ffffffc0fffffff
    s = int.from_bytes(otk[16:32], "little")
    p = (1 << 130) - 5
    acc = 0
    for i in range(0, len(msg), 16):
        blk = msg[i:i + 16]
        n = int.from_bytes(blk, "little") + (1 << (8 * len(blk)))  # append 0x01 bit
        acc = ((acc + n) * r) % p
    acc = (acc + s) & ((1 << 128) - 1)
    return acc.to_bytes(16, "little")


def _poly_keygen(key: bytes, nonce: bytes) -> bytes:
    return _chacha20_block(key, 0, nonce)[:32]


def _pad16(d: bytes) -> bytes:
    return b"" if len(d) % 16 == 0 else b"\x00" * (16 - len(d) % 16)


def aead_encrypt(key: bytes, nonce: bytes, plaintext: bytes, aad: bytes = b"") -> tuple[bytes, bytes]:
    """RFC 8439 §2.8 AEAD. 반환 (ciphertext, tag16)."""
    otk = _poly_keygen(key, nonce)
    ct = chacha20_xor(key, 1, nonce, plaintext)
    mac_data = (aad + _pad16(aad) + ct + _pad16(ct)
                + struct.pack("<Q", len(aad)) + struct.pack("<Q", len(ct)))
    return ct, _poly1305_mac(mac_data, otk)


def aead_decrypt(key: bytes, nonce: bytes, ciphertext: bytes, tag: bytes,
                 aad: bytes = b"") -> bytes:
    """태그 검증 후 평문 반환. 위조면 ValueError."""
    import hmac
    otk = _poly_keygen(key, nonce)
    mac_data = (aad + _pad16(aad) + ciphertext + _pad16(ciphertext)
                + struct.pack("<Q", len(aad)) + struct.pack("<Q", len(ciphertext)))
    expected = _poly1305_mac(mac_data, otk)
    if not hmac.compare_digest(expected, tag):  # 상수시간 비교
        raise ValueError("인증 실패: 태그 불일치(변조 또는 잘못된 키)")
    return chacha20_xor(key, 1, nonce, ciphertext)


def test_vectors() -> bool:
    """RFC 8439 §2.8.2 공식 AEAD 테스트벡터로 정확성 검증."""
    key = bytes(range(0x80, 0xa0))                      # 80..9f (32B)
    nonce = bytes.fromhex("070000004041424344454647")  # 12B
    aad = bytes.fromhex("50515253c0c1c2c3c4c5c6c7")
    plaintext = (b"Ladies and Gentlemen of the class of '99: If I could offer you "
                 b"only one tip for the future, sunscreen would be it.")
    exp_ct = bytes.fromhex(
        "d31a8d34648e60db7b86afbc53ef7ec2a4aded51296e08fea9e2b5a736ee62d6"
        "3dbea45e8ca9671282fafb69da92728b1a71de0a9e060b2905d6a5b67ecd3b36"
        "92ddbd7f2d778b8c9803aee328091b58fab324e4fad675945585808b4831d7bc"
        "3ff4def08e4b7a9de576d26586cec64b6116")
    exp_tag = bytes.fromhex("1ae10b594f09e26a7e902ecbd0600691")
    ct, tag = aead_encrypt(key, nonce, plaintext, aad)
    if ct != exp_ct or tag != exp_tag:
        return False
    # roundtrip + 위조 거부
    if aead_decrypt(key, nonce, ct, tag, aad) != plaintext:
        return False
    try:
        aead_decrypt(key, nonce, ct, bytes(16), aad)
        return False  # 위조 태그가 통과하면 실패
    except ValueError:
        pass
    return True


__all__ = ["chacha20_xor", "aead_encrypt", "aead_decrypt", "test_vectors"]


if __name__ == "__main__":
    print("RFC 8439 test vectors:", "PASS" if test_vectors() else "FAIL")
