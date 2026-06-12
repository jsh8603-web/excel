"""
fpna.crypto — passphrase 기반 텍스트 대칭 암복호화.

구성(외부 검토된 표준만): scrypt KDF(stdlib) + ChaCha20-Poly1305 AEAD(fpna._chacha,
RFC 8439 이식) + secrets nonce + base64 armor. 회사↔집 운반물(yaml/코드/csv) 보호용.

armored 포맷 = base64( MAGIC(6) | salt(16) | nonce(12) | tag(16) | ciphertext ).
passphrase 는 코드/메일과 다른 채널(구두·전화)로 공유한다.

무설치: 전부 stdlib + 동봉코어. 회사 PC 에서 `py main.py decrypt` 로 그대로 복호화.
"""
from __future__ import annotations

import base64
import hashlib
import secrets

import fpna._bootstrap  # noqa: F401

from fpna._chacha import aead_encrypt, aead_decrypt

_MAGIC = b"FPNAC1"
_SCRYPT_N = 2 ** 15      # 비용 파라미터(128*N*r = 32MB)
_SCRYPT_R = 8
_SCRYPT_P = 1
_MAXMEM = 64 * 1024 * 1024


def _derive(passphrase: str, salt: bytes) -> bytes:
    return hashlib.scrypt(passphrase.encode("utf-8"), salt=salt,
                          n=_SCRYPT_N, r=_SCRYPT_R, p=_SCRYPT_P,
                          dklen=32, maxmem=_MAXMEM)


def encrypt_text(passphrase: str, plaintext: str) -> str:
    """평문 문자열 → armored base64 텍스트(한 줄)."""
    if not passphrase:
        raise ValueError("passphrase 가 비었습니다")
    salt = secrets.token_bytes(16)
    nonce = secrets.token_bytes(12)
    key = _derive(passphrase, salt)
    ct, tag = aead_encrypt(key, nonce, plaintext.encode("utf-8"))
    blob = _MAGIC + salt + nonce + tag + ct
    return base64.b64encode(blob).decode("ascii")


def decrypt_text(passphrase: str, armored: str) -> str:
    """armored base64 텍스트 → 평문. 변조/오답 passphrase 면 ValueError."""
    blob = base64.b64decode("".join(armored.split()))  # 공백/줄바꿈 허용
    if blob[:6] != _MAGIC:
        raise ValueError("형식 불일치: FPNAC1 헤더 없음")
    salt, nonce, tag, ct = blob[6:22], blob[22:34], blob[34:50], blob[50:]
    if len(salt) != 16 or len(nonce) != 12 or len(tag) != 16:
        raise ValueError("형식 손상: 헤더 길이 불일치")
    key = _derive(passphrase, salt)
    return aead_decrypt(key, nonce, ct, tag).decode("utf-8")


def encrypt_file(passphrase: str, in_path: str, out_path: str) -> int:
    # newline="" : 원본 줄끝(LF/CRLF)을 그대로 보존해 복호 시 바이트 동일성 유지
    with open(in_path, "r", encoding="utf-8", newline="") as fh:
        text = fh.read()
    armored = encrypt_text(passphrase, text)
    with open(out_path, "w", encoding="utf-8", newline="") as fh:
        fh.write(armored + "\n")
    return len(armored)


def decrypt_file(passphrase: str, in_path: str, out_path: str) -> int:
    with open(in_path, "r", encoding="utf-8", newline="") as fh:
        armored = fh.read()
    text = decrypt_text(passphrase, armored)
    with open(out_path, "w", encoding="utf-8", newline="") as fh:
        fh.write(text)
    return len(text)


def selftest() -> bool:
    """코어 테스트벡터 + 래퍼 roundtrip + 위조/오답 거부."""
    from fpna._chacha import test_vectors
    if not test_vectors():
        return False
    msg = "예실대비 brief: 매출 1,234 (단위:천원) — 회사↔집 운반 테스트 ✓"
    arm = encrypt_text("hunter2-비밀", msg)
    if decrypt_text("hunter2-비밀", arm) != msg:
        return False
    try:
        decrypt_text("wrong-pw", arm)
        return False  # 오답 passphrase 가 통과하면 실패
    except ValueError:
        return True


__all__ = ["encrypt_text", "decrypt_text", "encrypt_file", "decrypt_file", "selftest"]
