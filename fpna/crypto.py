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
    """armored base64 텍스트 → 평문. 변조/오답 passphrase 면 ValueError.

    메일 part 마커(-----FPNAC1 ... PART k/n-----)가 섞여 있으면 자동으로
    추출·순서정렬·결합한다. 여러 메일 본문을 한 파일에 붙여넣어도 복원된다.
    """
    if "FPNAC1" in armored and "PART" in armored:
        armored = from_mail_text(armored)
    blob = base64.b64decode("".join(armored.split()))  # 공백/줄바꿈 허용
    if blob[:6] != _MAGIC:
        raise ValueError("형식 불일치: FPNAC1 헤더 없음")
    salt, nonce, tag, ct = blob[6:22], blob[22:34], blob[34:50], blob[50:]
    if len(salt) != 16 or len(nonce) != 12 or len(tag) != 16:
        raise ValueError("형식 손상: 헤더 길이 불일치")
    key = _derive(passphrase, salt)
    return aead_decrypt(key, nonce, ct, tag).decode("utf-8")


# -----------------------------------------------------------------------------
# 메일 본문 운반 (파일 첨부 X — 텍스트로 붙여넣기). 줄 wrap + 메일당 줄수 한정.
# -----------------------------------------------------------------------------
_WRAP = 76                 # PEM 스타일 줄 폭
_DEFAULT_MAX_LINES = 500   # 메일당 최대 줄수(헤더/푸터 포함). 초과분만 part 분할.

import re as _re


def to_mail_text(armored: str, *, wrap: int = _WRAP,
                 max_lines: int = _DEFAULT_MAX_LINES, msg_id: str = "MSG") -> list[str]:
    """armored(한 줄) → 메일 본문 part 리스트. 줄 wrap 후 메일당 max_lines 로 분할.

    각 part = 헤더 + wrapped base64 + 푸터. 한 통이면 part 1/1.
    너무 많이 초과하지 않게 본문 줄수를 max_lines-2(헤더/푸터)로 캡한다.
    """
    body = "".join(armored.split())
    lines = [body[i:i + wrap] for i in range(0, len(body), wrap)] or [""]
    cap = max(1, max_lines - 2)
    chunks = [lines[i:i + cap] for i in range(0, len(lines), cap)]
    n = len(chunks)
    parts = []
    for k, ch in enumerate(chunks, 1):
        head = "-----FPNAC1 %s PART %d/%d-----" % (msg_id, k, n)
        foot = "-----FPNAC1 %s END %d/%d-----" % (msg_id, k, n)
        parts.append("\n".join([head] + ch + [foot]))
    return parts


def from_mail_text(text: str) -> str:
    """part 마커가 섞인 텍스트(여러 메일 합본 가능) → 순서정렬·결합한 base64.

    마커가 없으면 공백만 제거해 반환(단일 armored 호환).
    """
    blocks = _re.findall(
        r"-----FPNAC1\s+\S+\s+PART\s+(\d+)/(\d+)-----\s*(.*?)\s*"
        r"-----FPNAC1\s+\S+\s+END\s+\1/\2-----",
        text, _re.S)
    if not blocks:
        return "".join(text.split())
    blocks.sort(key=lambda b: int(b[0]))
    return "".join("".join(b[2].split()) for b in blocks)


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
