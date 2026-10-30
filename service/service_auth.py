import hashlib
import hmac
import secrets
import sqlite3
from dataclasses import dataclass

from service.service_master import ADMIN_ROLE, Staff

HASH_ITERATIONS = 200_000
MIN_PASSWORD_LENGTH = 6
# 読み違えやすい 0/o/1/l/i を除いた文字
TEMPORARY_PASSWORD_CHARACTERS = 'abcdefghjkmnpqrstuvwxyz23456789'
TEMPORARY_PASSWORD_LENGTH = 8


@dataclass
class LoginResult:
    staff: Staff
    must_change_password: bool


def hash_password(password: str, salt: bytes) -> str:
    return hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, HASH_ITERATIONS).hex()


def make_temporary_password() -> str:
    return ''.join(secrets.choice(TEMPORARY_PASSWORD_CHARACTERS) for _ in range(TEMPORARY_PASSWORD_LENGTH))


def login_required(connection: sqlite3.Connection) -> bool:
    """パスワードを設定した在籍中の管理者がいれば、ログインを求める"""
    return connection.execute(
        'SELECT COUNT(*) FROM staff WHERE role = ? AND active = 1 AND password_hash IS NOT NULL',
        (ADMIN_ROLE,)).fetchone()[0] > 0


def set_password(connection: sqlite3.Connection, staff_id: int, password: str, must_change: bool) -> None:
    salt = secrets.token_bytes(16)
    connection.execute(
        'UPDATE staff SET password_hash = ?, password_salt = ?, must_change_password = ? WHERE id = ?',
        (hash_password(password, salt), salt.hex(), int(must_change), staff_id))


def authenticate(connection: sqlite3.Connection, emr_id: str, password: str) -> LoginResult | None:
    """電子カルテIDとパスワードが合う在籍中の担当者を返す。合わなければ None"""
    row = connection.execute(
        'SELECT id, name, display_order, active, emr_id, role, password_hash, password_salt, must_change_password'
        ' FROM staff WHERE emr_id = ? AND active = 1 AND password_hash IS NOT NULL', (emr_id,)).fetchone()
    if row is None:
        return None
    id_, name, order, active, emr_id, role, password_hash, salt, must_change = row
    if not hmac.compare_digest(hash_password(password, bytes.fromhex(salt)), password_hash):
        return None
    return LoginResult(Staff(id_, name, order, bool(active), emr_id, role), bool(must_change))
