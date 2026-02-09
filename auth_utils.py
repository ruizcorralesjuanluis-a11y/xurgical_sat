# auth_utils.py
from itsdangerous import URLSafeSerializer, BadSignature
from passlib.context import CryptContext

pwd_context = CryptContext(schemes=["pbkdf2_sha256"], deprecated="auto")

def hash_password(password: str) -> str:
    return pwd_context.hash(password)

def verify_password(password: str, password_hash: str) -> bool:
    return pwd_context.verify(password, password_hash)

def make_serializer(secret_key: str) -> URLSafeSerializer:
    return URLSafeSerializer(secret_key, salt="xurgical-sat-session")

def sign_session(serializer: URLSafeSerializer, user_id: int) -> str:
    return serializer.dumps({"user_id": user_id})

def read_session(serializer: URLSafeSerializer, token: str) -> int | None:
    try:
        data = serializer.loads(token)
        return int(data.get("user_id"))
    except (BadSignature, Exception):
        return None
