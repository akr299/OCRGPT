from __future__ import annotations

from pathlib import Path

from cryptography.fernet import Fernet, InvalidToken


class ApiKeyManager:
    """OpenAI APIキーを暗号化してローカル保存するヘルパー。"""

    def __init__(self, config_dir: Path | None = None):
        base_dir = Path(__file__).resolve().parent
        self.config_dir = config_dir or (base_dir / "config")
        self.key_file = self.config_dir / "key.bin"
        self.encrypted_api_key_file = self.config_dir / "api_key.enc"

    def generate_or_load_encryption_key(self) -> bytes:
        """暗号鍵を読み込む。なければ新規作成して保存する。"""
        self.config_dir.mkdir(parents=True, exist_ok=True)
        if self.key_file.exists():
            return self.key_file.read_bytes()

        key = Fernet.generate_key()
        self.key_file.write_bytes(key)
        return key

    def encrypt_api_key(self, api_key: str) -> bytes:
        """APIキー文字列をFernetで暗号化する。"""
        key = self.generate_or_load_encryption_key()
        fernet = Fernet(key)
        return fernet.encrypt(api_key.strip().encode("utf-8"))

    def decrypt_api_key(self) -> str:
        """保存済み暗号化APIキーを復号して返す。"""
        if not self.key_file.exists() or not self.encrypted_api_key_file.exists():
            raise FileNotFoundError("APIキー設定が見つかりません。")

        key = self.key_file.read_bytes()
        encrypted = self.encrypted_api_key_file.read_bytes()
        fernet = Fernet(key)
        decrypted = fernet.decrypt(encrypted)
        api_key = decrypted.decode("utf-8").strip()
        if not api_key:
            raise ValueError("復号したAPIキーが空です。")
        return api_key

    def save_api_key(self, api_key: str) -> None:
        """APIキーを暗号化して保存する。"""
        cleaned = api_key.strip()
        if not cleaned:
            raise ValueError("APIキーが空です。")

        encrypted = self.encrypt_api_key(cleaned)
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.encrypted_api_key_file.write_bytes(encrypted)

    def load_api_key(self) -> str | None:
        """保存済みAPIキーを返す。欠損・破損時は None を返す。"""
        try:
            return self.decrypt_api_key()
        except (FileNotFoundError, InvalidToken, ValueError):
            return None

    def delete_api_key(self) -> None:
        """保存済み暗号化APIキーを削除する。"""
        if self.encrypted_api_key_file.exists():
            self.encrypted_api_key_file.unlink()
