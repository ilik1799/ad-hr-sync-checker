import getpass
import sys

import keyring
from jira import JIRA

try:
    from hrad_config import CONFIG
except ImportError:
    print("КРИТИЧЕСКАЯ ОШИБКА: не найден файл hrad_config.py в папке скрипта!")
    sys.exit(1)
# --- Константы ---
SERVICE_NAME = CONFIG["jira_keyring_service_name"]
JIRA_SERVER = CONFIG["jira_server"]


def test_connection(token: str) -> bool:
    """Проверяет соединение с Jira."""
    print(f"\n[?] Проверяем соединение с {JIRA_SERVER}...")
    try:
        jira = JIRA(
            server=JIRA_SERVER,
            options={
                "headers": {"Authorization": f"Bearer {token}"},
                "server_info": False,
                "verify": True,
                "check_update": False,
            },
        )
        user = jira.myself()
        print("[V] УСПЕШНО! Токен работает.")
        print(f"    Вы авторизованы как: {user.get('displayName')}")
        return True
    except Exception as e:
        print(f"[X] ОШИБКА: {e}")
        return False


def main() -> None:
    print("--- Настройка токена Jira ---")
    print(f"Сервер: {JIRA_SERVER}")
    print("-" * 35)
    try:
        print("Введите API-токен (ввод скрыт): ", end="", flush=True)
        token = getpass.getpass(prompt="").strip()
        if not token:
            print("\n[!] Токен не введён.")
            sys.exit(1)
        if test_connection(token):
            keyring.set_password(SERVICE_NAME, "jira_api_token", token)
            print("-" * 35)
            print("[V] Токен сохранён.")
        else:
            print("-" * 35)
            print("[!] Токен не сохранён.")
    except KeyboardInterrupt:
        print("\nОтменено.")
    except Exception as e:
        print(f"\n[X] Ошибка: {e}")


if __name__ == "__main__":
    main()
