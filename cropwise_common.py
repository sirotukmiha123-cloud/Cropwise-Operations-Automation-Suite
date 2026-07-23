"""
cropwise_common.py — спільний модуль для допоміжних скриптів Cropwise Automation Suite.

Централізує:
- безпечне завантаження API-токена (ніколи не зберігається в коді),
- побудову заголовків запиту,
- обгортку над requests з обробкою HTTP 429 (rate limit) та мережевих збоїв.

Покладіть цей файл у ту саму папку, що й інші скрипти набору — вони імпортують його напряму.
"""

import json
import os
import sys
import time
import requests

API_BASE = "https://operations.cropwise.com"

CONFIG_EXAMPLE = '{\n  "cropwise_token": "ВАШ_ТОКЕН_ТУТ"\n}'

# За замовчуванням SSL-сертифікати перевіряються (безпечно). Якщо ваша корпоративна мережа
# використовує проксі з власним сертифікатом і requests відмовляється підключатись — можна
# ТИМЧАСОВО вимкнути перевірку змінною середовища, а не хардкодити verify=False в коді:
#   export CROPWISE_DISABLE_SSL_VERIFY=1
SSL_VERIFY = os.environ.get("CROPWISE_DISABLE_SSL_VERIFY", "").strip() != "1"
if not SSL_VERIFY:
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    print("⚠️  Перевірку SSL-сертифікатів вимкнено через CROPWISE_DISABLE_SSL_VERIFY=1 "
          "(не рекомендується поза довіреною мережею).")


def load_cropwise_token():
    """
    Порядок пошуку токена:
    1. Змінна середовища CROPWISE_TOKEN
    2. Файл config.json поруч зі скриптом (ключ "cropwise_token")
    3. Ручне введення (одноразово, для цього запуску)

    Токен НІКОЛИ не зберігається в коді. Отримати токен: особистий кабінет Cropwise ->
    Налаштування -> API.
    """
    env_token = os.environ.get("CROPWISE_TOKEN")
    if env_token:
        return env_token.strip()

    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            if cfg.get("cropwise_token"):
                return str(cfg["cropwise_token"]).strip()
        except Exception as e:
            print(f"⚠️  Не вдалося прочитати config.json ({e}).")

    print("🔐 Токен Cropwise не знайдено.")
    print("   Варіант 1: встановіть змінну середовища CROPWISE_TOKEN")
    print(f"   Варіант 2: створіть файл config.json поруч зі скриптом за зразком:\n   {CONFIG_EXAMPLE}")
    manual_token = input("   Або введіть токен зараз (одноразово, для цього запуску): ").strip()
    if manual_token:
        return manual_token
    print("❌ Токен не надано. Завершення роботи.")
    sys.exit(1)


CROPWISE_TOKEN = load_cropwise_token()


def get_headers():
    return {
        "X-User-Api-Token": str(CROPWISE_TOKEN).strip(),
        "Accept": "application/json",
        "Content-Type": "application/json; charset=utf-8",
    }


def api_request(method, url, max_retries=3, **kwargs):
    """
    Обгортка над requests.request з обробкою HTTP 429 (Too Many Requests): чекає Retry-After
    (або зростаючу паузу) і повторює запит до max_retries разів. Також повторює спробу при
    мережевих таймаутах/збоях з'єднання. Автоматично підставляє SSL_VERIFY, якщо verify не
    передано явно викликачем.
    """
    kwargs.setdefault("verify", SSL_VERIFY)
    last_exc = None
    for attempt in range(max_retries + 1):
        try:
            resp = requests.request(method, url, **kwargs)
        except requests.exceptions.RequestException as e:
            last_exc = e
            if attempt < max_retries:
                wait_s = 1.5 * (attempt + 1)
                print(f"  ⏳ Мережева помилка ({e.__class__.__name__}), повтор через {wait_s:.1f} с...")
                time.sleep(wait_s)
                continue
            raise
        if resp.status_code == 429 and attempt < max_retries:
            retry_after = resp.headers.get("Retry-After")
            try:
                wait_s = float(retry_after) if retry_after else 2.0 * (attempt + 1)
            except (TypeError, ValueError):
                wait_s = 2.0 * (attempt + 1)
            print(f"  ⏳ Сервер повернув 429 (забагато запитів). Пауза {wait_s:.1f} с. і повтор...")
            time.sleep(wait_s)
            continue
        return resp
    if last_exc:
        raise last_exc
    return resp


def clean_file_path(path_str):
    """Очищає шлях, введений через drag-and-drop, від лапок/амперсанда/пробілів."""
    if path_str is None:
        return ""
    cleaned = str(path_str).strip().lstrip("&").strip()
    cleaned = cleaned.strip('"').strip("'").strip()
    return cleaned
