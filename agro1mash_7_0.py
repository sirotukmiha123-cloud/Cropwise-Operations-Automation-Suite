import json
import os
import re
import shutil
import pandas as pd
import requests
import time
import sys
from datetime import datetime, timedelta

try:
    from openpyxl.utils import column_index_from_string
except Exception:
    column_index_from_string = None

try:
    import colorama
    colorama.init(autoreset=True)
    C_OK    = colorama.Fore.GREEN
    C_ERR   = colorama.Fore.RED
    C_WARN  = colorama.Fore.YELLOW
    C_RESET = colorama.Style.RESET_ALL
except Exception:
    C_OK = C_ERR = C_WARN = C_RESET = ""

def colorize_line(line):
    """Підсвічує рядок звіту зеленим/червоним/жовтим залежно від його змісту (якщо colorama доступна)."""
    if "✅" in line or "Успішно" in line:
        return f"{C_OK}{line}{C_RESET}"
    if "❌" in line or "ПОМИЛКА" in line or "Помилк" in line:
        return f"{C_ERR}{line}{C_RESET}"
    if "⚠️" in line or "Пропущено" in line or "⏭️" in line:
        return f"{C_WARN}{line}{C_RESET}"
    return line

# =====================================================
# КОНФІГУРАЦІЯ
# =====================================================
API_BASE = "https://operations.cropwise.com"

CONFIG_EXAMPLE = '{\n  "cropwise_token": "Ваш токен тут"\n}'

def load_cropwise_token():
    """
    Порядок пошуку токена (від найбезпечнішого до резервного):
    1. Змінна середовища CROPWISE_TOKEN
    2. Файл config.json поруч зі скриптом (ключ "cropwise_token")

    Токен НІКОЛИ не зберігається в самому коді. Щоб почати роботу, або встановіть
    змінну середовища CROPWISE_TOKEN, або створіть файл config.json поруч зі скриптом:
    """ + "\n    " + CONFIG_EXAMPLE.replace("\n", "\n    ") + """
    Отримати токен: особистий кабінет Cropwise -> Налаштування -> API.
    ВАЖЛИВО: якщо викладаєте цей скрипт на GitHub чи ділитесь ним, додайте config.json
    у .gitignore, щоб не опублікувати свій токен випадково.
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
    token = str(CROPWISE_TOKEN).strip().replace("\n", "").replace("\r", "")
    return {
        "X-User-API-Token": token,
        "Accept": "application/json",
        "Content-Type": "application/json; charset=utf-8"
    }

def api_request(method, url, max_retries=3, **kwargs):
    """
    Обгортка над requests.request з обробкою HTTP 429 (Too Many Requests):
    чекає Retry-After (або зростаючу паузу) і повторює запит до max_retries разів.
    Також повторює спробу при мережевих таймаутах/збоях з'єднання.
    """
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

def show_loading_animation(endpoint_path, duration=0.4):
    frames = ["📡 [=   ]", "📡 [==  ]", "🛰️ [=== ]", "🚀 [====]"]
    for frame in frames:
        sys.stdout.write(f"\r  {frame} Обробка {endpoint_path}...")
        sys.stdout.flush()
        time.sleep(duration / len(frames))
    sys.stdout.write("\r" + " " * 60 + "\r")
    sys.stdout.flush()

def draw_progress_bar(current, total, width=20):
    if total <= 0: return "[░░░░░░░░░░░░░░░░░░░░] 0%"
    percent = int((current / total) * 100)
    filled_len = int(width * current // total)
    bar = "█" * filled_len + "░" * (width - filled_len)
    return f"[{bar}] {percent}%"

# =====================================================
# ЗАВАНТАЖЕННЯ ДОВІДНИКІВ З СЕРВЕРА
# =====================================================
def fetch_fields_from_server():
    print("🛰️  Завантаження актуальних полів з сервера Cropwise...")
    try:
        params = {"limit": 500}
        resp = api_request("GET", f"{API_BASE}/api/v3/fields", headers=get_headers(), params=params, timeout=15)
        if resp.status_code == 200:
            fields = resp.json().get("data", [])
            print(f"✅ Успішно завантажено {len(fields)} полів з сервера!")
            return fields
        return None
    except:
        return None

def fetch_work_types_from_server():
    print("🛰️  Завантаження актуальних підтипів робіт з сервера Cropwise...")
    try:
        resp = api_request("GET", f"{API_BASE}/api/v3/work_types", headers=get_headers(), timeout=15)
        if resp.status_code == 200:
            w_types = resp.json().get("data", [])
            print(f"✅ Успішно завантажено {len(w_types)} підтипів робіт з сервера!")
            return w_types
        return None
    except:
        return None

def fetch_drivers_from_server():
    """
    У Cropwise API v3 немає окремого ендпоінту /drivers — водії це Users з ознакою driver=true
    (див. Group User, operations.cropwise.com/api/v3/users). Офіційний список "Filters by" для
    User не включає 'driver', тож фільтруємо локально після завантаження.
    """
    print("🛰️  Завантаження водіїв (User, driver=true) з сервера Cropwise...")
    try:
        resp = api_request("GET", f"{API_BASE}/api/v3/users", headers=get_headers(), params={"limit": 500}, timeout=15)
        if resp.status_code == 200:
            data = resp.json().get("data", [])
            drivers_only = [u for u in data if u.get("driver") is True]
            print(f"✅ Успішно завантажено {len(drivers_only)} водіїв (із {len(data)} користувачів) з сервера!")
            return drivers_only
        return None
    except:
        return None

def fetch_machines_from_server():
    print("🛰️  Завантаження техніки з сервера Cropwise...")
    try:
        resp = api_request("GET", f"{API_BASE}/api/v3/machines", headers=get_headers(), params={"limit": 500}, timeout=15)
        if resp.status_code == 200:
            data = resp.json().get("data", [])
            print(f"✅ Успішно завантажено {len(data)} одиниць техніки з сервера!")
            return data
        return None
    except:
        return None

def create_new_work_type_on_server(new_name):
    print(f"\n⚡ [РЕЖИМ ДЕБАГУ 00] Створення нового підтипу робіт '{new_name}'")
    try:
        resp_groups = api_request("GET", f"{API_BASE}/api/v3/work_type_groups", headers=get_headers(), params={"limit": 100}, timeout=10)
        groups = resp_groups.json().get("data", []) if resp_groups.status_code == 200 else []
    except:
        groups = []

    if not groups:
        groups = [
            {"id": 2, "name": "Тваринництво (Ферма)"},
            {"id": 1, "name": "Роботи в рослинництві (Польові)"},
            {"id": 3, "name": "Транспортні / Допоміжні роботи"}
        ]

    while True:
        print("\n📁 ДОСТУПНІ ГРУПИ В CROPWISE:")
        for i, g in enumerate(groups):
            print(f"  {i+1:>2} → {g.get('name')} (ID: {g.get('id')})")
        print("  11 → Скасувати створення та пропустити")

        choice = input("\n👉 Виберіть номер групи: ").strip()
        if choice == "11": return None

        if choice.isdigit() and 1 <= int(choice) <= len(groups):
            selected_group = groups[int(choice) - 1]
            g_id = selected_group.get("id")
            break
        print("⚠️  Невірний вибір.")

    payload = {"data": {"name": str(new_name), "work_type_group_id": int(g_id)}}
    try:
        resp = api_request("POST", f"{API_BASE}/api/v3/work_types", headers=get_headers(), data=json.dumps(payload, ensure_ascii=False).encode("utf-8"), timeout=15)
        if resp.status_code in [200, 201]:
            new_item = resp.json().get("data", {})
            print(f"✨ Створено успішно! ID: {new_item.get('id')}")
            return new_item
        return None
    except:
        return None

def parse_flexible_time(time_range_str):
    try:
        if hasattr(time_range_str, 'month') and hasattr(time_range_str, 'day'):
            time_range_str = f"{time_range_str.day}-{time_range_str.month}"
        clean_range = str(time_range_str).replace(" ", "").strip()
        if not clean_range or clean_range.lower() in ['nan', '0', ''] or "-" not in clean_range:
            return None
        start_str, end_str = clean_range.split("-")
        if ":" not in start_str: start_str += ":00"
        if ":" not in end_str:   end_str += ":00"
        start_time = datetime.strptime(start_str, "%H:%M")
        end_time = datetime.strptime(end_str, "%H:%M")
        return f"{start_time.strftime('%H:%M')}-{end_time.strftime('%H:%M')}"
    except:
        return None

# =====================================================
# ДОВІДНИКИ (JSON-кеш + сервер)
# =====================================================
def load_json(filename):
    if os.path.exists(filename):
        with open(filename, 'r', encoding='utf-8') as f: return json.load(f)
    return []

drivers_data = fetch_drivers_from_server() or load_json("ref_drivers.json")
drivers_data = [d for d in drivers_data if str(d.get("username", "")).lower() not in ("support", "support (readonly)", "user")]
machines_data = fetch_machines_from_server() or load_json("ref_machines.json")
fields_data = fetch_fields_from_server() or load_json("ref_fields.json")
work_types_data = fetch_work_types_from_server() or load_json("ref_work_types.json")

def get_driver_name(d):
    """
    Реальна схема User (v3) НЕ має full_name/last_name/first_name — лише username, email тощо.
    """
    if not d: return "Невідомо"
    username = d.get("username")
    email = d.get("email")
    if username and email:
        return f"{username} ({email})"
    return username or email or f"User #{d.get('id', '?')}"

def get_machine_name(m):
    """
    Реальна схема Machine (v3): 'name' (readonly, авто = manufacturer+model, або custom_name),
    'registration_number' (не 'reg_number'!).
    """
    if not m: return "Невідомо"
    display_name = m.get("custom_name") or m.get("name") or f"{m.get('manufacturer','') or ''} {m.get('model','') or ''}".strip() or "Без назви"
    reg_num = m.get("registration_number") or ""
    return f"{display_name} {reg_num}".strip()

MONTHS_UA = {1: "Січень", 2: "Лютий", 3: "Березень", 4: "Квітень", 5: "Травень", 6: "Червень",
             7: "Липень", 8: "Серпень", 9: "Вересень", 10: "Жовтень", 11: "Листопад", 12: "Грудень"}

# =====================================================
# ДОПОМІЖНІ ФУНКЦІЇ: ШЛЯХИ, КЛЮЧІ, ЛОГ
# =====================================================
def clean_file_path(path_str):
    """Очищає шлях, введений через drag-and-drop, від лапок, амперсандів та зайвих пробілів."""
    if path_str is None:
        return ""
    cleaned = str(path_str).strip()
    cleaned = cleaned.lstrip("&").strip()
    cleaned = cleaned.strip('"').strip("'")
    cleaned = cleaned.strip()
    return cleaned

def clean_field_name_for_key(field_full_name):
    """
    Приводить повну назву поля до короткого унікального вигляду для ключа логу.
    Приклад: "00 Коледж(Іллінці)" -> "коледж"
    """
    if not field_full_name:
        return "не_вказано"
    s = str(field_full_name).strip()
    if s.lower() in ("nan", "", "не вказано"):
        return "не_вказано"
    s = re.sub(r"\(.*?\)", "", s)          # прибрати дужки з уточненням
    s = re.sub(r"^[\d\s№./-]+", "", s)     # прибрати провідний номер/код поля
    s = s.strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s or "не_вказано"

def normalize_for_fuzzy_match(s):
    """
    Прибирає пробіли, дефіси, слеші, крапки та інші розділювачі — щоб '10к/1' і '10 К/1'
    (чи '10-К/1', '10.К/1' тощо) вважались однаковим значенням при пошуку поля.
    """
    s = str(s).strip().lower()
    return re.sub(r"[^\w]", "", s, flags=re.UNICODE).replace("_", "")

def find_matching_field(query_str, fields_data):
    """
    Шукає поле за назвою, толерантно до пробілів/дефісів/регістру (напр. '10к/1' знайде
    '10 К/1 (якесь уточнення)'). Спершу пробує звичайний підрядковий пошук (швидший і
    точніший при повному збігу форматування), а якщо не знайшло — нормалізований fuzzy-пошук.
    """
    query_str = (query_str or "").strip().lower()
    if not query_str:
        return None

    # 1. Звичайний підрядковий пошук (як і раніше)
    direct_match = next((f for f in fields_data if query_str in str(f.get("name") or f.get("title", "")).lower()), None)
    if direct_match:
        return direct_match

    # 2. Резервний фаззі-пошук без урахування пробілів/дефісів/слешів/крапок
    query_norm = normalize_for_fuzzy_match(query_str)
    if not query_norm:
        return None
    return next((f for f in fields_data if query_norm in normalize_for_fuzzy_match(f.get("name") or f.get("title", ""))), None)

def row_key(date_val, work_name, field_short, excel_row_num):
    """
    Унікальний ключ рядка логу.
    ВАЖЛИВО: свідомо НЕ включає обсяг роботи (площу/години) — це значення може динамічно
    підмінюватись під час імпорту (напр. заглушка 0.1 га для нерозпізнаного поля), через що
    ключі в Режимі 1 (створення) і Режимі 2 (пальне) могли б розійтися. Натомість
    використовується номер фізичного рядка в Excel-файлі — він стабільний, доки Режим 2
    запускається на тому самому (чи структурно ідентичному) файлі, що й Режим 1.
    """
    return f"{date_val}|{work_name}|{field_short}|row{excel_row_num}"

def load_log(log_path):
    if os.path.exists(log_path):
        with open(log_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_log(log_path, log_data):
    with open(log_path, 'w', encoding='utf-8') as f:
        json.dump(log_data, f, ensure_ascii=False, indent=2)

def generate_log_path_for_excel(excel_path):
    """Генерує шлях до JSON-логу поруч з Excel-файлом: Vodinnia_06.xlsx -> Vodinnia_06_LOG.json"""
    base_dir = os.path.dirname(os.path.abspath(excel_path))
    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    return os.path.join(base_dir, f"{base_name}_LOG.json")

def sanitize_folder_name(name):
    """Прибирає символи, заборонені у назвах папок/файлів Windows, і замінює пробіли на підкреслення."""
    name = str(name).strip()
    name = re.sub(r'[\\/:*?"<>|]', '', name)
    name = re.sub(r'\s+', '_', name)
    return name.strip('_')

def build_session_folder_name(driver_name_str, machine_name_str, session_month, session_year):
    """
    Формує назву папки сесії за шаблоном 'Водій_Машина_Місяць_Рік',
    напр. 'Іваненко_Петро_Трактор_МТЗ_82_АА1234ВВ_Червень_2026'.
    """
    driver_clean = sanitize_folder_name(driver_name_str.split('(')[0])  # прибираємо email у дужках
    machine_clean = sanitize_folder_name(machine_name_str)
    month_name = MONTHS_UA.get(session_month, str(session_month))
    return f"{driver_clean}_{machine_clean}_{month_name}_{session_year}"

def ensure_session_folder(excel_file, driver_name_str, machine_name_str, session_month, session_year):
    """
    Створює (якщо потрібно) окрему папку сесії поруч з Excel-файлом і переносить туди сам
    Excel-файл, якщо він ще не в ній. Повертає новий шлях до Excel-файлу (усередині папки сесії).
    Лог і фінальні звіти автоматично підуть у ту саму папку, бо їхні шляхи будуються від шляху
    до Excel (generate_log_path_for_excel / os.path.dirname(log_path)).
    """
    folder_name = build_session_folder_name(driver_name_str, machine_name_str, session_month, session_year)
    base_dir = os.path.dirname(os.path.abspath(excel_file))
    current_folder_name = os.path.basename(base_dir)

    # Якщо Excel вже лежить у папці з такою самою назвою сесії — нічого переносити не треба.
    if current_folder_name == folder_name:
        return excel_file

    session_folder = os.path.join(base_dir, folder_name)
    try:
        os.makedirs(session_folder, exist_ok=True)
    except Exception as e:
        print(f"⚠️  Не вдалося створити папку сесії ({e}), працюю з поточним розташуванням файлу.")
        return excel_file

    new_excel_path = os.path.join(session_folder, os.path.basename(excel_file))
    if os.path.abspath(excel_file) == os.path.abspath(new_excel_path):
        return excel_file

    try:
        if os.path.exists(new_excel_path):
            print(f"ℹ️  У папці сесії вже є файл з такою назвою — працюю з ним: {new_excel_path}")
        else:
            shutil.move(excel_file, new_excel_path)
            print(f"📦 Excel-файл перенесено в папку сесії:\n   📁 {session_folder}")
        return new_excel_path
    except Exception as e:
        print(f"⚠️  Не вдалося перенести Excel у папку сесії ({e}), працюю з поточним розташуванням файлу.")
        return excel_file

# =====================================================
# СПІЛЬНА ЛОГІКА РОЗБОРУ РЯДКА EXCEL (щоб ключі в Режимі 1 і 2 завжди збігались)
# =====================================================
STOP_WORDS = ["назва культури", "виконаної робот", "відпрацьовано", "разом", "підпис", "перевірив", "бригадир", "бухгалтер"]

def is_valid_work_row(row):
    day_val = str(row[2]).strip() if pd.notna(row[2]) else "nan"
    work_name = str(row[5]).strip() if pd.notna(row[5]) else "nan"

    if day_val == "nan" or work_name == "nan":
        return False

    for word in STOP_WORDS:
        if word in day_val.lower() or word in work_name.lower():
            return False

    # Якщо в комірці діапазон дат (напр. "17.06.26-18.06.26" — робота на два дні, часом
    # записана навіть у два рядки з переносом), для перевірки валідності беремо перший день.
    day_first = day_val.split("-")[0].strip()
    clean_day = day_first.split(".")[0].replace(",", ".").strip()
    try:
        day_num = int(float(clean_day))
        if 1 <= day_num <= 31:
            return True
    except ValueError:
        pass
    return False

def build_date_val(day_val, session_month, session_year):
    """Розбирає ОДНЕ значення дня (без діапазону, дефіс уже прибраний викликачем) у дд.мм.рррр."""
    day_val = str(day_val).strip()
    if "." in day_val:
        date_val = day_val.replace(".26", ".2026")
        # якщо в рядку вже стоїть повна дата "дд.мм.рррр" - лишаємо як є, інакше добудовуємо
        parts = date_val.split(".")
        if len(parts) == 2:
            date_val = f"{parts[0]}.{session_month:02d}.{session_year}"
        elif len(parts) == 3 and len(parts[2]) <= 2:
            date_val = f"{parts[0]}.{parts[1]}.{session_year}"
    else:
        try:
            date_val = f"{int(float(day_val)):02d}.{session_month:02d}.{session_year}"
        except:
            date_val = f"{day_val}.{session_month:02d}.{session_year}"
    return date_val

def build_date_range(day_val, session_month, session_year):
    """
    Розбирає значення дня, повертаючи (дата_початку, дата_кінця).
    Якщо в комірці діапазон дат (напр. "17.06.26-18.06.26", можливо з переносом рядка
    "17.06.26-\n18.06.26" — робота на два дні), повертає ОБИДВІ дати: перша йде як дата
    операції/ключ логу, а друга — як дата завершення завдання (start_time/end_time тоді
    коректно охоплюють обидва календарних дні замість штучного стиснення в один).
    Для звичайного одиничного дня обидва значення однакові.
    """
    raw = str(day_val).strip()
    if "-" in raw:
        parts = [p.strip() for p in raw.split("-") if p.strip()]
        start_part = parts[0] if parts else raw
        end_part = parts[1] if len(parts) > 1 else start_part
        start_date = build_date_val(start_part, session_month, session_year)
        end_date = build_date_val(end_part, session_month, session_year)
        return start_date, end_date
    single_date = build_date_val(raw, session_month, session_year)
    return single_date, single_date

def add_one_day_to_date_str(date_str):
    """Додає один календарний день до дати у форматі дд.мм.рррр (для нічних змін, що переходять за північ)."""
    dt = datetime.strptime(date_str, "%d.%m.%Y")
    dt_next = dt + timedelta(days=1)
    return dt_next.strftime("%d.%m.%Y")

def extract_row_essentials(row, session_month, session_year, excel_row_num):
    """
    Витягує з рядка Excel всі дані, потрібні для побудови унікального ключа логу:
    дату, назву роботи, коротку назву поля та фактичний обсяг (площу/години).
    Використовується ОДНАКОВО в Режимі 1 (створення) і Режимі 2 (пальне),
    щоб ключі гарантовано збігались. Ключ будується на основі номера рядка Excel
    (excel_row_num), а не обсягу роботи — див. коментар у row_key().
    """
    day_val = str(row[2]).strip() if pd.notna(row[2]) else ""
    field_num_xl = str(row[4]).strip() if pd.notna(row[4]) else "Не вказано"
    work_name = str(row[5]).strip() if pd.notna(row[5]) else ""
    unit_val = str(row[13]).strip() if pd.notna(row[13]) else "га"

    date_val, end_date_val = build_date_range(day_val, session_month, session_year)

    try:
        actual_work_amount = float(str(row[22]).replace(",", "."))
    except:
        actual_work_amount = 0.0

    excel_time_val = str(row[15]).strip() if pd.notna(row[15]) else ""
    parsed_range = parse_flexible_time(excel_time_val)
    if parsed_range:
        try:
            sh_str, eh_str = parsed_range.split("-")
            actual_hours = float(int(eh_str.split(":")[0]) - int(sh_str.split(":")[0]))
        except:
            actual_hours = 1.0
    else:
        try:
            actual_hours = float(excel_time_val.replace(",", ".")) if excel_time_val else 1.0
        except:
            actual_hours = 1.0

    if actual_work_amount == 0.0 and actual_hours > 0:
        actual_work_amount = actual_hours

    field_short = clean_field_name_for_key(field_num_xl)
    key = row_key(date_val, work_name, field_short, excel_row_num)

    return {
        "key": key,
        "date_val": date_val,
        "end_date_val": end_date_val,
        "work_name": work_name,
        "field_num_xl": field_num_xl,
        "field_short": field_short,
        "unit_val": unit_val,
        "actual_work_amount": actual_work_amount,
        "actual_hours": actual_hours,
        "excel_time_val": excel_time_val,
        "excel_row_num": excel_row_num,
    }

# =====================================================
# СЕРВЕРНІ ПЕРЕВІРКИ ДУБЛІКАТІВ
# =====================================================
def fetch_existing_tasks(date_iso, driver_id, machine_id, work_type_id):
    params = {"start_time_gt_eq": f"{date_iso}T00:00:00", "start_time_lt_eq": f"{date_iso}T23:59:59", "limit": 100}
    if driver_id:    params["driver_id"] = driver_id
    if machine_id:   params["machine_id"] = machine_id
    if work_type_id: params["work_type_id"] = work_type_id
    try:
        resp = api_request("GET", f"{API_BASE}/api/v3/machine_tasks", headers=get_headers(), params=params, timeout=10)
        return resp.json().get("data", []) if resp.status_code == 200 else []
    except:
        return []

def fetch_field_for_task(task_id):
    try:
        resp = api_request("GET", f"{API_BASE}/api/v3/machine_task_field_mapping_items", headers=get_headers(), params={"machine_task_id": task_id, "limit": 10}, timeout=8)
        return [item.get("field_id") for item in resp.json().get("data", [])] if resp.status_code == 200 else []
    except:
        return []

def fetch_agro_for_task(task_id):
    try:
        resp = api_request("GET", f"{API_BASE}/api/v3/machine_task_agro_operation_mapping_items", headers=get_headers(), params={"machine_task_id": task_id, "limit": 5}, timeout=8)
        return len(resp.json().get("data", [])) > 0 if resp.status_code == 200 else False
    except:
        return False

def check_duplicate(date_iso, driver_id, machine_id, field_id, work_type_id, work_name):
    tasks = fetch_existing_tasks(date_iso, driver_id, machine_id, work_type_id)
    if not tasks: return []
    duplicates = []
    for t in tasks:
        t_id = t.get("id")
        if field_id and str(field_id) != "0":
            if str(field_id) not in [str(x) for x in fetch_field_for_task(t_id)]: continue
        t_desc = (t.get("description") or "").strip()
        if work_name and t_desc and work_name.lower() not in t_desc.lower(): continue
        duplicates.append({"id": t_id, "has_agro": fetch_agro_for_task(t_id), "description": t_desc})
    return duplicates

def delete_task(task_id):
    try:
        return api_request("DELETE", f"{API_BASE}/api/v3/machine_tasks/{task_id}", headers=get_headers(), timeout=10).status_code in [200, 204]
    except:
        return False

# =====================================================
# ВІДПРАВКА РОБОТИ + АГРООПЕРАЦІЇ В CROPWISE
# =====================================================
def get_kyiv_utc_offset(dt_naive):
    """
    Повертає зміщення часу Києва відносно UTC у форматі "+0300"/"+0200" для заданої наївної дати.
    Спочатку намагається використати системну базу IANA через zoneinfo (коректно враховує
    перехід на літній/зимовий час). Якщо zoneinfo/tzdata недоступні (типово на "чистому" Windows
    без пакета tzdata) — використовує резервний розрахунок за правилом ЄС (яке Україна наслідує):
    літній час з останньої неділі березня по останню неділю жовтня.
    """
    try:
        from zoneinfo import ZoneInfo
        aware = dt_naive.replace(tzinfo=ZoneInfo("Europe/Kyiv"))
        return aware.strftime("%z")
    except Exception:
        return _eu_dst_offset_fallback(dt_naive)

def _eu_dst_offset_fallback(dt_naive):
    year = dt_naive.year

    def last_sunday(month):
        if month == 12:
            next_month_first = datetime(year + 1, 1, 1)
        else:
            next_month_first = datetime(year, month + 1, 1)
        last_day = next_month_first - timedelta(days=1)
        days_back = (last_day.weekday() - 6) % 7  # weekday(): Monday=0 ... Sunday=6
        return last_day - timedelta(days=days_back)

    dst_start = last_sunday(3).replace(hour=1, minute=0, second=0, microsecond=0)   # 01:00 UTC, остання неділя березня
    dst_end = last_sunday(10).replace(hour=1, minute=0, second=0, microsecond=0)    # 01:00 UTC, остання неділя жовтня

    if dst_start <= dt_naive < dst_end:
        return "+0300"
    return "+0200"

def format_iso_with_kyiv_offset(date_str, time_str_hms):
    """Формує ISO-рядок з коректним (сезонним) зміщенням Києва, напр. 2026-06-04T08:00:00.000+03:00"""
    dt_naive = datetime.strptime(f"{date_str} {time_str_hms}", "%Y-%m-%d %H:%M:%S")
    offset_raw = get_kyiv_utc_offset(dt_naive)         # напр. "+0300"
    offset_fmt = f"{offset_raw[:3]}:{offset_raw[3:]}"  # -> "+03:00"
    return f"{date_str}T{time_str_hms}.000{offset_fmt}"

# Емпірично підтверджено сервером (422-помилка): технічне поле-заглушка "00 Інші не польові
# операції (н.п.)" (ID 24) приймає площу не більше 0.1 га, НЕЗАЛЕЖНО від того, що показує наш
# власний кеш fields_data (динамічний пошук tillable_area для цього поля виявився ненадійним —
# можливо, поле в кеші має tillable_area=0 чи не заповнене, тоді як реальне обмеження на сервері
# інше). Тому для відомих технічних полів застосовуємо гарантований резервний ліміт.
KNOWN_STUB_FIELD_MAX_AREA = {24: 0.1, 25: 0.1}

def get_field_tillable_area(field_id):
    """
    Шукає реальну оброблювану площу (tillable_area) поля. Спершу перевіряє відомі технічні
    поля-заглушки (гарантований ліміт, підтверджений сервером), інакше шукає в уже завантаженому
    довіднику fields_data. Потрібно, щоб коректно обмежувати planned_area/completed_area
    агрооперації для маленьких технічних об'єктів (сміттєзвалище, бурт для гною тощо) — сервер
    відхиляє площу, що перевищує tillable_area поля.
    """
    if field_id in KNOWN_STUB_FIELD_MAX_AREA:
        return KNOWN_STUB_FIELD_MAX_AREA[field_id]

    for f in fields_data:
        if str(f.get("id")) == str(field_id):
            area = f.get("tillable_area")
            if area in (None, ""):
                area = f.get("calculated_area")
            try:
                return float(area) if area not in (None, "") else None
            except (TypeError, ValueError):
                return None
    return None

def send_to_cropwise(driver_id, machine_id, work_type_id, field_id, iso_start, iso_end, season_year, send_amount, work_name, description):
    agro_op_id = None
    if field_id and str(field_id) != "0":
        start_date = iso_start.split("T")[0]
        end_date = iso_end.split("T")[0]
        raw_start_time = iso_start.split("T")[1]
        raw_end_time = iso_end.split("T")[1]

        s_time = raw_start_time if len(raw_start_time.split(':')) == 3 else f"{raw_start_time}:00"
        e_time = raw_end_time if len(raw_end_time.split(':')) == 3 else f"{raw_end_time}:00"

        actual_start_dt = format_iso_with_kyiv_offset(start_date, s_time)
        completed_dt = format_iso_with_kyiv_offset(end_date, e_time)

        # Обмежуємо площу агрооперації реальною оброблюваною площею поля (якщо вона менша за
        # надіслану) — сервер інакше відхиляє запит для маленьких технічних об'єктів
        # (сміттєзвалище, бурт для гною тощо), а не лише для заглушки ID 24/25.
        tillable_area = get_field_tillable_area(field_id)
        if tillable_area is not None and tillable_area > 0 and float(send_amount) > tillable_area:
            agro_area = tillable_area
        else:
            agro_area = float(send_amount)

        # ПРИМІТКА: 'custom_name' — валідне поле лише для /api/v3a/agro_operations, у v3 його немає.
        # Назву роботи й так передаємо через 'description'.
        agro_payload = {
            "data": {
                "field_id": int(field_id),
                "work_type_id": int(work_type_id) if str(work_type_id).isdigit() else None,
                "season": season_year,
                "status": "done",
                "description": f"Операція з облікового листа: {work_name}",
                "planned_area": agro_area,
                "completed_area": agro_area,
                "planned_start_date": start_date,
                "planned_end_date": end_date,
                "completed_date": end_date,
                "actual_start_datetime": str(actual_start_dt),
                "completed_datetime": str(completed_dt)
            }
        }
        try:
            print(f"\n🔍 [ДЕБАГ АГРООП] Спроба відправки для: '{work_name}'")
            print(f"   📊 Параметри: field_id={field_id}, work_type_id={work_type_id}, season={season_year}")

            show_loading_animation("/api/v3/agro_operations", duration=0.2)

            url_agro = f"{API_BASE}/api/v3/agro_operations"
            r = api_request("POST", url_agro, headers=get_headers(), data=json.dumps(agro_payload, ensure_ascii=False).encode("utf-8"), timeout=15)

            print(f"   📡 Статус відповіді сервера: {r.status_code}")

            if r.status_code in [200, 201]:
                agro_op_id = r.json().get("data", {}).get("id")
                print(f"   ✨ Успішно створено AgroOp! ID: {agro_op_id}")
            else:
                print(f"   ❌ Сервер Cropwise відхилив АгроОп: {r.text}")
        except Exception as agro_err:
            print(f"   💥 Критична помилка відправки АгроОп: {agro_err}")

    # ПРИМІТКА: MachineTask (v3) НЕ має власних атрибутів 'field_id' та 'status' — зв'язок з полем
    # робиться через окремий MachineTaskFieldMappingItem (нижче), а обсяг роботи зберігається як
    # 'work_area' (в оригінальному коді помилково називалось 'work_amount').
    payload = {"data": {
        "driver_id": int(driver_id) if str(driver_id).isdigit() else None,
        "machine_id": int(machine_id) if str(machine_id).isdigit() else None,
        "work_type_id": int(work_type_id) if str(work_type_id).isdigit() else None,
        "work_area": float(send_amount),
        "description": str(description),
        "season": season_year,
        "start_time": iso_start,
        "end_time": iso_end
    }}
    try:
        show_loading_animation("/api/v3/machine_tasks", duration=0.3)
        r = api_request("POST", f"{API_BASE}/api/v3/machine_tasks", headers=get_headers(), data=json.dumps(payload, ensure_ascii=False).encode("utf-8"), timeout=15)
        if r.status_code not in [200, 201]: return None, None, None
        mt_id = r.json().get("data", {}).get("id")
        field_mapping_item_id = None
        if mt_id and field_id and str(field_id) != "0":
            # MachineTaskFieldMappingItem: поле обсягу роботи, яке МОЖНА редагувати, зветься
            # 'covered_area' (не 'work_amount' — того поля в схемі взагалі немає).
            # ВАЖЛИВО: 'covered_area' САМ ПО СОБІ ігнорується сервером — Cropwise за замовчуванням
            # розраховує оброблену площу з GPS/телеметрії треку машини. Оскільки в нас немає
            # реального треку (дані створюються через API, не з бортового обладнання), система
            # показує порожнє/дефолтне значення (площу самого поля), а наше covered_area мовчки
            # ігнорується — САМЕ ТАК і виявилась ця проблема (площа "не відображалась" в UI).
            # За документацією API, covered_area починає діяти лише коли manually_defined_covered_area=true.
            fm = {"data": {
                "machine_task_id": int(mt_id),
                "field_id": int(field_id),
                "covered_area": float(send_amount),
                "manually_defined_covered_area": True,
            }}
            fm_resp = api_request("POST", f"{API_BASE}/api/v3/machine_task_field_mapping_items", headers=get_headers(), data=json.dumps(fm, ensure_ascii=False).encode("utf-8"), timeout=10)
            if fm_resp.status_code in [200, 201]:
                field_mapping_item_id = fm_resp.json().get("data", {}).get("id")
            if agro_op_id:
                am = {"data": {"machine_task_id": int(mt_id), "agro_operation_id": int(agro_op_id)}}
                api_request("POST", f"{API_BASE}/api/v3/machine_task_agro_operation_mapping_items", headers=get_headers(), data=json.dumps(am, ensure_ascii=False).encode("utf-8"), timeout=10)
        return mt_id, agro_op_id, field_mapping_item_id
    except:
        return None, None, None

def get_field_mapping_item_id(machine_task_id, field_id=None):
    """
    Резервний пошук ID MachineTaskFieldMappingItem за machine_task_id — потрібен для старих
    записів логу (v7/v8), створених ДО того, як цей ID почав зберігатися в лозі напряму.
    """
    try:
        params = {"machine_task_id": machine_task_id, "limit": 5}
        if field_id:
            params["field_id"] = field_id
        resp = api_request("GET", f"{API_BASE}/api/v3/machine_task_field_mapping_items", headers=get_headers(), params=params, timeout=10)
        if resp.status_code == 200:
            items = resp.json().get("data", [])
            if items:
                return items[0].get("id")
        return None
    except:
        return None

def get_field_id_for_task(machine_task_id):
    """
    Резервний пошук field_id за machine_task_id — потрібен для старих записів логу (v7),
    створених ДО того, як field_id почав зберігатися в лозі напряму. Без цього такі записи
    в Режимі 3 помилково падали на консumption_target='Machine' замість коректного 'Field'.
    """
    try:
        resp = api_request("GET", f"{API_BASE}/api/v3/machine_task_field_mapping_items", headers=get_headers(),
                            params={"machine_task_id": machine_task_id, "limit": 5}, timeout=10)
        if resp.status_code == 200:
            items = resp.json().get("data", [])
            if items:
                return items[0].get("field_id")
        return None
    except:
        return None

def get_machine_task_end_time(machine_task_id):
    """
    Резервний пошук end_time MachineTask за його ID — потрібен для старих записів логу,
    створених ДО того, як end_time почав зберігатися в лозі напряму (Режим 1). Без цього
    Режим 3 ставив би у списанні пального фіксовану полудень-годину замість реального часу
    завершення роботи.
    """
    try:
        resp = api_request("GET", f"{API_BASE}/api/v3/machine_tasks/{machine_task_id}", headers=get_headers(), timeout=10)
        if resp.status_code == 200:
            return resp.json().get("data", {}).get("end_time")
        return None
    except:
        return None

def patch_fuel_used(machine_task_id, fuel_liters, field_id=None, field_mapping_item_id=None):
    """
    ВАЖЛИВО: у реальній схемі API немає поля 'fuel_used' на MachineTask!
    Витрата пального ('fuel_consumption') зберігається на MachineTaskFieldMappingItem і є
    readonly, ЯКЩО не встановити 'manually_defined_fuel_consumption': true в тому самому запиті —
    саме так документація описує ручне перевизначення цього значення.
    """
    mapping_id = field_mapping_item_id or get_field_mapping_item_id(machine_task_id, field_id)
    if not mapping_id:
        return False, None

    payload = {"data": {
        "fuel_consumption": float(fuel_liters),
        "manually_defined_fuel_consumption": True
    }}
    try:
        r = api_request("PATCH", f"{API_BASE}/api/v3/machine_task_field_mapping_items/{mapping_id}", headers=get_headers(),
                            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"), timeout=15)
        return (r.status_code == 200), mapping_id
    except:
        return False, mapping_id

# =====================================================
# ІНІЦІАЛІЗАЦІЯ СЕСІЇ (спільна для обох режимів)
# =====================================================
def init_session():
    if not drivers_data:
        print(f"{C_ERR}❌ Довідник водіїв порожній: не вдалося завантажити дані ні з сервера, ні з кешу "
              f"ref_drivers.json. Перевірте з'єднання з інтернетом/токен і спробуйте знову.{C_RESET}")
        sys.exit(1)
    if not machines_data:
        print(f"{C_ERR}❌ Довідник техніки порожній: не вдалося завантажити дані ні з сервера, ні з кешу "
              f"ref_machines.json. Перевірте з'єднання з інтернетом/токен і спробуйте знову.{C_RESET}")
        sys.exit(1)

    print("\n📅 Місяць та рік облікового листа:")
    while True:
        m_in = input("  Місяць (1-12): ").strip()
        y_in = input("  Рік (наприклад 2026): ").strip()
        if m_in.isdigit() and y_in.isdigit() and 1 <= int(m_in) <= 12:
            session_month = int(m_in)
            session_year = int(y_in)
            break
        print("  ⚠️  Невірний формат.")

    print("\n👨 ВОДІЇ:")
    for i, d in enumerate(drivers_data): print(f"  {i+1:>3}. {get_driver_name(d)}")
    while True:
        ch = input("  Номер водія (або частина прізвища, 'вихід' для виходу): ").strip()
        if ch.lower() in ("вихід", "exit", "quit"):
            print("👋 Роботу перервано користувачем.")
            sys.exit(0)
        if ch.isdigit() and 1 <= int(ch) <= len(drivers_data):
            selected_driver = drivers_data[int(ch) - 1]
            break
        matches = [d for d in drivers_data if ch.lower() in get_driver_name(d).lower()]
        if len(matches) == 1:
            selected_driver = matches[0]
            break
        elif len(matches) > 1:
            print("  Знайдено декілька збігів, уточніть номер зі списку вище.")
        else:
            print("  ⚠️  Не знайдено. Спробуйте ще раз.")

    print("\n🚜 ТЕХНІКА:")
    for i, m in enumerate(machines_data): print(f"  {i+1:>3}. {get_machine_name(m)}")
    while True:
        ch = input("  Номер машини (або частина назви, 'вихід' для виходу): ").strip()
        if ch.lower() in ("вихід", "exit", "quit"):
            print("👋 Роботу перервано користувачем.")
            sys.exit(0)
        if ch.isdigit() and 1 <= int(ch) <= len(machines_data):
            selected_machine = machines_data[int(ch) - 1]
            break
        matches = [m for m in machines_data if ch.lower() in get_machine_name(m).lower()]
        if len(matches) == 1:
            selected_machine = matches[0]
            break
        elif len(matches) > 1:
            print("  Знайдено декілька збігів, уточніть номер зі списку вище.")
        else:
            print("  ⚠️  Не знайдено. Спробуйте ще раз.")

    driver_id, machine_id = selected_driver.get("id"), selected_machine.get("id")
    driver_name_str, machine_name_str = get_driver_name(selected_driver), get_machine_name(selected_machine)

    return {
        "month": session_month,
        "year": session_year,
        "driver_id": driver_id,
        "machine_id": machine_id,
        "driver_name": driver_name_str,
        "machine_name": machine_name_str,
    }

# =====================================================
# РЕЖИМ 1: ПЕРВИННИЙ ІМПОРТ РОБІТ + СТВОРЕННЯ ЛОГУ
# =====================================================
def run_mode_primary_import(session):
    session_month, session_year = session["month"], session["year"]
    driver_id, machine_id = session["driver_id"], session["machine_id"]
    driver_name_str, machine_name_str = session["driver_name"], session["machine_name"]

    raw_path = input("\n📂 Шлях до Excel-файлу (можна перетягнути файл у консоль): ")
    excel_file = clean_file_path(raw_path)
    if not os.path.exists(excel_file):
        print("❌ Файл не знайдено!")
        return

    excel_file = ensure_session_folder(excel_file, driver_name_str, machine_name_str, session_month, session_year)

    log_path = generate_log_path_for_excel(excel_file)
    log_data = load_log(log_path)
    print(f"📁 Лог сесії: {log_path}")

    try:
        df = pd.read_excel(excel_file, header=None, engine='openpyxl')
    except Exception as e:
        print(f"❌ Помилка Excel: {e}")
        return

    valid_rows = [(index, row) for index, row in df.iterrows() if is_valid_work_row(row)]

    if not valid_rows:
        print("⚠️  У файлі не знайдено жодного валідного робочого рядка після фільтрації!")
        return

    print(f"\n✅ Знайдено {len(valid_rows)} реальних робочих операцій.")
    print("⏩ З якої операції починати?")
    print("  0     → з самого початку")
    print("  n     → з порядкового номера операції")
    start_ch = input("  → ").strip()
    start_index = int(start_ch) - 1 if (start_ch.isdigit() and int(start_ch) > 0) else 0

    stats = {"sent": 0, "skipped": 0, "error": 0, "total_hours": 0.0}
    final_registry = []
    day_time_memory = {}
    total_valid_count = len(valid_rows)

    print("\n🤖 АВТОПІЛОТ ЗАПУЩЕНО... Обробка йде у фоні.")

    work_row_counter = 0

    for current_pos in range(start_index, total_valid_count):
        original_index, row = valid_rows[current_pos]
        excel_row_num = original_index + 1

        essentials = extract_row_essentials(row, session_month, session_year, excel_row_num)
        r_key = essentials["key"]
        date_val = essentials["date_val"]
        end_date_val = essentials["end_date_val"]
        work_name = essentials["work_name"]
        field_num_xl = essentials["field_num_xl"]
        unit_val = essentials["unit_val"]
        actual_work_amount = essentials["actual_work_amount"]
        actual_hours = essentials["actual_hours"]
        excel_time_val = essentials["excel_time_val"]

        work_row_counter += 1

        parsed_range = parse_flexible_time(excel_time_val)
        if parsed_range:
            time_formatted = parsed_range
            end_h, end_m = map(int, parsed_range.split("-")[1].split(":"))
            day_time_memory[date_val] = end_h * 60 + end_m
        else:
            # ВАЖЛИВО: раніше тут було `end_hour = int(start_hour + actual_hours)`, що ОБРІЗАЛО
            # дробову частину годин замість округлення. Для будь-якої операції коротшої за годину
            # (напр. 0.28 год = 17 хв) це давало end_hour == start_hour (той самий час!), що потім
            # хибно розпізнавалось як "перехід через північ" і додавало ЦІЛУ добу до дати
            # завершення. Тепер рахуємо в хвилинах з округленням — короткі операції (хвилини)
            # більше не перетворюються помилково на добові.
            start_total_min = day_time_memory.get(date_val, 8 * 60)
            duration_min = round(actual_hours * 60)
            end_total_min = start_total_min + duration_min

            start_h, start_m = divmod(start_total_min % (24 * 60), 60)
            end_h, end_m = divmod(end_total_min % (24 * 60), 60)

            time_formatted = f"{start_h:02d}:{start_m:02d}-{end_h:02d}:{end_m:02d}"
            # Зберігаємо БЕЗ обрізання по модулю 24 год — щоб ланцюжок послідовних операцій
            # того самого дня рахувався коректно, навіть якщо сумарно вони переходять за північ.
            day_time_memory[date_val] = end_total_min

        if log_data.get(r_key, {}).get("status") == "sent":
            stats["skipped"] += 1
            saved_mt = log_data[r_key].get("machine_task_id", "ПРОПУЩЕНО")
            saved_ao = log_data[r_key].get("agro_operation_id", "ПРОПУЩЕНО")
            final_registry.append([work_row_counter, work_name, f"MT-{saved_mt}", f"AO-{saved_ao}" if saved_ao else "-", time_formatted, f"{actual_work_amount} {unit_val}", "Пропущено (Дубль)"])
            continue

        prog_bar = draw_progress_bar(current_pos + 1, total_valid_count, width=20)
        print(f"⚙️  [Робоча операція №{work_row_counter}] {prog_bar} | Сплавляю: '{work_name[:30]}...'")

        work_type_id = None
        for wt in work_types_data:
            if str(wt.get("name", "")).strip().lower() == work_name.strip().lower():
                work_type_id = wt.get("id")
                break

        if not work_type_id:
            print(f"\n⚠️  [РЕЖИМ ДЕБАГУ 00] Роботи '{work_name}' немає на сервері.")
            new_wt = create_new_work_type_on_server(work_name)
            if new_wt:
                work_type_id = new_wt.get("id")
                work_types_data.append(new_wt)
            else:
                final_registry.append([work_row_counter, work_name, "-", "-", time_formatted, f"{actual_work_amount} {unit_val}", "Пропущено (Немає підтипу)"])
                continue

        stats["total_hours"] += actual_hours

        excel_field_val = field_num_xl.strip().lower()
        field_id = None
        field_name = "Тех. поле (Заглушка 0.1 га)"
        send_amount = actual_work_amount  # ЗАВЖДИ реальна оброблена площа/обсяг — ніколи не обрізаємо тут.
                                           # Обрізання (якщо потрібне) відбувається лише для AgroOperation
                                           # нижче, через get_field_tillable_area + анотацію в назві.

        if excel_field_val and excel_field_val != "не вказано" and excel_field_val != "nan":
            matched_field = find_matching_field(excel_field_val, fields_data)
            if matched_field:
                field_id = matched_field.get("id")
                field_name = matched_field.get("name") or matched_field.get("title")

        if not field_id and excel_field_val and excel_field_val != "не вказано" and excel_field_val != "nan":
            print(f"\n🗺️  [РЕЖИМ ДЕБАГУ 00] Поле '{field_num_xl}' не розпізнано автоматично.")
            print("  0  → Поставити заглушку (ID: 24, 0.1 га)")
            field_input = input("  Введіть назву поля або 0: ").strip().lower()
            if field_input == "0" or not field_input:
                field_id, field_name = 24, "Тех. поле (Заглушка 0.1 га)"
            else:
                manual_match = find_matching_field(field_input, fields_data)
                if manual_match:
                    field_id = manual_match.get("id")
                    field_name = manual_match.get("name") or manual_match.get("title")
                else:
                    field_id, field_name = 24, "Тех. поле (Заглушка 0.1 га)"
        elif not field_id:
            field_id = 24

        # Якщо реальний обсяг перевищує оброблювану площу поля (типово для технічних полів-заглушок,
        # де tillable_area лише 0.1 га, але в обліковому листі пишуть скільки РЕАЛЬНО оброблено) —
        # AgroOperation однаково прийме лише обрізане значення (нижче, в send_to_cropwise), тож тут
        # додаємо позначку з фактичним обсягом прямо в назву/опис, щоб ця інформація не губилась.
        display_work_name = work_name
        field_tillable_area = get_field_tillable_area(field_id)
        if field_tillable_area is not None and field_tillable_area > 0 and actual_work_amount > field_tillable_area:
            amount_str = f"{actual_work_amount:g}"
            unit_label = (unit_val or "од.").strip()
            display_work_name = f"{work_name} (фактично виконано: {amount_str} {unit_label})"

        description = (
            f"Операція №{work_row_counter} "
            f"(рядок Excel {excel_row_num}, "
            f"фактично виконано {actual_work_amount:g} {(unit_val or 'од.').strip()})"
            )
        try:
            d, mo, y = date_val.split(".")
            date_iso = f"{y}-{mo}-{d}"
            sh, eh = time_formatted.split("-")
            sh_h, sh_m = map(int, sh.strip().split(":"))
            eh_h, eh_m = map(int, eh.strip().split(":"))

            end_date_for_iso = end_date_val
            # Якщо час "перегорнувся" через північ (напр. 16:00-00:00 чи 22:00-02:00) і при цьому
            # рядок Excel НЕ задавав власного багатоденного діапазону (end_date_val == date_val),
            # то це наш розрахований перехід через північ — переносимо кінцеву дату на наступний
            # день, інакше iso_end виходить РАНІШЕ за iso_start і сервер відхиляє запит.
            if (eh_h * 60 + eh_m) <= (sh_h * 60 + sh_m) and end_date_val == date_val:
                end_date_for_iso = add_one_day_to_date_str(date_val)

            ed, emo, ey = end_date_for_iso.split(".")
            iso_start = f"{y}-{mo}-{d}T{sh.strip()}:00"
            iso_end = f"{ey}-{emo}-{ed}T{eh.strip()}:00"
            season_year = int(y)
        except:
            date_iso = iso_start = iso_end = None
            season_year = session_year

        duplicates = check_duplicate(date_iso, driver_id, machine_id, field_id, work_type_id, work_name)
        replace_ids = []
        if duplicates:
            print(f"\n⚠️  [РЕЖИМ ДЕБАГУ 00] Знайдено дублікати для робочої операції №{work_row_counter}!")
            print("  1  → Пропустити (Залишити старе завдання)")
            print("  11 → Перезаписати (Видалити старі з сервера та залити нове)")
            dup_ch = input("  Вибір: ").strip()
            if dup_ch == "11": replace_ids = [d["id"] for d in duplicates]
            else:
                final_registry.append([work_row_counter, work_name, "ІСНУЄ", "ІСНУЄ", time_formatted, f"{actual_work_amount} {unit_val}", "Пропущено за вибором"])
                continue

        for old_id in replace_ids: delete_task(old_id)

        mt_id, agro_id, field_mapping_item_id = send_to_cropwise(driver_id, machine_id, work_type_id, field_id, iso_start, iso_end, season_year, send_amount, display_work_name, description)

        if mt_id:
            log_data[r_key] = {
                "status": "sent",
                "machine_task_id": mt_id,
                "agro_operation_id": agro_id,
                "field_mapping_item_id": field_mapping_item_id,
                "field_id": field_id,
                "date": date_val,
                "end_time": iso_end,
                "work": work_name,
                "field": field_name,
                "actual_area": actual_work_amount,
                "unit_val": unit_val,
                "sent_at": datetime.now().isoformat(),
                "fuel_updated": False
            }
            save_log(log_path, log_data)
            stats["sent"] += 1
            final_registry.append([work_row_counter, work_name, f"MT-{mt_id}", f"AO-{agro_id}" if agro_id else "-", time_formatted, f"{actual_work_amount} {unit_val}", "✅ Успішно"])
        else:
            stats["error"] += 1
            final_registry.append([work_row_counter, work_name, "ПОМИЛКА", "ПОМИЛКА", time_formatted, f"{actual_work_amount} {unit_val}", "❌ Відхилено сервером"])

    print_final_report_mode1(driver_name_str, machine_name_str, session_month, session_year, log_path, stats, final_registry)

def print_final_report_mode1(driver_name_str, machine_name_str, session_month, session_year, log_path, stats, final_registry):
    report_lines = []
    report_lines.append("\n" + "=" * 105)
    report_lines.append("🏁 ОБРОБКА ОБЛІКОВОГО ЛИСТА ЗАВЕРШЕНА! (Режим 1: Первинний імпорт)")
    report_lines.append("=" * 105)
    report_lines.append(f"👨 Водій: {driver_name_str:<25} | 🚜 Техніка: {machine_name_str}")
    report_lines.append(f"📅 Період: {MONTHS_UA.get(session_month)} {session_year:<21} | 📁 Лог: {log_path}")
    report_lines.append("-" * 105)
    report_lines.append(f"{'№ Робочий':<9} | {'Робота / Підтип операції':<32} | {'ID Завдання':<11} | {'ID АгроОп':<10} | {'Час':<11} | {'Обсяг':<10} | Стан")
    report_lines.append("-" * 105)
    for item in final_registry:
        report_lines.append(f"Опер. №{item[0]:<2} | {item[1][:32]:<32} | {item[2]:<11} | {item[3]:<10} | {item[4]:<11} | {item[5]:<10} | {item[6]}")
    report_lines.append("-" * 105)
    report_lines.append("📊 ПІДСУМКОВА СТАТИСТИКА СЕСІЇ:")
    report_lines.append(f"  ✅ Надіслано успішно:     {stats['sent']} рядк.")
    report_lines.append(f"  ⏭️  Пропущено дублів:      {stats['skipped']} рядк.")
    report_lines.append(f"  ❌ Помилки відправки:      {stats['error']} рядк.")
    report_lines.append(f"  ⏱️  Загальний облікований час за сесію: {stats['total_hours']} год.")
    report_lines.append("=" * 105)
    report_lines.append("Далі: коли отримаєте дані про пальне, запустіть Режим 2 і вкажіть цей самий файл логу. ⛽")
    report_lines.append("=" * 105)

    for line in report_lines:
        print(colorize_line(line))

    report_file_path = os.path.join(os.path.dirname(log_path), "фінальний_звіт.txt")
    try:
        with open(report_file_path, "w", encoding="utf-8") as rf:
            rf.write("\n".join(report_lines))
        print(f"💾 Текстову копію цього звіту збережено:\n   📁 {report_file_path}")
    except:
        print(f"❌ Не вдалося зберегти фінальний звіт у файл:\n   📁 {report_file_path}")

# =====================================================
# РЕЖИМ 2: ДОДАВАННЯ ПАЛЬНОГО (ОНОВЛЕННЯ ЗА ЛОГОМ)
# =====================================================
# Підтверджено користувачем: колонка з пальним знаходиться в R10C51
# (нотація R1C1: рядок 10, стовпець 51, обидва 1-based) -> 0-based індекс колонки 50.
DEFAULT_FUEL_HEADER_ROW = 10   # 1-based
DEFAULT_FUEL_COLUMN     = 51   # 1-based
DEFAULT_FUEL_COL_IDX    = DEFAULT_FUEL_COLUMN - 1   # 0-based -> 50

def detect_fuel_column(df):
    """Пробує автоматично знайти колонку з пальним за ключовими словами в перших рядках."""
    keywords = ["пальне", "паливо", "пмм", "дизел", "літр", "витрачено пальн"]
    max_scan_rows = min(15, len(df))
    for col in df.columns:
        for row_i in range(max_scan_rows):
            val = df.iat[row_i, col] if col < df.shape[1] else None
            if pd.notna(val) and any(kw in str(val).lower() for kw in keywords):
                return col
    return None

def ask_fuel_column(df):
    # 1) Пробуємо відомий за замовчуванням стовпець R10C51 (0-based індекс 50)
    if df.shape[1] > DEFAULT_FUEL_COL_IDX:
        header_val = None
        header_row_idx = DEFAULT_FUEL_HEADER_ROW - 1
        if df.shape[0] > header_row_idx:
            header_val = df.iat[header_row_idx, DEFAULT_FUEL_COL_IDX]
        print(f"⛽ За замовчуванням використовую колонку пального: R{DEFAULT_FUEL_HEADER_ROW}C{DEFAULT_FUEL_COLUMN} "
              f"(0-based індекс колонки {DEFAULT_FUEL_COL_IDX}).")
        if pd.notna(header_val):
            print(f"   Заголовок у цій позиції: '{header_val}'")
        confirm = input("   Використати цю колонку? (Enter = так, або введіть свою літеру/номер колонки): ").strip()
        if not confirm:
            return DEFAULT_FUEL_COL_IDX
        if column_index_from_string:
            try:
                return column_index_from_string(confirm.upper()) - 1
            except:
                pass
        if confirm.isdigit():
            return int(confirm)
        print("   ⚠️  Не розпізнано введене значення, використовую колонку за замовчуванням.")
        return DEFAULT_FUEL_COL_IDX

    # 2) Якщо в файлі менше колонок, ніж очікувалось — пробуємо автовизначення за ключовими словами
    print(f"⚠️  У файлі менше {DEFAULT_FUEL_COLUMN} колонок, тому типова позиція пального (R{DEFAULT_FUEL_HEADER_ROW}C{DEFAULT_FUEL_COLUMN}) недоступна.")
    auto_col = detect_fuel_column(df)
    if auto_col is not None:
        print(f"⛽ Автоматично визначено колонку пального: індекс {auto_col} (0-based).")
        confirm = input("   Використати цю колонку? (Enter = так, або введіть свою літеру колонки): ").strip()
        if not confirm:
            return auto_col
        if column_index_from_string:
            try:
                return column_index_from_string(confirm.upper()) - 1
            except:
                pass
        if confirm.isdigit():
            return int(confirm)
        print("   ⚠️  Не розпізнано, використовую автоматично визначену колонку.")
        return auto_col

    # 3) Останній варіант — питаємо вручну
    print("⛽ Не вдалося автоматично визначити колонку з пальним.")
    while True:
        col_in = input("   Введіть літеру колонки Excel з пальним (наприклад V) або її номер (0-based): ").strip()
        if column_index_from_string:
            try:
                return column_index_from_string(col_in.upper()) - 1
            except:
                pass
        if col_in.isdigit():
            return int(col_in)
        print("   ⚠️  Невірний формат, спробуйте ще раз (наприклад: V або 21).")

def run_mode_fuel_update(session):
    session_month, session_year = session["month"], session["year"]

    raw_excel = input("\n📂 Шлях до Excel-файлу з даними по пальному: ")
    excel_file = clean_file_path(raw_excel)
    if not os.path.exists(excel_file):
        print("❌ Excel-файл не знайдено!")
        return

    raw_log = input("📂 Шлях до існуючого JSON-логу (з Режиму 1): ")
    log_path = clean_file_path(raw_log)
    if not os.path.exists(log_path):
        print("❌ Файл логу не знайдено!")
        return

    log_data = load_log(log_path)
    if not log_data:
        print("⚠️  Лог порожній або не вдалося прочитати. Перевірте файл.")
        return

    try:
        df = pd.read_excel(excel_file, header=None, engine='openpyxl')
    except Exception as e:
        print(f"❌ Помилка Excel: {e}")
        return

    fuel_col = ask_fuel_column(df)

    valid_rows = [(index, row) for index, row in df.iterrows() if is_valid_work_row(row)]
    if not valid_rows:
        print("⚠️  У файлі не знайдено жодного валідного робочого рядка після фільтрації!")
        return

    print(f"\n✅ Знайдено {len(valid_rows)} робочих рядків для перевірки пального.")

    stats = {"updated": 0, "skipped_done": 0, "not_found": 0, "error": 0, "no_fuel_value": 0}
    report_rows = []

    for original_index, row in valid_rows:
        excel_row_num = original_index + 1
        essentials = extract_row_essentials(row, session_month, session_year, excel_row_num)
        r_key = essentials["key"]
        work_name = essentials["work_name"]
        date_val = essentials["date_val"]
        field_short = essentials["field_short"]
        actual_work_amount = essentials["actual_work_amount"]

        try:
            fuel_raw = row[fuel_col] if fuel_col < len(row) else None
        except Exception:
            fuel_raw = None

        if fuel_raw is None or (isinstance(fuel_raw, float) and pd.isna(fuel_raw)) or str(fuel_raw).strip() in ("", "nan"):
            stats["no_fuel_value"] += 1
            continue

        try:
            fuel_liters = float(str(fuel_raw).replace(",", "."))
        except:
            stats["error"] += 1
            report_rows.append([excel_row_num, work_name, date_val, str(fuel_raw), "❌ Невірний формат пального"])
            continue

        # 1. Спроба знайти за новим стандартом v8 (ключ з '...|rowN' в кінці)
        matched_key = r_key
        log_entry = log_data.get(matched_key)

        # 2. Якщо не знайдено — пробуємо старий стандарт v7 (ключ з обсягом роботи в кінці, напр. "...|6.0")
        if not log_entry:
            old_key = f"{date_val}|{work_name}|{field_short}|{float(actual_work_amount)}"
            log_entry = log_data.get(old_key)
            if log_entry:
                matched_key = old_key
                print(f"   ℹ️  [Сумісність] Знайдено за старим форматом v7 для рядка {excel_row_num}")

        if not log_entry:
            stats["not_found"] += 1
            report_rows.append([excel_row_num, work_name, date_val, fuel_liters, "⚠️  Не знайдено в лозі (пропущено)"])
            continue

        if log_entry.get("fuel_updated") is True:
            stats["skipped_done"] += 1
            report_rows.append([excel_row_num, work_name, date_val, fuel_liters, "⏭️  Вже оновлено раніше"])
            continue

        machine_task_id = log_entry.get("machine_task_id")
        if not machine_task_id:
            stats["error"] += 1
            report_rows.append([excel_row_num, work_name, date_val, fuel_liters, "❌ Немає machine_task_id в лозі"])
            continue

        print(f"⛽ [Рядок Excel {excel_row_num}] Оновлюю пальне для '{work_name[:30]}' → {fuel_liters} л (MT-{machine_task_id})")
        ok, mapping_id = patch_fuel_used(
            machine_task_id,
            fuel_liters,
            field_id=log_entry.get("field_id"),
            field_mapping_item_id=log_entry.get("field_mapping_item_id"),
        )
        if ok:
            log_entry["fuel_used"] = fuel_liters
            log_entry["fuel_updated"] = True
            if mapping_id and not log_entry.get("field_mapping_item_id"):
                log_entry["field_mapping_item_id"] = mapping_id  # донасичуємо старий запис для майбутнього
            log_data[matched_key] = log_entry
            stats["updated"] += 1
            report_rows.append([excel_row_num, work_name, date_val, fuel_liters, "✅ Пальне оновлено"])
        elif mapping_id is None:
            stats["error"] += 1
            report_rows.append([excel_row_num, work_name, date_val, fuel_liters, "❌ Не знайдено MachineTaskFieldMappingItem на сервері"])
        else:
            stats["error"] += 1
            report_rows.append([excel_row_num, work_name, date_val, fuel_liters, "❌ Сервер відхилив оновлення"])

    save_log(log_path, log_data)
    print_final_report_mode2(log_path, stats, report_rows)

def print_final_report_mode2(log_path, stats, report_rows):
    lines = []
    lines.append("\n" + "=" * 100)
    lines.append("🏁 ОНОВЛЕННЯ ПАЛЬНОГО ЗАВЕРШЕНО! (Режим 2)")
    lines.append("=" * 100)
    lines.append(f"📁 Лог: {log_path}")
    lines.append("-" * 100)
    lines.append(f"{'Excel рядок':<11} | {'Робота':<28} | {'Дата':<12} | {'Пальне (л)':<10} | Стан")
    lines.append("-" * 100)
    for item in report_rows:
        lines.append(f"{item[0]:<11} | {str(item[1])[:28]:<28} | {str(item[2]):<12} | {str(item[3]):<10} | {item[4]}")
    lines.append("-" * 100)
    lines.append("📊 ПІДСУМОК:")
    lines.append(f"  ✅ Оновлено пальне:        {stats['updated']} рядк.")
    lines.append(f"  ⏭️  Вже було оновлено:      {stats['skipped_done']} рядк.")
    lines.append(f"  ⚠️  Не знайдено в лозі:     {stats['not_found']} рядк.")
    lines.append(f"  ➖ Без значення пального:  {stats['no_fuel_value']} рядк.")
    lines.append(f"  ❌ Помилки:                {stats['error']} рядк.")
    lines.append("=" * 100)

    for line in lines:
        print(colorize_line(line))

    report_file_path = os.path.join(os.path.dirname(log_path), "звіт_пальне.txt")
    try:
        with open(report_file_path, "w", encoding="utf-8") as rf:
            rf.write("\n".join(lines))
        print(f"💾 Текстову копію цього звіту збережено:\n   📁 {report_file_path}")
    except:
        print(f"❌ Не вдалося зберегти звіт у файл:\n   📁 {report_file_path}")

# =====================================================
# РЕЖИМ 3: АВТОМАТИЧНЕ СПИСАННЯ ПАЛЬНОГО ЗІ СКЛАДУ
# =====================================================
# ВАЖЛИВО: реальні об'єкти Cropwise API для складського обліку — це WhDocument
# (operations.cropwise.com/api/v3/wh_documents) та WhDocumentItem
# (operations.cropwise.com/api/v3/wh_document_items). Ендпоінт /api/v3/warehouse_documents,
# поле 'additional_type': 'goods_issue_mtfmi_fuel' та концепція 'fields_work_result_id' З
# ОФІЦІЙНОЇ ДОКУМЕНТАЦІЇ НЕ ІСНУЮТЬ — це й була причина, чому попередня версія Режиму 3 не могла
# працездатно списувати пальне (сервер або відхиляв запит, або мовчки ігнорував невідомі поля).
#
# Списання (goods_issue) тепер прив'язується до вже наявної AgroOperation (яку Режим 1 створює
# і зберігає agro_operation_id в лозі) через related_document_type='AgroOperation' — це єдиний
# задокументований спосіб пов'язати складський документ з конкретною роботою.
DEFAULT_WAREHOUSE_ID = 13       # Склад, підтверджений користувачем
DEFAULT_FUEL_WH_ITEM_ID = 171   # WhItem "Пальне", підтверджений користувачем

def get_wh_item_unit_id(wh_item_id):
    """Дізнається базову одиницю виміру (unit_id) для WhItem — потрібна при створенні WhDocumentItem."""
    try:
        resp = api_request("GET", f"{API_BASE}/api/v3/wh_items/{wh_item_id}", headers=get_headers(), timeout=10)
        if resp.status_code == 200:
            item = resp.json().get("data", {})
            return item.get("base_inventory_unit_id") or item.get("wh_item_base_unit_id")
        return None
    except:
        return None

def create_wh_related_document_assignment(wh_doc_id, related_document_type, related_document_id):
    """
    Окремий об'єкт WhRelatedDocumentAssignment (operations.cropwise.com/api/v3/wh_related_document_assignments).
    Судячи з поведінки реального акаунту, саме ЦЕЙ окремий запис — а не поле related_document_type/id
    на самому WhDocument — і є тим, що показує рядок "На підставі: #... / Робота" в інтерфейсі.
    """
    payload = {"data": {
        "wh_document_id": int(wh_doc_id),
        "related_document_type": related_document_type,
        "related_document_id": int(related_document_id),
    }}
    try:
        r = api_request("POST", f"{API_BASE}/api/v3/wh_related_document_assignments", headers=get_headers(),
                         data=json.dumps(payload, ensure_ascii=False).encode("utf-8"), timeout=15)
        return r.status_code in (200, 201), (r.text if r.status_code not in (200, 201) else None)
    except Exception as e:
        return False, str(e)

def create_fuel_goods_issue(agro_operation_id, date_iso, amount_liters, machine_id=None, field_id=None, machine_task_id=None,
                             end_time_iso=None, warehouse_id=DEFAULT_WAREHOUSE_ID, wh_item_id=DEFAULT_FUEL_WH_ITEM_ID):
    """
    Створює складське списання пального (саме так влаштована реальна схема):
    1. WhDocument (шапка документа: тип, склад, статус, прив'язка споживання, час документа)
    2. WhDocumentItem (сам рядок: яка позиція складу (пальне) і скільки списано)
    3. WhRelatedDocumentAssignment (прив'язка до AgroOperation — дає підпис "На підставі" в UI)

    Прив'язка до техніки/поля робиться через 'consumption_target_type'/'consumption_target_id'
    ('Field', 'Machine', 'Implement' або 'Counterparty' — прямого 'MachineTask' в API немає).

    document_time: якщо переданий end_time_iso (реальний час завершення роботи, збережений
    Режимом 1 чи підтягнутий з сервера для старих записів) — використовуємо його. Інакше
    (резервний варіант, якщо часу ніде взяти) ставимо полудень обраної дати.
    """
    document_time = end_time_iso if end_time_iso else f"{date_iso}T12:00:00"

    doc_payload = {
        "data": {
            "document_type": "goods_issue",
            "document_status": "accepted",
            "from_warehouse_id": int(warehouse_id),  # ПІДТВЕРДЖЕНО СЕРВЕРОМ: для 'goods_issue' саме from_warehouse_id
            "document_time": document_time,
        }
    }
    if field_id and str(field_id) != "0":
        doc_payload["data"]["consumption_target_type"] = "Field"
        doc_payload["data"]["consumption_target_id"] = int(field_id)
    elif machine_id:
        doc_payload["data"]["consumption_target_type"] = "Machine"
        doc_payload["data"]["consumption_target_id"] = int(machine_id)

    if agro_operation_id:
        doc_payload["data"]["related_document_type"] = "AgroOperation"
        doc_payload["data"]["related_document_id"] = int(agro_operation_id)

    try:
        r = api_request("POST", f"{API_BASE}/api/v3/wh_documents", headers=get_headers(),
                         data=json.dumps(doc_payload, ensure_ascii=False).encode("utf-8"), timeout=15)
        if r.status_code not in (200, 201):
            return False, None, f"WhDocument відхилено: {r.status_code} {r.text}"
        wh_doc_id = r.json().get("data", {}).get("id")
    except Exception as e:
        return False, None, f"Помилка з'єднання (WhDocument): {e}"

    unit_id = get_wh_item_unit_id(wh_item_id)
    item_payload = {"data": {
        "wh_document_id": int(wh_doc_id),
        "wh_item_id": int(wh_item_id),
        "position": 1,  # обов'язкове поле — порядковий номер позиції в документі (у нас завжди одна позиція)
        "quantity": float(amount_liters),
    }}
    if unit_id:
        item_payload["data"]["unit_id"] = unit_id

    try:
        show_loading_animation("/api/v3/wh_document_items", duration=0.3)
        r2 = api_request("POST", f"{API_BASE}/api/v3/wh_document_items", headers=get_headers(),
                          data=json.dumps(item_payload, ensure_ascii=False).encode("utf-8"), timeout=15)
        if r2.status_code not in (200, 201):
            return False, wh_doc_id, f"WhDocumentItem відхилено: {r2.status_code} {r2.text}"
    except Exception as e:
        return False, wh_doc_id, f"Помилка з'єднання (WhDocumentItem): {e}"

    if agro_operation_id:
        ok_assign, assign_err = create_wh_related_document_assignment(wh_doc_id, "AgroOperation", agro_operation_id)
        if not ok_assign:
            print(f"   ⚠️  Документ створено, але прив'язку 'На підставі' до AgroOperation не вдалося встановити: {assign_err}")

    # ПРИМІТКА: пробували додатково прив'язати "MachineTask" через WhRelatedDocumentAssignment —
    # сервер повертає 500 Internal Server Error (а не акуратну 422-валідацію), тобто цей тип
    # дійсно не підтримується. Прив'язка через AgroOperation (вище) — це задокументована межа
    # можливого; вона вже показує "На підставі: #... Робота" в інтерфейсі, чого достатньо для
    # трасування списання до конкретної роботи.

    return True, wh_doc_id, None

def run_mode_fuel_write_off(session):
    """
    РЕЖИМ 3: Списання пального зі складу (WhDocument/WhDocumentItem) за даними логу.
    Використовує agro_operation_id, вже збережений у лозі Режимом 1, та fuel_used,
    збережений Режимом 2 — обидва режими мають бути виконані заздалегідь для кожного рядка.
    """
    print("\n" + "=" * 70)
    print("📦 РЕЖИМ 3: АВТОМАТИЧНЕ СПИСАННЯ ПАЛЬНОГО ЗІ СКЛАДУ")
    print("=" * 70)

    raw_log = input("📂 Шлях до існуючого JSON-логу: ")
    log_path = clean_file_path(raw_log)
    if not os.path.exists(log_path):
        print("❌ Файл логу не знайдено!")
        return

    log_data = load_log(log_path)
    if not log_data:
        print("⚠️  Лог порожній.")
        return

    stats = {"written_off": 0, "skipped_already": 0, "errors": 0}
    report_rows = []

    for key, entry in log_data.items():
        work_name = entry.get("work", "Невідома робота")
        date_raw = entry.get("date")  # формат DD.MM.YYYY

        if entry.get("fuel_written_off") is True:
            stats["skipped_already"] += 1
            report_rows.append([date_raw, work_name, entry.get("fuel_used", 0), "⏭️ Вже списано"])
            continue

        fuel_amount = entry.get("fuel_used")  # очікуємо значення, збережене Режимом 2
        agro_operation_id = entry.get("agro_operation_id")

        if fuel_amount is None or float(fuel_amount) <= 0:
            print(f"⚠️  Пропущено '{work_name}' від {date_raw}: немає збереженого значення пального (спершу запустіть Режим 2).")
            stats["errors"] += 1
            report_rows.append([date_raw, work_name, 0, "❌ Немає літрів у лозі"])
            continue

        if not agro_operation_id:
            print(f"⚠️  '{work_name}' від {date_raw}: немає agro_operation_id — списання буде без прив'язки до операції.")

        field_id = entry.get("field_id")
        if not field_id and entry.get("machine_task_id"):
            field_id = get_field_id_for_task(entry.get("machine_task_id"))
            if field_id:
                entry["field_id"] = field_id  # донасичуємо старий запис логу для наступних запусків
                print(f"   ℹ️  [Сумісність] field_id підтягнуто з сервера для старого запису логу (MT-{entry.get('machine_task_id')})")

        end_time_iso = entry.get("end_time")
        if not end_time_iso and entry.get("machine_task_id"):
            end_time_iso = get_machine_task_end_time(entry.get("machine_task_id"))
            if end_time_iso:
                entry["end_time"] = end_time_iso  # донасичуємо старий запис логу для наступних запусків
                print(f"   ℹ️  [Сумісність] end_time підтягнуто з сервера для старого запису логу (MT-{entry.get('machine_task_id')})")

        try:
            parts = date_raw.split('.')
            date_formatted = f"{parts[2]}-{parts[1]}-{parts[0]}"
        except Exception:
            date_formatted = date_raw

        print(f"\n🔄 Обробка '{work_name}' від {date_raw} ({fuel_amount} л)...")
        success, doc_id, err = create_fuel_goods_issue(
            agro_operation_id, date_formatted, fuel_amount,
            machine_id=session.get("machine_id"),
            field_id=field_id,
            machine_task_id=entry.get("machine_task_id"),
            end_time_iso=end_time_iso,
        )
        if success:
            print(f"   ✅ Успішно списано {fuel_amount} л пального! Створено складський документ № {doc_id}")
            entry["fuel_written_off"] = True
            entry["warehouse_doc_id"] = doc_id
            log_data[key] = entry
            stats["written_off"] += 1
            report_rows.append([date_raw, work_name, fuel_amount, f"✅ Списано (Doc {doc_id})"])
            save_log(log_path, log_data)  # зберігаємо прогрес одразу після кожного успіху
        else:
            print(f"   ❌ {err}")
            stats["errors"] += 1
            report_rows.append([date_raw, work_name, fuel_amount, "❌ Помилка створення документа"])

    print_final_report_mode3(log_path, stats, report_rows)
    mark_session_folder_complete_if_done(log_path, log_data)

def mark_session_folder_complete_if_done(log_path, log_data):
    """
    Якщо всі рядки логу, для яких уже є пальне (fuel_used, тобто пройшли Режим 2),
    успішно списані (fuel_written_off=True), додає до назви папки сесії позначку
    завершення — щоб у провіднику Windows одразу було видно, які сесії повністю
    закриті (Режими 1+2+3 виконано). До Режиму 3 дійти без 1 і 2 неможливо,
    тож ця позначка — надійний індикатор повного циклу.

    На Windows перейменування папки іноді падає з "Access is denied" (WinError 5),
    якщо якийсь файл усередині ще "тримається" іншою програмою (відкритий Excel,
    антивірус щойно сканує звіт тощо) — зазвичай це тимчасово, тож пробуємо кілька
    разів з паузою. Якщо й це не допомагає — не валимо процес, а ставимо маркер-файл
    усередині папки як надійний резервний варіант позначення завершення.
    """
    eligible = [e for e in log_data.values() if e.get("fuel_used") is not None]
    if not eligible:
        return  # ще нема жодного рядка з пальним — позначати нема за що
    if not all(e.get("fuel_written_off") is True for e in eligible):
        return  # ще не всі рядки списані — рано позначати

    session_folder = os.path.dirname(os.path.abspath(log_path))
    folder_basename = os.path.basename(session_folder)
    marker = " [✅ ЗАВЕРШЕНО]"
    if folder_basename.endswith(marker):
        return  # вже позначено раніше

    marker_file_path = os.path.join(session_folder, "_ЗАВЕРШЕНО.txt")
    if os.path.exists(marker_file_path):
        return  # вже позначено раніше через резервний маркер-файл

    parent_dir = os.path.dirname(session_folder)
    new_folder = os.path.join(parent_dir, folder_basename + marker)

    last_err = None
    for attempt in range(3):
        try:
            os.rename(session_folder, new_folder)
            print(f"\n🏷️  Усі рядки пального списано — папку сесії позначено як завершену:\n   📁 {new_folder}")
            return
        except Exception as e:
            last_err = e
            if attempt < 2:
                time.sleep(1.5)

    # Перейменування не вдалося (типово через відкритий у Windows файл усередині папки) —
    # ставимо маркер-файл як надійний резервний індикатор завершення, щоб не втратити прогрес.
    try:
        with open(marker_file_path, "w", encoding="utf-8") as f:
            f.write("Усі рядки пального в цій сесії успішно списано (Режими 1+2+3 виконано).\n")
            f.write(f"Час: {datetime.now().isoformat()}\n")
        print(f"⚠️  Не вдалося перейменувати папку сесії (можливо, якийсь файл усередині відкритий у "
              f"Windows/Excel): {last_err}")
        print(f"   ℹ️  Замість цього створено маркер-файл завершення: {marker_file_path}")
        print("   Можете закрити відкриті файли й перейменувати папку вручну, якщо потрібно.")
    except Exception as e2:
        print(f"⚠️  Не вдалося ні перейменувати папку сесії ({last_err}), ні створити маркер-файл ({e2}).")

def print_final_report_mode3(log_path, stats, report_rows):
    lines = []
    lines.append("\n" + "=" * 100)
    lines.append("🏁 АВТОМАТИЧНЕ СПИСАННЯ ПАЛЬНОГО ЗАВЕРШЕНО! (Режим 3)")
    lines.append("=" * 100)
    lines.append(f"📁 Лог: {log_path}")
    lines.append("-" * 100)
    lines.append(f"{'Дата':<12} | {'Робота':<35} | {'Обсяг (л)':<10} | Стан")
    lines.append("-" * 100)
    for item in report_rows:
        lines.append(f"{str(item[0]):<12} | {str(item[1])[:35]:<35} | {str(item[2]):<10} | {item[3]}")
    lines.append("-" * 100)
    lines.append("📊 ПІДСУМОК:")
    lines.append(f"  ✅ Списано успішно:        {stats['written_off']} рядк.")
    lines.append(f"  ⏭️  Вже було списано:      {stats['skipped_already']} рядк.")
    lines.append(f"  ❌ Помилки:                {stats['errors']} рядк.")
    lines.append("=" * 100)

    for line in lines:
        print(colorize_line(line))

    report_file_path = os.path.join(os.path.dirname(log_path), "звіт_списання.txt")
    try:
        with open(report_file_path, "w", encoding="utf-8") as rf:
            rf.write("\n".join(lines))
        print(f"💾 Текстову копію цього звіту збережено:\n   📁 {report_file_path}")
    except:
        print(f"❌ Не вдалося зберегти звіт у файл:\n   📁 {report_file_path}")

# =====================================================
# СТАРТ ПРОГРАМИ
# =====================================================
def main():
    print("=" * 70)
    print("🚀 CROPWISE АВТОПІЛОТ — Версія v10.2 (Імпорт + Пальне + Списання (Додана площа))")
    print("=" * 70)

    session = init_session()

    while True:
        print("\n" + "=" * 70)
        print("🧭 ОБЕРІТЬ РЕЖИМ РОБОТИ:")
        print("  1 → Первинний імпорт робіт та створення логу")
        print("  2 → Додавання пального (оновлення за логом)")
        print("  3 → Автоматичне списання пального зі складу 🚀")
        print("  0 → Вихід")
        print("=" * 70)
        mode = input("👉 Ваш вибір: ").strip()

        if mode == "1":
            run_mode_primary_import(session)
        elif mode == "2":
            run_mode_fuel_update(session)
        elif mode == "3":
            run_mode_fuel_write_off(session)
        elif mode == "0":
            print("👋 Роботу завершено.")
            break
        else:
            print("⚠️  Невірний вибір, спробуйте ще раз.")
            continue

        again = input("\n🔁 Виконати ще одну дію в цій самій сесії (той самий водій/машина)? (т/н): ").strip().lower()
        if again not in ("т", "так", "y", "yes"):
            print("👋 Роботу завершено.")
            break

if __name__ == "__main__":
    main()
