import json
import os
import sys
import time
import requests

# =====================================================
# КОНФІГУРАЦІЯ (той самий підхід до токена, що й в основному скрипті —
# може використовувати той самий config.json, якщо покласти цей файл поруч)
# =====================================================
API_BASE = "https://operations.cropwise.com"

def load_cropwise_token():
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
        except Exception:
            pass

    print("⚠️  Токен не знайдено ні в змінній середовища CROPWISE_TOKEN, ні в config.json поруч зі скриптом.")
    print("   (Покладіть цей файл в ту саму папку, що й основний скрипт agro1mash — тоді config.json підхопиться сам.)")
    return input("🔑 Введіть API-токен Cropwise вручну: ").strip()

CROPWISE_TOKEN = load_cropwise_token()

def get_headers():
    return {
        "X-User-API-Token": str(CROPWISE_TOKEN).strip(),
        "Accept": "application/json",
        "Content-Type": "application/json; charset=utf-8",
    }

def api_request(method, url, max_retries=3, **kwargs):
    """Обгортка з обробкою 429 (Too Many Requests) та мережевих збоїв — так само, як в основному скрипті."""
    last_exc = None
    for attempt in range(max_retries + 1):
        try:
            resp = requests.request(method, url, **kwargs)
        except requests.exceptions.RequestException as e:
            last_exc = e
            if attempt < max_retries:
                time.sleep(1.5 * (attempt + 1))
                continue
            raise
        if resp.status_code == 429 and attempt < max_retries:
            retry_after = resp.headers.get("Retry-After")
            try:
                wait_s = float(retry_after) if retry_after else 2.0 * (attempt + 1)
            except (TypeError, ValueError):
                wait_s = 2.0 * (attempt + 1)
            print(f"  ⏳ Сервер повернув 429, пауза {wait_s:.1f} с...")
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

# =====================================================
# ВИДАЛЕННЯ
# =====================================================
def delete_machine_task(task_id):
    try:
        r = api_request("DELETE", f"{API_BASE}/api/v3/machine_tasks/{task_id}", headers=get_headers(), timeout=15)
        if r.status_code in (200, 204):
            return True, None
        return False, f"{r.status_code} {r.text[:200]}"
    except Exception as e:
        return False, str(e)

def delete_agro_operation(agro_id, retries=2, retry_delay=1.5):
    """
    Видаляє AgroOperation. Якщо об'єкт "прив'язаний" (ще посилається на щойно видалене
    MachineTask, поки Cropwise не встиг оновити зв'язок), повторює спробу через паузу —
    замість того щоб одразу здаватись.
    """
    last_err = None
    for attempt in range(retries + 1):
        try:
            r = api_request("DELETE", f"{API_BASE}/api/v3/agro_operations/{agro_id}", headers=get_headers(), timeout=15)
            if r.status_code in (200, 204):
                return True, None
            last_err = f"{r.status_code} {r.text[:200]}"
        except Exception as e:
            last_err = str(e)

        if attempt < retries:
            time.sleep(retry_delay)

    return False, last_err

def get_agro_operation_ids_for_task(machine_task_id):
    """
    Шукає ID УСІХ AgroOperation, пов'язаних із конкретним MachineTask, через мапінг-таблицю.
    ВАЖЛИВО: одне завдання машини може мати КІЛЬКА пов'язаних агроопераций (напр. коли поля
    вводились вручну по одному) — раніше ця функція повертала лише items[0], тому видалялась
    тільки одна з кількох.
    """
    try:
        resp = api_request("GET", f"{API_BASE}/api/v3/machine_task_agro_operation_mapping_items",
                            headers=get_headers(), params={"machine_task_id": machine_task_id, "limit": 100}, timeout=10)
        if resp.status_code == 200:
            items = resp.json().get("data", [])
            return [item.get("agro_operation_id") for item in items if item.get("agro_operation_id")]
        return []
    except Exception:
        return []

def load_log(log_path):
    with open(log_path, "r", encoding="utf-8") as f:
        return json.load(f)

# =====================================================
# РЕЖИМ 1: ВИДАЛЕННЯ ЗА ФАЙЛОМ ЛОГУ
# =====================================================
def mode_delete_by_log():
    raw = input("\n📂 Шлях до JSON-логу (можна перетягнути файл у консоль): ")
    log_path = clean_file_path(raw)
    if not os.path.exists(log_path):
        print("❌ Файл логу не знайдено!")
        return

    try:
        log_data = load_log(log_path)
    except Exception as e:
        print(f"❌ Не вдалося прочитати лог: {e}")
        return

    if not log_data:
        print("⚠️  Лог порожній.")
        return

    items = []
    for key, entry in log_data.items():
        mt_id = entry.get("machine_task_id")
        ao_id = entry.get("agro_operation_id")
        if mt_id or ao_id:
            items.append((key, entry, mt_id, ao_id))

    if not items:
        print("⚠️  У лозі немає жодного запису з machine_task_id/agro_operation_id.")
        return

    # Лог зберігає лише ОДИН agro_operation_id на завдання, але на сервері до того самого
    # MachineTask могло прив'язатись кілька агроопераций (напр. поля вводились вручну по
    # одному) — тому додатково опитуємо сервер за кожним завданням, щоб не пропустити зайві.
    print("\n🔎 Перевіряю сервер на додаткові агрооперації, не відображені в лозі...")
    enriched_items = []
    for key, entry, mt_id, ao_id in items:
        all_ao_ids = set()
        if ao_id:
            all_ao_ids.add(ao_id)
        if mt_id:
            server_ao_ids = get_agro_operation_ids_for_task(mt_id)
            extra = set(server_ao_ids) - all_ao_ids
            if extra:
                print(f"  • MT-{mt_id}: у лозі AO-{ao_id}, на сервері ще знайдено {', '.join(f'AO-{a}' for a in extra)}")
            all_ao_ids.update(server_ao_ids)
        enriched_items.append((key, entry, mt_id, sorted(all_ao_ids)))
    items = enriched_items

    total_ao = sum(len(ao_list) for _, _, _, ao_list in items)
    print(f"\n📋 Знайдено {len(items)} запис(ів) у лозі, які можна видалити з Cropwise ({total_ao} агроопераций сумарно):")
    for _, entry, mt_id, ao_list in items:
        work = str(entry.get("work", "?"))[:40]
        ao_str = ", ".join(f"AO-{a}" for a in ao_list) if ao_list else "—"
        print(f"  • {entry.get('date','?'):<12} | {work:<40} | MT-{mt_id} | {ao_str}")

    print(f"\n⚠️  УВАГА: це НЕЗВОРОТНО видалить {len(items)} завдань машин та {total_ao} пов'язаних агроопераций з Cropwise!")
    confirm = input("Введіть 'ВИДАЛИТИ' (великими літерами), щоб підтвердити: ").strip()
    if confirm != "ВИДАЛИТИ":
        print("Скасовано.")
        return

    stats = {"mt_deleted": 0, "mt_error": 0, "ao_deleted": 0, "ao_error": 0, "ao_skipped": 0}

    # Прохід 1: спочатку ВСІ завдання машин — доки завдання ще "прив'язане" до агрооперації,
    # видалити саму агрооперацію Cropwise не дасть.
    print("\n🔻 Крок 1/2: видаляю завдання машин (MachineTask)...")
    for _, entry, mt_id, _ in items:
        if not mt_id:
            continue
        label = f"{entry.get('date','?')} | {str(entry.get('work','?'))[:40]}"
        ok, err = delete_machine_task(mt_id)
        if ok:
            print(f"  ✅ Видалено MachineTask MT-{mt_id} ({label})")
            stats["mt_deleted"] += 1
        else:
            print(f"  ❌ Не вдалося видалити MachineTask MT-{mt_id}: {err}")
            stats["mt_error"] += 1

    # Прохід 2: тепер, коли завдання вже видалені, видаляємо УСІ пов'язані агрооперації
    # (delete_agro_operation додатково сама повторює спробу з паузою про всяк випадок)
    print("\n🔻 Крок 2/2: видаляю агрооперації (AgroOperation)...")
    for _, entry, _, ao_list in items:
        label = f"{entry.get('date','?')} | {str(entry.get('work','?'))[:40]}"
        if not ao_list:
            stats["ao_skipped"] += 1
            continue
        for ao_id in ao_list:
            ok2, err2 = delete_agro_operation(ao_id)
            if ok2:
                print(f"  ✅ Видалено AgroOperation AO-{ao_id} ({label})")
                stats["ao_deleted"] += 1
            else:
                print(f"  ❌ Не вдалося видалити AgroOperation AO-{ao_id}: {err2}")
                stats["ao_error"] += 1

    print("\n" + "=" * 70)
    print("📊 ПІДСУМОК ВИДАЛЕННЯ:")
    print(f"  MachineTask   → видалено: {stats['mt_deleted']}, помилок: {stats['mt_error']}")
    print(f"  AgroOperation → видалено: {stats['ao_deleted']}, помилок: {stats['ao_error']}, без ID: {stats['ao_skipped']}")
    print("=" * 70)

    if stats["mt_error"] == 0 and stats["ao_error"] == 0:
        clear_choice = input("\nВидалити також сам файл логу (щоб не переплутати зі старою сесією)? (т/н): ").strip().lower()
        if clear_choice in ("т", "так", "y", "yes"):
            try:
                os.remove(log_path)
                print(f"🗑️  Файл логу видалено: {log_path}")
            except Exception as e:
                print(f"⚠️  Не вдалося видалити файл логу: {e}")

# =====================================================
# РЕЖИМ 2: РУЧНЕ ВИДАЛЕННЯ ЗА СПИСКОМ ID
# =====================================================
def parse_id_spec(spec_str):
    """
    Розбирає рядок ID у відсортований список цілих чисел без дублікатів. Підтримує:
    - Точкові ID через кому: "1234,5436"
    - Діапазони через дефіс: "1500-1534"
    - Змішано: "1234,1500-1534,5436"
    """
    ids = set()
    spec_str = (spec_str or "").strip()
    if not spec_str:
        return []
    parts = [p.strip() for p in spec_str.split(",") if p.strip()]
    for part in parts:
        if "-" in part:
            range_parts = part.split("-")
            if len(range_parts) == 2 and range_parts[0].strip().isdigit() and range_parts[1].strip().isdigit():
                start, end = int(range_parts[0].strip()), int(range_parts[1].strip())
                if start > end:
                    start, end = end, start
                ids.update(range(start, end + 1))
            else:
                print(f"⚠️  Не вдалося розпізнати діапазон '{part}', пропускаю.")
        elif part.isdigit():
            ids.add(int(part))
        else:
            print(f"⚠️  Не вдалося розпізнати значення '{part}', пропускаю.")
    return sorted(ids)

def mode_delete_manual():
    print("\nВведіть ID: точково через кому (1234,5436), діапазоном (1500-1534) або змішано")
    print("(1234,1500-1534,5436). Enter — пропустити, якщо не потрібно видаляти цей тип:")
    mt_ids_raw = input("  ID завдань машин (MachineTask): ").strip()
    ao_ids_raw = input("  ID агроопераций (AgroOperation, можна лишити порожнім — знайду сам за завданнями): ").strip()

    mt_ids = parse_id_spec(mt_ids_raw)
    ao_ids = parse_id_spec(ao_ids_raw)

    if not mt_ids and not ao_ids:
        print("Нічого не вказано — вихід.")
        return

    if mt_ids:
        print("\n🔎 Шукаю прив'язані агрооперації для вказаних завдань машин...")
        discovered_ao_ids = []
        for mt_id in mt_ids:
            found_aos = get_agro_operation_ids_for_task(mt_id)
            if found_aos:
                discovered_ao_ids.extend(found_aos)
                if len(found_aos) == 1:
                    print(f"  • MT-{mt_id} → знайдено AO-{found_aos[0]}")
                else:
                    print(f"  • MT-{mt_id} → знайдено {len(found_aos)} агроопераций: {', '.join(f'AO-{a}' for a in found_aos)}")
            else:
                print(f"  • MT-{mt_id} → агрооперацію не знайдено (можливо, її немає або вже видалена)")
        ao_ids = sorted(set(ao_ids) | set(discovered_ao_ids))

    def preview(ids, label):
        if not ids:
            return f"0 {label}"
        if len(ids) <= 30:
            return f"{len(ids)} {label} ({', '.join(str(i) for i in ids)})"
        return f"{len(ids)} {label} (від {ids[0]} до {ids[-1]}, список завеликий для показу повністю)"

    print(f"\n📋 До видалення: {preview(mt_ids, 'завдань машин')}")
    print(f"             та: {preview(ao_ids, 'агроопераций')}")

    total = len(mt_ids) + len(ao_ids)
    if total > 50:
        print(f"\n🚨 Це ВЕЛИКА кількість об'єктів ({total})! Перевірте діапазон ще раз перед підтвердженням.")

    confirm = input("Введіть 'ВИДАЛИТИ' (великими літерами), щоб підтвердити: ").strip()
    if confirm != "ВИДАЛИТИ":
        print("Скасовано.")
        return

    for mt_id in mt_ids:
        ok, err = delete_machine_task(mt_id)
        print(f"  {'✅' if ok else '❌'} MachineTask MT-{mt_id}" + (f": {err}" if err else ""))

    for ao_id in ao_ids:
        ok, err = delete_agro_operation(ao_id)
        print(f"  {'✅' if ok else '❌'} AgroOperation AO-{ao_id}" + (f": {err}" if err else ""))

# =====================================================
# СТАРТ ПРОГРАМИ
# =====================================================
def main():
    print("=" * 70)
    print("🗑️  CROPWISE — Видалення завдань машин та агроопераций")
    print("=" * 70)
    print("  1 → Видалити все за файлом логу (лог.json з основного скрипта)")
    print("  2 → Видалити вручну за списком ID")
    print("  0 → Вихід")
    print("=" * 70)
    choice = input("👉 Ваш вибір: ").strip()

    if choice == "1":
        mode_delete_by_log()
    elif choice == "2":
        mode_delete_manual()
    else:
        print("👋 Вихід.")
        sys.exit(0)

if __name__ == "__main__":
    main()
