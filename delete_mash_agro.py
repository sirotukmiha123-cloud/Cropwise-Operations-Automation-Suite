import json
import os
import sys
from cropwise_common import API_BASE, get_headers, api_request, clean_file_path

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

def get_agro_operation_id_for_task(machine_task_id):
    """Шукає ID AgroOperation, пов'язаної з конкретним MachineTask, через мапінг-таблицю."""
    try:
        resp = api_request("GET", f"{API_BASE}/api/v3/machine_task_agro_operation_mapping_items",
                            headers=get_headers(), params={"machine_task_id": machine_task_id, "limit": 5}, timeout=10)
        if resp.status_code == 200:
            items = resp.json().get("data", [])
            if items:
                return items[0].get("agro_operation_id")
        return None
    except Exception:
        return None

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

    print(f"\n📋 Знайдено {len(items)} запис(ів) у лозі, які можна видалити з Cropwise:")
    for _, entry, mt_id, ao_id in items:
        work = str(entry.get("work", "?"))[:40]
        print(f"  • {entry.get('date','?'):<12} | {work:<40} | MT-{mt_id} | AO-{ao_id}")

    print(f"\n⚠️  УВАГА: це НЕЗВОРОТНО видалить {len(items)} завдань машин та пов'язаних агроопераций з Cropwise!")
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

    # Прохід 2: тепер, коли завдання вже видалені, видаляємо агрооперації
    # (delete_agro_operation додатково сама повторює спробу з паузою про всяк випадок)
    print("\n🔻 Крок 2/2: видаляю агрооперації (AgroOperation)...")
    for _, entry, _, ao_id in items:
        label = f"{entry.get('date','?')} | {str(entry.get('work','?'))[:40]}"
        if not ao_id:
            stats["ao_skipped"] += 1
            continue
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
            found_ao = get_agro_operation_id_for_task(mt_id)
            if found_ao:
                discovered_ao_ids.append(found_ao)
                print(f"  • MT-{mt_id} → знайдено AO-{found_ao}")
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
    