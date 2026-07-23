import pandas as pd
import os
import json
from cropwise_common import API_BASE, get_headers, api_request

API_DOCS_URL = f"{API_BASE}/api/v3/wh_documents"
API_ITEMS_URL = f"{API_BASE}/api/v3/wh_document_items"
API_WH_ITEMS_URL = f"{API_BASE}/api/v3/wh_items"

CACHE_FILE = "mapping_cache.txt"

def load_permanent_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_permanent_cache(cache_data):
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache_data, f, ensure_ascii=False, indent=4)
    except:
        pass

def find_item_in_cropwise_online(item_name):
    params = {"q[name_cont]": item_name.strip()} 
    try:
        res = api_request("GET", API_WH_ITEMS_URL, headers=get_headers(), params=params, timeout=10)
        if res.status_code == 200:
            items_list = res.json().get("data", [])
            for item in items_list:
                if item.get("name", "").strip().lower() == item_name.strip().lower():
                    return int(item["id"]), item["name"], int(item.get("wh_item_base_unit_id") or 1)
    except:
        pass
    return None

def run_waiting_import():
    print("=== 🌐 ІМПОРТ ТМЦ (v4.6 - РЕЖИМ ОЧІКУВАННЯ СТВОРЕННЯ) ===")
    
    raw_path = input("📁 Перетягніть файл Excel у консоль: ").strip()
    excel_file = raw_path.strip("'\"").strip() if raw_path else "oborotka.xlsx"
    if raw_path.startswith("& "): excel_file = raw_path[2:].strip().strip("'\"")

    if not os.path.exists(excel_file):
        print(f"❌ Файл не знайдено!")
        return

    try: target_warehouse_id = int(input("🏢 Введіть ID Складу призначення [13]: ").strip())
    except: target_warehouse_id = 13

    try: counterparty_id = int(input("🤝 Введіть ID Контрагента [9]: ").strip())
    except: counterparty_id = 9

    df = pd.read_excel(excel_file, header=None, engine='openpyxl')
    print(" ✅ Файл Excel успішно зчитано!")

    permanent_cache = load_permanent_cache()
    print("\n📦 Створюю новий документ надходження...")
    
    doc_payload = {
        "data": {
            "document_type": "goods_receipt",
            "document_status": "accepted",
            "document_time": "2026-06-01T09:00:00.000+03:00",
            "to_warehouse_id": target_warehouse_id,
            "counterparty_id": counterparty_id,
            "additional_info": f"Імпорт через v4.6: {os.path.basename(excel_file)}"
        }
    }
    
    doc_res = api_request("POST", API_DOCS_URL, headers=get_headers(), json=doc_payload, timeout=10)
    doc_id = doc_res.json().get('data', {}).get('id')
    print(f" 🎉 Документ створено! ID: {doc_id}")

    TARGET_COL_INDEX = 6  
    NAME_COL_INDEX = 0    
    attempt_count = 0
    total_rows = len(df)

    print("\n🚀 Починаємо розносити позиції...")

    for index in range(9, total_rows, 2):
        row_sum = df.iloc[index]
        if index + 1 >= total_rows: break
        row_qty = df.iloc[index + 1]

        cell_0 = str(row_sum.iloc[NAME_COL_INDEX]).strip() if pd.notna(row_sum.iloc[NAME_COL_INDEX]) else ""
        if not cell_0 or any(x in cell_0.lower() for x in ["разом", "всього", "рахунок", "тмц"]):
            continue

        try:
            sum_str = str(row_sum.iloc[TARGET_COL_INDEX]).replace(' ', '').replace(',', '').strip()
            total_sum = float(sum_str) if sum_str and sum_str != 'nan' else 0.0
            qty_str = str(row_qty.iloc[TARGET_COL_INDEX]).replace(' ', '').replace(',', '').strip()
            quantity = float(qty_str) if qty_str and qty_str != 'nan' else 0.0
        except:
            continue

        if quantity <= 0:
            continue

        cell_0_lower = cell_0.lower()
        
        # 1. Перевіряємо локальну пам'ять (минулі місяці / поточна сесія)
        if cell_0_lower in permanent_cache:
            cached_data = permanent_cache[cell_0_lower]
            item_id = cached_data["id"]
            unit_id = cached_data["unit_id"]
            cropwise_name = cached_data["name"]
        else:
            # 2. Пробуємо знайти в Cropwise (може вже є точний збіг)
            found = find_item_in_cropwise_online(cell_0)
            if found:
                item_id, cropwise_name, unit_id = found
                permanent_cache[cell_0_lower] = {"id": item_id, "name": cropwise_name, "unit_id": unit_id}
            else:
                # 3. ⏳ РЕЖИМ ОЧІКУВАННЯ ТВОЇХ ДІЙ
                print(f"\n⏳ ТМЦ відсутній у Cropwise: '{cell_0}'")
                print(f"   [Кількість: {quantity} | Сума: {total_sum} - {total_sum / quantity:.2f} за од.]")
                print(f"   🛠️  Зроби паузу, зайди в браузер і створи цю картку. Назва МАЄ БУТИ ТОЧНО ТАКОЮ Ж.")
                
                while True:
                    input("   👉 Як тільки створиш картку в браузері, натисни [ENTER] для перевірки... ")
                    
                    # Скрипт знову шукає в онлайні
                    refound = find_item_in_cropwise_online(cell_0)
                    if refound:
                        item_id, cropwise_name, unit_id = refound
                        print(f"   ✅ Знайшов! Новий ID в системі: {item_id}. Продовжую імпорт...")
                        
                        # Записуємо в кеш, щоб наступного разу взагалі не питати
                        permanent_cache[cell_0_lower] = {"id": item_id, "name": cropwise_name, "unit_id": unit_id}
                        save_permanent_cache(permanent_cache)
                        break
                    else:
                        print("   ❌ Сервер Cropwise досі не бачить товар з такою назвою. Перевір, чи правильно вказав ім'я на сайті.")
                        print("      Якщо хочеш пропустити перевірку і ввести ID вручну, просто введи число-ID, або натисни Enter щоб спробувати знайти знову.")
                        backup_input = input("      Твій вибір (ID або Enter): ").strip()
                        if backup_input.isdigit():
                            item_id = int(backup_input)
                            cropwise_name = f"Ручний ID {item_id}"
                            unit_id = 1
                            permanent_cache[cell_0_lower] = {"id": item_id, "name": cell_0, "unit_id": unit_id}
                            save_permanent_cache(permanent_cache)
                            break

        attempt_count += 1
        total_cents = int(round(total_sum * 100))

        item_payload = {
            "data": {
                "wh_document_id": int(doc_id),
                "wh_item_id": int(item_id),
                "quantity": float(quantity),
                "position": int(attempt_count),
                "total_cents": total_cents,
                "unit_id": unit_id,
                "wh_unit_id": unit_id,
            }
        }
        api_request("POST", API_ITEMS_URL, headers=get_headers(), json=item_payload, timeout=10)

    save_permanent_cache(permanent_cache)
    print(f"\n🏆 РОБОТУ ЗАВЕРШЕНО! Документ повністю сформовано в Cropwise ID: {doc_id}")

if __name__ == "__main__":
    run_waiting_import()