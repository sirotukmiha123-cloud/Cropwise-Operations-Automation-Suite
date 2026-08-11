import pandas as pd
import json
import os
import re
from docx import Document
from cropwise_common import API_BASE, get_headers, api_request, clean_file_path as clean_path

JSON_FILE_NAME = "nomenclature.json"

API_DOCS_URL = f"{API_BASE}/api/v3/wh_documents"
API_ITEMS_URL = f"{API_BASE}/api/v3/wh_document_items"
API_FIELDS_URL = f"{API_BASE}/api/v3/fields"

EXCEL_COLS = {
    'G': 6, 'H': 7, 'I': 8, 'J': 9,
    'K': 10, 'L': 11, 'M': 12, 'N': 13,
    'O': 14, 'P': 15, 'Q': 16, 'R': 17,
    'S': 18, 'T': 19, 'U': 20, 'V': 21
}

def load_cropwise_fields():
    print("🔄 Завантажую список полів та культур з Cropwise...")
    fields_list = []
    url = f"{API_FIELDS_URL}?calculated_status=active&page=1&per_page=100"
    try:
        res = api_request("GET", url, headers=get_headers(), timeout=15)
        if res.status_code == 200:
            for field in res.json().get('data', []):
                crop_name = field.get('current_crop_name')
                crop_str = f"({crop_name})" if crop_name else "(Культуру не вказано)"
                
                fields_list.append({
                    "id": field.get('id'), 
                    "name": field.get('name', '').strip(),
                    "crop": crop_str
                })
            fields_list = sorted(fields_list, key=lambda x: x['name'])
    except Exception as e:
        print(f" ⚠️ Помилка завантаження полів: {e}")
    return fields_list


def load_fields_from_excel(filepath="ref_fields.json"):
    if not os.path.exists(filepath):
        print(f"⚠️  Файл полів '{filepath}' не знайдено.")
        return []
    try:
        ext = os.path.splitext(filepath)[1].lower()
        if ext in (".xlsx", ".xls"):
            df = pd.read_excel(filepath, header=0)
            result = []
            for _, row in df.iterrows():
                fid   = row.iloc[0]
                fname = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ""
                if fname and fname != "nan" and str(fid).isdigit() if isinstance(fid, str) else True:
                    try:
                        result.append({"id": int(fid), "name": fname})
                    except:
                        pass
            print(f"  📋 Завантажено полів з Excel: {len(result)}")
            return result
        else:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return [{"id": d.get("id"), "name": d.get("name","")} for d in data]
            return []
    except Exception as e:
        print(f"❌ Помилка читання файлу полів: {e}")
        return []

def load_local_nomenclature():
    if not os.path.exists(JSON_FILE_NAME): 
        print(f"❌ Помилка: Файл {JSON_FILE_NAME} не знайдено!")
        return {}
    try:
        with open(JSON_FILE_NAME, "r", encoding="utf-8") as f:
            raw_data = json.load(f)
            return {str(k).lower().strip(): (v.get("id") if isinstance(v, dict) else v) for k, v in raw_data.items()}
    except Exception as e: 
        print(f"❌ Помилка зчитування JSON: {e}")
        return {}

def process_single_field(df, nomenclature_items, field_id, field_name, qty_col_letter, user_act_number, start_row):
    qty_idx = EXCEL_COLS[qty_col_letter]
    sum_idx = qty_idx + 2  
    
    print(f"\n📦 Створюю чернетку в Cropwise для поля: {field_name} (Рядок початку: {start_row}, Стовпчик: {qty_col_letter})...")
    additional_info_text = f"Акт №{user_act_number}. Автоімпорт Поля. Об'єкт: {field_name}"
    
    doc_payload = {
        "data": {
            "document_type": "goods_issue",
            "document_status": "draft",
            "document_time": "2026-05-31T23:59:00.000+03:00",
            "from_warehouse_id": id,
            "additional_info": additional_info_text
        }
    }
    
    if field_id:
        doc_payload["data"]["consumption_target_type"] = "Field"
        doc_payload["data"]["consumption_target_id"] = int(field_id)

    try:
        doc_res = api_request("POST", API_DOCS_URL, headers=get_headers(), json=doc_payload, timeout=15)
        if doc_res.status_code not in [200, 201]:
            print(f"❌ Помилка створення документу: {doc_res.text}")
            return
        doc_id = doc_res.json()['data']['id']
        print(f"✅ Чернетку акту №{doc_id} успішно створено!")
    except Exception as e:
        print(f"❌ Помилка API: {e}")
        return

    idx_pos = 1
    start_index = max(0, start_row - 1)
    
    for idx in range(start_index, len(df)):
        if idx >= len(df): break
        row = df.iloc[idx]
        if pd.isna(row.iloc[0]) or row.iloc[0] is None: continue
        
        name = str(row.iloc[0]).strip()
        if "всього" in name.lower() or "разом" in name.lower() or "голова" in name.lower():
            break
            
        item_id = nomenclature_items.get(name.lower())
        if not item_id:
            for nom_name, nom_id in nomenclature_items.items():
                if nom_name in name.lower() or name.lower() in nom_name:
                    item_id = nom_id
                    break
        
        forced_helper = False
        if not item_id:
            item_id = id
            display_description = f"- {name}"
            forced_helper = True
        else:
            display_description = name

        try:
            val_qty = row.iloc[qty_idx]
            val_sum = row.iloc[sum_idx]
            qty = float(str(val_qty).replace(",", ".")) if not pd.isna(val_qty) else 0.0
            total_sum = float(str(val_sum).replace(",", ".")) if not pd.isna(val_sum) else 0.0
        except:
            continue
            
        if qty > 0:
            active_unit_id = 1
            if not forced_helper:
                try:
                    item_info_res = api_request("GET", f"{API_BASE}/api/v3/wh_items/{item_id}", headers=get_headers(), timeout=10)
                    if item_info_res.status_code == 200:
                        item_data = item_info_res.json().get("data", {})
                        active_unit_id = item_data.get("wh_item_base_unit", {}).get("id") or item_data.get("base_unit", {}).get("id") or 1
                except: pass

            item_payload = {
                "data": {
                    "wh_document_id": int(doc_id),
                    "wh_item_id": int(item_id),
                    "sku_id": int(item_id),
                    "quantity": float(qty),
                    "position": idx_pos,
                    "total_cents": int(round(total_sum * 100)),
                    "unit_id": int(active_unit_id),
                    "wh_unit_id": int(active_unit_id),
                    "description": f"{display_description}"
                }
            }
            
            item_res = api_request("POST", API_ITEMS_URL, headers=get_headers(), json=item_payload, timeout=15)
            if item_res.status_code not in [200, 201] and not forced_helper:
                print(f"   ⚠️ Товар [{name[:20]}] заблоковано. Страховка ID 169...")
                item_payload["data"]["wh_item_id"] = 169
                item_payload["data"]["sku_id"] = 169
                item_payload["data"]["unit_id"] = 1
                item_payload["data"]["wh_unit_id"] = 1
                item_payload["data"]["description"] = f"- {name}"
                item_res = api_request("POST", API_ITEMS_URL, headers=get_headers(), json=item_payload, timeout=15)

            if item_res.status_code in [200, 201]:
                print(f"     ➔ ➕ [{idx_pos}] Додано: {name[:25]}... ➔ {qty} од.")
                idx_pos += 1

    print(f"🎉 Акт №{doc_id} для поля [{field_name}] повністю укомплектовано!")


# ==================== МОДУЛЬ: АКТ СПИСАННЯ НАСІННЯ (WORD) ====================

def parse_word_docx_file(docx_path):
    """Зчитує сирі дані таблиці акту насіння з Word за нашими індексами."""
    try:
        doc = Document(docx_path)
        if not doc.tables:
            print("❌ У документі не знайдено жодної таблиці.")
            return []

        data_table = max(doc.tables, key=lambda t: len(t.columns))
        HEADER_ROWS = 3   
        rows = []

        for i, row in enumerate(data_table.rows):
            if i < HEADER_ROWS: continue

            cells = [c.text.strip() for c in row.cells]
            while len(cells) < 11: cells.append("")

            name    = cells[0]
            sort_   = cells[1]
            field_s = cells[2]   

            if not name: continue
            if any(kw in name.lower() for kw in ("всього", "разом", "голова", "підпис", "додаток")): continue
            if any(kw in (field_s or "").lower() for kw in ("всього", "разом")): continue

            try:
                qty = float(cells[9].replace(",", ".").replace(" ", "").replace("\xa0","")) if cells[9] else 0.0
            except: qty = 0.0
            try:
                price = float(cells[10].replace(",", ".").replace(" ", "").replace("\xa0","")) if cells[10] else 0.0
            except: price = 0.0

            if qty <= 0: continue

            full_name = name
            if sort_ and sort_.lower() not in name.lower():
                full_name = f"{name} ({sort_})"

            rows.append({
                "name":    full_name,
                "field_s": field_s,
                "qty":     qty,
                "price":   price,
            })
        return rows
    except Exception as e:
        print(f"❌ Помилка читання Word-файлу: {e}")
        return []


def create_word_wh_document(field_id, field_name, act_number, doc_date):
    """Створює чернетку wh_document для Ворд-модуля."""
    payload = {
        "data": {
            "document_type":    "goods_issue",
            "document_status":  "draft",
            "document_time":    f"{doc_date}T23:59:00.000+03:00",
            "from_warehouse_id": 13,
            "additional_info":  f"Акт №{act_number}. Автоімпорт. Насіння. Поле: {field_name}"
        }
    }
    if field_id:
        payload["data"]["consumption_target_type"] = "Field"
        payload["data"]["consumption_target_id"]   = int(field_id)
    try:
        res = api_request("POST", API_DOCS_URL, headers=get_headers(), json=payload, timeout=15)
        if res.status_code in [200, 201]:
            return res.json()["data"]["id"]
        return None
    except:
        return None


def send_single_word_item(doc_id, item, nomenclature_items, position):
    """Відправляє одну позицію насіння до документу doc_id з фолбеком на ID id."""
    name  = item["name"]
    qty   = item["qty"]
    price = item["price"]
    total_cents = int(round(qty * price * 100))

    item_id = nomenclature_items.get(name.lower().strip())
    if not item_id:
        for nom_name, nom_id in nomenclature_items.items():
            if nom_name in name.lower() or name.lower() in nom_name:
                item_id = nom_id
                break

    forced_helper = False
    if not item_id:
        item_id = id
        display_name = f"- {name}"
        forced_helper = True
    else:
        display_name = name

    active_unit_id = id
    if not forced_helper:
        try:
            r = api_request("GET", f"{API_BASE}/api/v3/wh_items/{item_id}", headers=get_headers(), timeout=10)
            if r.status_code == 200:
                d = r.json().get("data", {})
                active_unit_id = d.get("wh_item_base_unit", {}).get("id") or d.get("base_unit", {}).get("id") or 1
        except: pass

    payload = {
        "data": {
            "wh_document_id": int(doc_id),
            "wh_item_id":     int(item_id),
            "sku_id":         int(item_id),
            "quantity":       float(qty),
            "position":       position,
            "total_cents":    total_cents,
            "unit_id":        int(active_unit_id),
            "wh_unit_id":     int(active_unit_id),
            "description":    display_name,
        }
    }

    try:
        res = api_request("POST", API_ITEMS_URL, headers=get_headers(), json=payload, timeout=15)
        if res.status_code not in [200, 201] and not forced_helper:
            print(f"    ⚠️ [{name[:20]}] заблоковано — страховка на ID 169...")
            payload["data"].update({
                "wh_item_id": 169, "sku_id": 169, "unit_id": 1, "wh_unit_id": 1, "description": f"- {name}"
            })
            res = api_request("POST", API_ITEMS_URL, headers=get_headers(), json=payload, timeout=15)

        if res.status_code in [200, 201]:
            print(f"    ➔ ➕ [{position}] {name[:30]}... → {qty} од. успішно додано!")
            return True
        return False
    except:
        return False


def process_word_act_interactive(docx_path, nomenclature_items, ref_fields):
    """🔥 Інтерактивна логіка з ДИНАМІЧНИМ ПОШУКОМ полів під час вибору."""
    print("\n" + "="*60)
    print("🌱 МОДУЛЬ: ІНТЕРАКТИВНЕ СПИСАННЯ НАСІННЯ (WORD)")
    print("="*60)

    rows = parse_word_docx_file(docx_path)
    if not rows:
        print("❌ Не вдалося зчитати дані з Word-файлу.")
        return

    act_number = input("\n📝 Введіть номер акту (для опису документу): ").strip()
    doc_date   = input("📅 Дата акту (Enter = 2026-05-31): ").strip()
    if not doc_date: doc_date = "2026-05-31"

    current_doc_id = None
    current_field_id = None
    current_field_name = ""
    position_counter = 1

    total_sent = 0
    total_skip = 0

    print(f"\n📊 Починаю покрокову обробку. Знайдено {len(rows)} позицій насіння.")

    for idx, item in enumerate(rows, start=1):
        print(f"\n{'─'*60}")
        print(f"📋 ТОВАР [{idx}/{len(rows)}]: \033[1;32m{item['name']}\033[0m ➔ Кількість: \033[1;33m{item['qty']}\033[0m од.")
        print(f"📌 Пропозиція з Ворду: '{item['field_s']}'")

        selected_field = None

        # Цикл пошуку поля, поки користувач не зробить правильний вибір
        while True:
            search_query = input("\n🔍 Введіть назву поля для пошуку (або Enter для автопошуку, 0 - пропустити, E - вихід): ").strip()

            if search_query.lower() == 'e':
                print("⏹️  Обробку перервано.")
                return
            if search_query == '0':
                print(f"⏭️  Позицію [{item['name']}] пропущено.")
                total_skip += 1
                break

            # Якщо користувач просто тицьнув Enter — беремо назву поля з Ворду як запит
            if not search_query:
                # Очищаємо "№ 5 Хм" -> "5 хм"
                search_query = item['field_s'].replace("№", "").replace("No", "").replace("#", "").strip()
                print(f"🤖 Автопошук за запитом: '{search_query}'")

            # Фільтруємо наш локальний список полів по введеному тексту
            filtered_fields = [f for f in ref_fields if search_query.lower() in f['name'].lower()]

            if not filtered_fields:
                print("❌ Полів із такою назвою не знайдено! Спробуйте інший запит (наприклад: '5 хм', '1к').")
                continue

            # Виводимо тільки знайдені поля
            print("\n🗺️  ЗНАЙДЕНІ ПОЛЯ:")
            for i, f in enumerate(filtered_fields, start=1):
                print(f"  [{i}] {f['name']}")
            print("  [0] Спробувати знайти інше поле")

            sub_choice = input("\n👉 Твій вибір (номер зі списку знайдених або 0): ").strip()

            if sub_choice == '0':
                continue

            try:
                f_idx = int(sub_choice) - 1
                if 0 <= f_idx < len(filtered_fields):
                    selected_field = filtered_fields[f_idx]
                    break
                else:
                    print("❌ Невірний номер у підсписку!")
            except:
                print("❌ Будь ласка, введіть число.")

        # Якщо позицію було пропущено (choice == 0), переходимо до наступного товару
        if not selected_field:
            continue

        target_field_id = selected_field["id"]
        target_field_name = selected_field["name"]

        # ПЕРЕВІРКА ПАМ'ЯТІ ПОЛЯ (ОБ'ЄДНАННЯ В ОДИН АКТ):
        if target_field_id == current_field_id and current_doc_id is not None:
            print(f"♻️  Поле те саме [{target_field_name}]. Додаю в існуючу чернетку #{current_doc_id}...")
        else:
            print(f"✨ Нове поле. Створюю новий акт вибуття для [{target_field_name}]...")
            new_doc_id = create_word_wh_document(target_field_id, target_field_name, act_number, doc_date)
            if new_doc_id:
                current_doc_id = new_doc_id
                current_field_id = target_field_id
                current_field_name = target_field_name
                position_counter = 1
                print(f"  ✅ Чернетку документу #{current_doc_id} успішно створено!")
            else:
                print("❌ Не вдалося створити документ у Cropwise. Пропускаю.")
                total_skip += 1
                continue

        success = send_single_word_item(current_doc_id, item, nomenclature_items, position_counter)
        if success:
            total_sent += 1
            position_counter += 1
        else:
            total_skip += 1

    print("\n" + "="*60)
    print(f"🎉 Обробку Ворд-акту завершено! Успішно додано: {total_sent}, пропущено: {total_skip}")
    print("="*60)


def main():
    AGRI_ITEMS = load_local_nomenclature()
    
    # 🛰️ ТЯГНЕМО ПОЛЯ НАПРЯМУ З СЕРВЕРА CROPWISE
    FIELDS = load_cropwise_fields()
    
    print("\n" + "="*60)
    print("📥 ПЕРЕТЯГНИ ФАЙЛ (EXCEL або WORD .docx) І НАТИСНИ ENTER")
    print("="*60)
    file_path = clean_path(input("👉 Шлях до файлу: "))

    if not os.path.exists(file_path):
        print("❌ Файл не знайдено.")
        return

    base_name = os.path.basename(file_path)  # Отримуємо назву файлу
    ext = os.path.splitext(file_path)[1].lower()
    
    # 🔥 НОВИЙ ЗАХИСТ ВІД ДУБЛЮВАННЯ ТА ПОВТОРНОГО ЗАПУСКУ
    if base_name.startswith("✅_"):
        print("\n" + "!"*60)
        print(f"⚠️  УВАГА! Файл [{base_name}] уже замаркований як ВИКОНАНИЙ!")
        print("!"*60)
        print(" [1] Проігнорувати і все одно обробити (запустити з нуля)")
        print(" [2] Скасувати, я помилився файлом")
        
        double_check = input("\n👉 Твій вибір (1 або 2): ").strip()
        if double_check != "1":
            print("👋 Обробку скасовано. Робот захистив тебе від дублювання!")
            return
        print("🔄 Окей, запускаємо повторну обробку за твоїм запитом...")

    # Прапорець, який покаже, чи успішно пройшла обробка для перейменування
    processing_success = False
    
    # ── ПЕРЕХІД НА ВОРД (НАСІННЯ) ───────────────────────────
    if ext == ".docx":
        if not FIELDS:
            print("❌ Список полів порожній! Перевір інтернет або токен.")
            return
        print(f"🛰️  Успішно підтягнуто {len(FIELDS)} живих полів з сервера Cropwise!")
        process_word_act_interactive(file_path, AGRI_ITEMS, FIELDS)
        processing_success = True  # Ворд оброблено
        
    # ── СТАРА EXCEL-ЛОГІКА (ДОБРИВА / ЗЗР) ───────────────────
    else:
        df = pd.read_excel(file_path, header=None)
        user_act_number = input("📝 Введіть реальний номер акту для фільтру (наприклад, 524): ").strip()

        while True:
            print("\n%s" % ("🗺️  ОБЕРИ ПОЛЕ ДЛЯ ОБРОБКИ:"))
            for i, f in enumerate(FIELDS, start=1):
                if i <= 45:
                    print(f"  [{i:<2}] {f['name']} {f['crop']:<40} (ID: {f['id']})")
            print("  [0] Закрити програму й вийти")
            
            try:
                choice = input("\n👉 Введіть номер поля з консолі: ").strip()
                if choice == "0": 
                    print("👋 Роботу завершено. Відпочивай!")
                    break
                    
                idx = int(choice) - 1
                field_id = FIELDS[idx]["id"]
                field_name = FIELDS[idx]["name"]
            except:
                print("❌ Невірний вибір номера поля. Спробуй ще раз.")
                continue
                
            while True:
                user_input = input(f"👉 Введіть стовпчик для поля [{field_name}] (наприклад, H або H52): ").strip().upper()
                match = re.match(r"^([A-Z]+)(\d*)$", user_input)
                if match:
                    col_letter = match.group(1)
                    row_str = match.group(2)
                    if col_letter in EXCEL_COLS:
                        start_row = int(row_str) if row_str else 30  
                        break
                print("❌ Невірний формат стовпчика!")

            process_single_field(df, AGRI_ITEMS, field_id, field_name, col_letter, user_act_number, start_row)
            processing_success = True  # Ексель успішно обробив хоча б одне поле
            
            print("\n" + "-"*50)
            print("🔄 Бажаєш обробити ще одне поле з цього ж Excel-файлу?")
            print("  [1] Так, обробити наступне поле")
            print("  [2] Ні, закінчити роботу з цим файлом")
            
            next_step = input("\n👉 Твій вибір (1 або 2): ").strip()
            if next_step != '1':
                print("🏁 Усі акти розкидані по окремих чернетках. Роботу закінчено!")
                break

    # 🔥 УНІВЕРСАЛЬНИЙ БЛОК АВТО-МАРКУВАННЯ ДЛЯ ВСІХ ФАЙЛІВ (і Word, і Excel)
    if processing_success:
        try:
            folder = os.path.dirname(file_path)      # Папка файлу
            
            # Перевіряємо ще раз, чи файл уже не має галочки, щоб не дублювати "✅_✅_"
            if not base_name.startswith("✅_"):
                new_base_name = f"✅_{base_name}"
                new_file_path = os.path.join(folder, new_base_name)
                
                os.rename(file_path, new_file_path)
                print(f"\n💾 Файл автоматично марковано як виконаний:")
                print(f"   ➔ {new_base_name}")
        except Exception as e:
            print(f"\n⚠️  Не вдалося автоматично перейменувати файл: {e}")
if __name__ == "__main__":
    main()
