import json
import os
import pandas as pd
from cropwise_common import API_BASE, get_headers, api_request

# =====================================================
# ФУНКЦІЇ ДЛЯ РОБОТИ З ГРУПАМИ ТА ПІДТИПАМИ
# =====================================================
def fetch_work_type_groups():
    """Завантажує всі існуючі групи робіт з сервера"""
    try:
        resp = api_request("GET", f"{API_BASE}/api/v3/work_type_groups", headers=get_headers(), params={"limit": 100}, timeout=10)
        if resp.status_code == 200:
            return resp.json().get("data", [])
    except Exception as e:
        print(f"⚠️ Помилка завантаження груп робіт: {e}")
    return []

def create_new_group(group_name):
    """Створює нову групу робіт через API (Режим 00)"""
    payload = {"data": {"name": group_name}}
    try:
        resp = api_request("POST", f"{API_BASE}/api/v3/work_type_groups", headers=get_headers(), json=payload, timeout=10)
        if resp.status_code in [200, 201]:
            new_group = resp.json().get("data", {})
            print(f"✅ Створено нову групу: '{group_name}' (ID: {new_group.get('id')})")
            return new_group
        else:
            print(f"❌ Не вдалося створити групу. Код: {resp.status_code}")
    except Exception as e:
        print(f"❌ Помилка створення групи: {e}")
    return None

def upload_work_type(name, kbs_code, group_id):
    """Створює підтип роботи у вказаній групі"""
    payload = {
        "data": {
            "name": name,
            "external_id": str(kbs_code),  # Надсилаємо як рядок з усіма нулями
            "work_type_group_id": int(group_id)
        }
    }
    try:
        resp = api_request("POST", f"{API_BASE}/api/v3/work_types", headers=get_headers(), json=payload, timeout=10)
        if resp.status_code in [200, 201]:
            print(f"  🚀 Успішно завантажено: {name} -> ID: {resp.json().get('data', {}).get('id')}")
            return True
        elif resp.status_code == 422:
            print(f"  ℹ️ Робота з кодом {kbs_code} вже існує в Cropwise.")
            return True
        else:
            print(f"  ❌ Помилка сервера ({resp.status_code}): {resp.text}")
    except Exception as e:
        print(f"  ❌ Помилка мережі при завантаженні: {e}")
    return False

# =====================================================
# ГОЛОВНИЙ СЦЕНАРІЙ
# =====================================================
def main():
    print("=" * 60)
    print("🚜 ІНТЕРАКТИВНИЙ ЗАВАНТАЖУВАЧ ДОВІДНИКА КБС В CROPWISE")
    print("=" * 60)

    excel_file = "Список робіт.xlsx"
    if not os.path.exists(excel_file):
        print(f"❌ Файл {excel_file} не знайдено в поточній папці!")
        return

    # Зчитуємо файл, примусово переводячи Код у формат рядка (str), щоб зберегти нулі
    df = pd.read_excel(excel_file, dtype={"Код": str})

    # Отримуємо поточні групи з сервера
    groups = fetch_work_type_groups()
    if not groups:
        print("❌ Не вдалося отримати групи робіт з сервера. Перевірте токен або мережу.")
        return

    print(f"📊 Знайдено {len(df)} рядків у файлі Excel.")
    start_row = input("▶️ З якого за рахунком рядка почати? (натисніть Enter для старту з 1): ").strip()
    start_idx = int(start_row) - 1 if start_row.isdigit() else 0

    for idx, row in df.iloc[start_idx:].iterrows():
        raw_name = str(row.iloc[0]).strip()
        raw_code = str(row.iloc[1]).strip()

        # Повертаємо провідні нулі, якщо Excel їх десь затер (робимо довжину 9 знаків)
        if raw_code.isdigit():
            kbs_code = raw_code.zfill(9)
        else:
            kbs_code = raw_code

        print("\n" + "-" * 60)
        print(f"📋 Позиція №{idx + 1} з {len(df)}")
        print(f"🌾 Робота з КБС:  {raw_name}")
        print(f"🔢 Код КБС:       {kbs_code}")
        print("-" * 60)

        while True:
            print("\n📁 ДОСТУПНІ ГРУПИ В CROPWISE:")
            for i, g in enumerate(groups):
                print(f"  {i+1:>2} → {g.get('name')} (ID: {g.get('id')})")
            print("  00 → Створити нову групу")
            print("  11 → Пропустити цю роботу")

            choice = input("\n👉 Твій вибір: ").strip()

            if choice == "11":
                print("⏭️  Позицію пропущено.")
                break
            
            elif choice == "00":
                new_name = input("✍️  Введіть назву для нової групи: ").strip()
                if new_name:
                    new_g = create_new_group(new_name)
                    if new_g:
                        groups = fetch_work_type_groups()  # Оновлюємо список груп з сервера
                continue

            elif choice.isdigit() and 1 <= int(choice) <= len(groups):
                selected_group = groups[int(choice) - 1]
                g_id = selected_group.get("id")
                
                # Відправка в Cropwise
                success = upload_work_type(raw_name, kbs_code, g_id)
                if success:
                    break
                else:
                    retry = input("⚠️  Помилка відправки. Спробувати ще раз? (1 - так, інше - пропустити): ").strip()
                    if retry != "1":
                        break
            else:
                print("⚠️  Невірний вибір. Введіть номер групи, 00 або 11.")

    print("\n🏁 Обробка списку завершена!")

if __name__ == "__main__":
    main()