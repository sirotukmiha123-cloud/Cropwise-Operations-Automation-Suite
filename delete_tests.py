"""
delete_tests.py — швидка зачистка тестових складських документів (WhDocument) за діапазоном ID.

Використовується для видалення документів, створених під час тестування скриптів набору
(наприклад, "oborotka_tmc_3_1.py" чи "cropwise_act.py" в режимі чернетки).
"""

from cropwise_common import API_BASE, get_headers, api_request

# Вкажи діапазон ID документів, які треба стерти (від якого і до якого включно)
START_ID = id
END_ID = id


def clean_test_documents():
    total = END_ID - START_ID + 1
    print(f"🧹 До видалення: {total} документ(ів) з №{START_ID} по №{END_ID} включно.")
    print("⚠️  УВАГА: це НЕЗВОРОТНО видалить складські документи (WhDocument) із зазначеного діапазону!")
    confirm = input("Введіть 'ВИДАЛИТИ' (великими літерами), щоб підтвердити: ").strip()
    if confirm != "ВИДАЛИТИ":
        print("Скасовано.")
        return

    deleted_count = 0
    for doc_id in range(START_ID, END_ID + 1):
        url = f"{API_BASE}/api/v3/wh_documents/{doc_id}"
        try:
            response = api_request("DELETE", url, headers=get_headers())
            if response.status_code in (200, 204):
                print(f"   🗑️ Документ №{doc_id} успішно видалено.")
                deleted_count += 1
            elif response.status_code == 404:
                print(f"   ℹ️ Документ №{doc_id} не знайдено (можливо, вже видалений).")
            else:
                print(f"   ❌ Не вдалося видалити №{doc_id}. Код: {response.status_code}")
        except Exception as e:
            print(f"   ❌ Збій системи на документі №{doc_id}: {e}")

    print(f"\n✅ Зачистка завершена! Всього видалено документів: {deleted_count}")


if __name__ == "__main__":
    clean_test_documents()
