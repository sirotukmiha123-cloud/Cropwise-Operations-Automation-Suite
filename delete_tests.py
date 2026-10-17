"""
delete_tests.py — зачистка складських документів (WhDocument) за діапазоном ID.

Використовується для відкату тестових документів, створених скриптами набору
(наприклад, "oborotka_tmc_3_1.py" чи "cropwise_act.py" в режимі чернетки).

Приклади:
    python delete_tests.py --start 100 --end 110 --dry-run   # лише показати, що буде видалено
    python delete_tests.py --start 100 --end 110             # видалити після підтвердження
"""
import argparse

CONFIRM_WORD = "ВИДАЛИТИ"
LARGE_RANGE = 100  # від цієї кількості документів просимо додаткове підтвердження


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Видаляє складські документи Cropwise за діапазоном ID (включно)."
    )
    parser.add_argument("--start", type=int, required=True, help="перший ID документа")
    parser.add_argument("--end", type=int, required=True, help="останній ID документа (включно)")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="нічого не видаляти й не звертатись до сервера, лише показати діапазон",
    )
    args = parser.parse_args(argv)
    if args.start < 1 or args.end < args.start:
        parser.error("має виконуватись умова: 1 <= --start <= --end")
    return args


def confirm_deletion(total):
    """Повертає True, лише якщо користувач явно підтвердив видалення."""
    print("⚠️ УВАГА: це НЕЗВОРОТНО видалить складські документи (WhDocument) із зазначеного діапазону!")
    answer = input(f"Введіть '{CONFIRM_WORD}' (великими літерами), щоб підтвердити: ").strip()
    if answer != CONFIRM_WORD:
        return False
    if total > LARGE_RANGE:
        answer = input(f"Діапазон великий. Введіть кількість документів ({total}) ще раз: ").strip()
        if answer != str(total):
            return False
    return True


def clean_test_documents(start_id, end_id, dry_run=False):
    total = end_id - start_id + 1
    print(f"🧹 До видалення: {total} документ(ів) з №{start_id} по №{end_id} включно.")

    if dry_run:
        print("ℹ️ Режим --dry-run: запитів до сервера не надсилається, нічого не видалено.")
        return 0

    if not confirm_deletion(total):
        print("Скасовано.")
        return 0

    # Імпорт тут, а не нагорі: спільний модуль запитує токен під час імпорту,
    # а для --dry-run токен не потрібен.
    from cropwise_common import API_BASE, get_headers, api_request

    deleted_count = 0
    for doc_id in range(start_id, end_id + 1):
        url = f"{API_BASE}/api/v3/wh_documents/{doc_id}"
        try:
            response = api_request("DELETE", url, headers=get_headers())
            if response.status_code in (200, 204):
                print(f" 🗑️ Документ №{doc_id} успішно видалено.")
                deleted_count += 1
            elif response.status_code == 404:
                print(f" ℹ️ Документ №{doc_id} не знайдено (можливо, вже видалений).")
            else:
                print(f" ❌ Не вдалося видалити №{doc_id}. Код: {response.status_code}")
        except Exception as e:
            print(f" ❌ Збій системи на документі №{doc_id}: {e}")

    print(f"\n✅ Зачистка завершена! Всього видалено документів: {deleted_count}")
    return deleted_count


if __name__ == "__main__":
    args = parse_args()
    clean_test_documents(args.start, args.end, dry_run=args.dry_run)
