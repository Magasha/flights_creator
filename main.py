"""
Главный скрипт: создание рейсов на yamagistrali.ru из Excel файла.

Роль пользователя: ИСПОЛНИТЕЛЬ (executor).

Алгоритм:
  1. Читаем Excel — список заказов + данные внешнего исполнителя
  2. Получаем все доступные задачи через getForFlightAddForExecutor
  3. Сопоставляем cargoId из Excel с taskName из API
  4. Показываем превью и ждём подтверждения
  5. Для каждой строки Excel:
       а. Создаём рейс
       б. Добавляем задачу в рейс (по taskId)
       в. Назначаем внешнего исполнителя ("Повезу не я")
"""

import logging
import sys

import config
from api_client import MagistraliClient, MagistraliAPIError
from excel_reader import read_orders_excel

logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


def print_separator(char="=", width=60):
    print(char * width)


def main():
    print_separator()
    print("  Создание рейсов — yamagistrali.ru")
    print_separator()
    print()

    # --- Читаем Excel ---
    excel_path = config.EXCEL_PATH
    print(f"  Читаю файл: {excel_path}")
    rows = read_orders_excel(excel_path)

    if not rows:
        print("  Файл пустой или не содержит корректных строк. Выход.")
        sys.exit(1)

    cargo_ids = [r.order_id for r in rows]
    print(f"  Строк в файле:   {len(rows)}")
    print(f"  Уникальных заказов: {len(set(cargo_ids))}")
    print()

    # --- Подключаемся к API ---
    client = MagistraliClient(token=config.YANDEX_OAUTH_TOKEN, base_url=config.BASE_URL)

    # --- Получаем доступные задачи ---
    print("  Запрашиваю доступные задачи через API...")
    try:
        task_map = client.find_tasks_by_cargo_ids(config.EXECUTOR_ID, cargo_ids)
    except MagistraliAPIError as e:
        print(f"\n  ОШИБКА при обращении к API: {e}")
        sys.exit(1)

    found_ids = set(task_map.keys())
    missing_ids = [cid for cid in cargo_ids if cid not in found_ids]

    print()
    print_separator("-")
    print(f"  Найдено в системе:    {len(found_ids)} из {len(rows)}")
    print(f"  Не найдено:           {len(missing_ids)}")
    print_separator("-")

    if missing_ids:
        print("\n  Не найдены следующие заказы:")
        for cid in missing_ids:
            print(f"    - {cid}")
        print()

    if not found_ids:
        print("  Ни один заказ не найден. Проверьте EXECUTOR_ID в config.py и правильность номеров заказов.")
        sys.exit(1)

    # --- Показываем превью ---
    print("\n  Будут созданы рейсы:")
    print_separator("-")
    ok_rows = [r for r in rows if r.order_id in found_ids]
    for r in ok_rows:
        print(f"  {r.order_id}  →  рейс '{r.flight_name}'  |  {r.company_name} ({r.inn})")
    print_separator("-")
    print()

    # --- Подтверждение ---
    try:
        input("  Нажмите Enter для создания рейсов (Ctrl+C для отмены)...")
    except KeyboardInterrupt:
        print("\n\n  Отменено пользователем.")
        sys.exit(0)

    print()

    # --- Создаём рейсы ---
    success_count = 0
    error_count = 0

    for row in ok_rows:
        task = task_map[row.order_id]
        task_id = task["taskId"]

        print(f"  [{row.order_id}] Создаю рейс '{row.flight_name}'...")

        try:
            # Шаг 1: создать рейс
            flight_id = client.create_flight(
                executor_id=config.EXECUTOR_ID,
                flight_name=row.flight_name,
            )
            print(f"    ✓ Рейс создан: {flight_id}")

            # Шаг 2: добавить задачу в рейс
            client.add_tasks_to_flight(flight_id=flight_id, task_ids=[task_id])
            print(f"    ✓ Заказ добавлен в рейс")

            # Шаг 3: назначить внешнего исполнителя
            client.assign_external_executor(
                flight_id=flight_id,
                company_name=row.company_name,
                org_type=row.org_type,
                country=row.country,
                tin=row.inn,
                currency=row.currency,
                price=row.price,
                vat=row.vat,
                price_with_vat=row.price_with_vat,
                is_vat_payer=row.is_vat_payer,
                contact_name=row.contact_name,
                contact_phone=row.contact_phone,
                contact_email=row.contact_email,
                additional_comment=row.comment,
            )
            print(f"    ✓ Внешний исполнитель назначен: {row.company_name}")
            success_count += 1

        except MagistraliAPIError as e:
            print(f"    ✗ ОШИБКА: {e}")
            error_count += 1

        print()

    # --- Итог ---
    print_separator()
    print(f"  Готово!")
    print(f"  Успешно:  {success_count}")
    print(f"  Ошибки:   {error_count}")
    print(f"  Пропущено (не найдены): {len(missing_ids)}")
    print_separator()


if __name__ == "__main__":
    main()
