"""
Главный скрипт: читает Excel → создаёт рейсы в yamagistrali.ru

Логика:
1. Читает строки из Excel файла
2. Для каждого заказа проверяет статус (должен быть onExecute — "В работе")
3. Проверяет, что у заказа ещё нет рейса
4. Создаёт рейс
5. Добавляет заказ в рейс
6. Назначает внешнего исполнителя ("Повезу не я")

Запуск:
    python main.py --file orders.xlsx
    python main.py --file orders.xlsx --dry-run
    python main.py --file orders.xlsx --sheet "Мой лист"
"""

import argparse
import logging
import sys
from pathlib import Path

import config
from api_client import MagistraliClient, MagistraliAPIError
from excel_reader import read_excel, OrderRow


# ----------------------------------------------------------------
# Настройка логирования
# ----------------------------------------------------------------
def setup_logging():
    log_level = getattr(logging, config.LOG_LEVEL.upper(), logging.INFO)
    handlers = [logging.StreamHandler(sys.stdout)]
    if config.LOG_TO_FILE:
        handlers.append(logging.FileHandler(config.LOG_FILE, encoding="utf-8"))

    logging.basicConfig(
        level=log_level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=handlers,
    )


logger = logging.getLogger(__name__)


# ----------------------------------------------------------------
# Вспомогательные функции
# ----------------------------------------------------------------
def check_order_has_flight(order: dict) -> bool:
    """
    Проверяет, есть ли у заказа уже созданный рейс.
    В ответе getFlatForCustomer рейсы отражаются в поле flightIds / execOrderId.
    """
    # Проверяем наличие рейсов через поля ответа
    flights = order.get("flightIds") or order.get("flights") or []
    if flights:
        return True
    # Некоторые версии API возвращают execOrderId когда есть исполнение
    # Дополнительная проверка через список рейсов делается в основном цикле
    return False


def process_row(
    client: MagistraliClient,
    row: OrderRow,
    forwarder_id: str,
    dry_run: bool,
) -> dict:
    """
    Обрабатывает одну строку Excel:
    - Проверяет заказ
    - Создаёт рейс
    - Добавляет заказ в рейс
    - Назначает внешнего исполнителя

    Returns:
        dict с ключами: success (bool), flight_id (str|None), message (str)
    """
    order_id = row.order_id
    result = {"success": False, "flight_id": None, "message": ""}

    # --- Шаг 1: получаем заказ и проверяем статус ---
    logger.info("[Строка %d] Обрабатываю заказ %s", row.row_number, order_id)

    try:
        orders = client.get_orders(order_ids=[order_id])
    except MagistraliAPIError as e:
        result["message"] = f"Ошибка получения заказа: {e}"
        logger.error("[Строка %d] %s", row.row_number, result["message"])
        return result

    if not orders:
        result["message"] = f"Заказ {order_id} не найден в системе"
        logger.warning("[Строка %d] %s", row.row_number, result["message"])
        return result

    order = orders[0]
    status = order.get("status", "")

    if status != "onExecute":
        status_names = {
            "draft": "В черновиках",
            "onMatch": "На подборе исполнителя",
            "onExecute": "В работе",
            "completeExecute": "Исполнение завершено",
            "cancel": "Отменён",
            "matchHistory": "История торгов",
            "revoke": "Отзыв из исполнения",
        }
        human_status = status_names.get(status, status)
        result["message"] = (
            f"Заказ {order_id} имеет статус '{human_status}' — ожидается 'В работе' (onExecute). Пропускаю."
        )
        logger.warning("[Строка %d] %s", row.row_number, result["message"])
        return result

    # --- Шаг 2: проверяем, нет ли уже рейса ---
    try:
        existing_flights = client.get_flights_for_orders([order_id])
    except MagistraliAPIError as e:
        logger.warning(
            "[Строка %d] Не удалось проверить существующие рейсы: %s. Продолжаю.", row.row_number, e
        )
        existing_flights = []

    if existing_flights:
        flight_ids = [f.get("id", "?") for f in existing_flights]
        result["message"] = (
            f"У заказа {order_id} уже есть рейс(ы): {flight_ids}. Пропускаю."
        )
        logger.warning("[Строка %d] %s", row.row_number, result["message"])
        return result

    logger.info(
        "[Строка %d] Заказ %s: статус ОК, рейсов нет — создаю рейс '%s'",
        row.row_number, order_id, row.flight_name,
    )

    if dry_run:
        result["success"] = True
        result["flight_id"] = "DRY_RUN"
        result["message"] = (
            f"[DRY RUN] Создан бы рейс '{row.flight_name}' для заказа {order_id} "
            f"с исполнителем '{row.company_name}' (ИНН: {row.inn})"
        )
        logger.info("[Строка %d] %s", row.row_number, result["message"])
        return result

    # --- Шаг 3: создаём рейс ---
    try:
        flight_id = client.create_flight(
            executor_id=forwarder_id,
            flight_name=row.flight_name,
        )
        logger.info("[Строка %d] Рейс создан, ID: %s", row.row_number, flight_id)
    except MagistraliAPIError as e:
        result["message"] = f"Ошибка создания рейса: {e}"
        logger.error("[Строка %d] %s", row.row_number, result["message"])
        return result

    # --- Шаг 4: добавляем заказ в рейс ---
    try:
        client.add_orders_to_flight(flight_id=flight_id, order_ids=[order_id])
        logger.info("[Строка %d] Заказ %s добавлен в рейс %s", row.row_number, order_id, flight_id)
    except MagistraliAPIError as e:
        result["message"] = f"Рейс {flight_id} создан, но ошибка добавления заказа: {e}"
        result["flight_id"] = flight_id
        logger.error("[Строка %d] %s", row.row_number, result["message"])
        return result

    # --- Шаг 5: назначаем внешнего исполнителя ("Повезу не я") ---
    try:
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
        logger.info(
            "[Строка %d] Внешний исполнитель '%s' (ИНН: %s) назначен на рейс %s",
            row.row_number, row.company_name, row.inn, flight_id,
        )
    except MagistraliAPIError as e:
        result["message"] = (
            f"Рейс {flight_id} создан, заказ добавлен, "
            f"но ошибка назначения внешнего исполнителя: {e}"
        )
        result["flight_id"] = flight_id
        result["success"] = False
        logger.error("[Строка %d] %s", row.row_number, result["message"])
        return result

    # --- Успех ---
    result["success"] = True
    result["flight_id"] = flight_id
    result["message"] = (
        f"Рейс '{row.flight_name}' (ID: {flight_id}) создан для заказа {order_id}. "
        f"Внешний исполнитель: {row.company_name} (ИНН: {row.inn})."
    )
    logger.info("[Строка %d] ✓ %s", row.row_number, result["message"])
    return result


# ----------------------------------------------------------------
# Основная функция
# ----------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Создаёт рейсы в Яндекс Магистралях по данным из Excel файла"
    )
    parser.add_argument(
        "--file", "-f",
        required=True,
        help="Путь к Excel файлу (.xlsx)",
    )
    parser.add_argument(
        "--sheet", "-s",
        default=config.EXCEL_SHEET_NAME,
        help=f"Название листа (по умолчанию: {config.EXCEL_SHEET_NAME})",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Режим тестирования: проверяет данные без реального создания рейсов",
    )
    parser.add_argument(
        "--forwarder-id",
        default=config.FORWARDER_ID,
        help="ID экспедитора (переопределяет значение из config.py)",
    )
    args = parser.parse_args()

    setup_logging()

    # ---- Проверка конфигурации ----
    if not config.YANDEX_OAUTH_TOKEN or config.YANDEX_OAUTH_TOKEN.startswith("y0_AgAAAA"):
        logger.error(
            "Токен не настроен! Укажите YANDEX_OAUTH_TOKEN в config.py или "
            "переменной окружения YANDEX_TOKEN."
        )
        sys.exit(1)

    forwarder_id = args.forwarder_id
    if not forwarder_id:
        logger.error(
            "FORWARDER_ID не указан! Заполните его в config.py или передайте через --forwarder-id."
        )
        sys.exit(1)

    if args.dry_run:
        logger.info("=" * 60)
        logger.info("РЕЖИМ DRY RUN — реальные запросы создания НЕ выполняются")
        logger.info("=" * 60)

    # ---- Читаем Excel ----
    try:
        rows = read_excel(args.file, sheet_name=args.sheet)
    except (FileNotFoundError, ValueError) as e:
        logger.error("Ошибка чтения файла: %s", e)
        sys.exit(1)

    if not rows:
        logger.warning("Файл не содержит строк с данными. Нечего обрабатывать.")
        sys.exit(0)

    # ---- Проверяем строки на ошибки ----
    invalid_rows = [r for r in rows if not r.is_valid]
    valid_rows = [r for r in rows if r.is_valid]

    if invalid_rows:
        logger.warning("Найдено %d строк с ошибками валидации (будут пропущены):", len(invalid_rows))
        for row in invalid_rows:
            logger.warning("  Строка %d (заказ '%s'): %s", row.row_number, row.order_id, "; ".join(row.errors))

    logger.info("Всего строк для обработки: %d", len(valid_rows))

    if not valid_rows:
        logger.error("Нет корректных строк для обработки. Выход.")
        sys.exit(1)

    # ---- Создаём клиент API ----
    client = MagistraliClient(
        token=config.YANDEX_OAUTH_TOKEN,
        base_url=config.BASE_URL,
    )

    # ---- Обрабатываем строки ----
    success_count = 0
    skip_count = 0
    error_count = 0

    for row in valid_rows:
        result = process_row(
            client=client,
            row=row,
            forwarder_id=forwarder_id,
            dry_run=args.dry_run,
        )
        if result["success"]:
            success_count += 1
        elif "Пропускаю" in result["message"] or "уже есть рейс" in result["message"]:
            skip_count += 1
        else:
            error_count += 1

    # ---- Итог ----
    logger.info("")
    logger.info("=" * 60)
    logger.info("ИТОГ:")
    logger.info("  ✓ Успешно создано рейсов:  %d", success_count)
    logger.info("  ~ Пропущено (уже есть/статус): %d", skip_count)
    logger.info("  ✗ Ошибок:                  %d", error_count)
    logger.info("  ⚠ Строк с ошибками Excel:  %d", len(invalid_rows))
    logger.info("=" * 60)

    if error_count > 0:
        sys.exit(2)


if __name__ == "__main__":
    main()
