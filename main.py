"""
Главный скрипт: читает Excel → создаёт рейсы в yamagistrali.ru

Логика:
1. Читает строки из Excel файла
2. Запрашивает все заказы из списка одним пакетным запросом
3. Выводит preview: что найдено, статусы, у каких уже есть рейс
4. Ждёт нажатия Enter — только после этого начинает создавать рейсы
5. Для каждого подходящего заказа: создаёт рейс, добавляет заказ, назначает внешнего исполнителя

Запуск:
    python main.py --file orders.xlsx
    python main.py --file orders.xlsx --dry-run
    python main.py --file orders.xlsx --sheet "Мой лист"
"""

import argparse
import logging
import sys
from dataclasses import dataclass
from typing import Optional

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
# Статусы заказов
# ----------------------------------------------------------------
STATUS_NAMES = {
    "draft":           "В черновиках",
    "onMatch":         "На подборе исполнителя",
    "onExecute":       "В работе ✓",
    "completeExecute": "Исполнение завершено",
    "cancel":          "Отменён",
    "matchHistory":    "История торгов",
    "revoke":          "Отзыв из исполнения",
}


@dataclass
class EnrichedRow:
    """Строка Excel, обогащённая данными из API."""
    row: OrderRow
    system_order_id: Optional[str] = None   # UUID из API
    status: Optional[str] = None            # статус заказа
    existing_flight_ids: list = None        # уже существующие рейсы
    api_error: Optional[str] = None         # ошибка при запросе к API

    def __post_init__(self):
        if self.existing_flight_ids is None:
            self.existing_flight_ids = []

    @property
    def not_found(self) -> bool:
        return self.system_order_id is None and self.api_error is None

    @property
    def wrong_status(self) -> bool:
        return self.status is not None and self.status != "onExecute"

    @property
    def has_flight(self) -> bool:
        return len(self.existing_flight_ids) > 0

    @property
    def ready(self) -> bool:
        """Заказ готов к созданию рейса."""
        return (
            self.api_error is None
            and not self.not_found
            and not self.wrong_status
            and not self.has_flight
        )


# ----------------------------------------------------------------
# Preview и подтверждение
# ----------------------------------------------------------------
SEP = "=" * 70

def print_sep():
    print(SEP)

def print_preview(enriched: list[EnrichedRow], dry_run: bool):
    """Выводит подробный отчёт о состоянии заказов перед обработкой."""
    ready       = [e for e in enriched if e.ready]
    not_found   = [e for e in enriched if e.not_found]
    wrong_st    = [e for e in enriched if e.wrong_status]
    has_flight  = [e for e in enriched if e.has_flight]
    api_errors  = [e for e in enriched if e.api_error]

    print()
    print_sep()
    mode_tag = "  [РЕЖИМ DRY RUN — рейсы создаваться НЕ будут]" if dry_run else ""
    print(f"  ПРЕДВАРИТЕЛЬНАЯ ПРОВЕРКА ЗАКАЗОВ{mode_tag}")
    print_sep()
    print(f"  Всего строк в Excel:           {len(enriched)}")
    print(f"  ✓ Готово к созданию рейса:     {len(ready)}")
    print(f"  ✗ Не найдены в системе:        {len(not_found)}")
    print(f"  ~ Не тот статус (не «В работе»): {len(wrong_st)}")
    print(f"  ~ Уже есть рейс:               {len(has_flight)}")
    print(f"  ! Ошибки API:                  {len(api_errors)}")
    print_sep()

    # --- Готовые к обработке ---
    if ready:
        print(f"\n  ✓ БУДУТ ОБРАБОТАНЫ ({len(ready)} шт.):")
        print(f"  {'№ строки':<10} {'Номер заказа':<30} {'Рейс':<28} {'Компания'}")
        print("  " + "-" * 96)
        for e in ready:
            r = e.row
            print(f"  {str(r.row_number):<10} {r.order_id:<30} {r.flight_name:<28} {r.company_name}")

    # --- Не найдены ---
    if not_found:
        print(f"\n  ✗ НЕ НАЙДЕНЫ В СИСТЕМЕ ({len(not_found)} шт.):")
        print(f"  {'№ строки':<10} {'Номер заказа'}")
        print("  " + "-" * 50)
        for e in not_found:
            print(f"  {str(e.row.row_number):<10} {e.row.order_id}")
        print("  → Проверьте правильность номеров заказов.")
        print("    Используйте номер из интерфейса Магистралей (например, 0-MPPEE-...).")

    # --- Неправильный статус ---
    if wrong_st:
        print(f"\n  ~ НЕПОДХОДЯЩИЙ СТАТУС ({len(wrong_st)} шт.):")
        print(f"  {'№ строки':<10} {'Номер заказа':<30} {'Текущий статус'}")
        print("  " + "-" * 70)
        for e in wrong_st:
            status_str = STATUS_NAMES.get(e.status, e.status)
            print(f"  {str(e.row.row_number):<10} {e.row.order_id:<30} {status_str}")
        print("  → Рейс можно создать только для заказов в статусе «В работе».")

    # --- Уже есть рейс ---
    if has_flight:
        print(f"\n  ~ УЖЕ ЕСТЬ РЕЙС, ПРОПУСКАЕМ ({len(has_flight)} шт.):")
        print(f"  {'№ строки':<10} {'Номер заказа':<30} {'Существующий рейс'}")
        print("  " + "-" * 70)
        for e in has_flight:
            flights_str = ", ".join(e.existing_flight_ids)
            print(f"  {str(e.row.row_number):<10} {e.row.order_id:<30} {flights_str}")

    # --- Ошибки API ---
    if api_errors:
        print(f"\n  ! ОШИБКИ ПРИ ЗАПРОСЕ К API ({len(api_errors)} шт.):")
        for e in api_errors:
            print(f"  Строка {e.row.row_number} / заказ {e.row.order_id}: {e.api_error}")

    print()
    print_sep()


def wait_for_confirmation(ready_count: int, dry_run: bool) -> bool:
    """Показывает запрос подтверждения. Возвращает True если надо продолжать."""
    if ready_count == 0:
        print("  Нет заказов для обработки. Выход.")
        print_sep()
        return False

    if dry_run:
        print(f"  [DRY RUN] Нажмите Enter для симуляции создания {ready_count} рейс(ов),")
        print("            или Ctrl+C для отмены... ")
    else:
        print(f"  Будет создано {ready_count} рейс(ов).")
        print("  Нажмите Enter для запуска, или Ctrl+C для отмены... ")

    print_sep()
    try:
        input()
    except KeyboardInterrupt:
        print("\n  Отменено пользователем.")
        return False
    return True


# ----------------------------------------------------------------
# Основная обработка одной строки
# ----------------------------------------------------------------
def process_row(
    client: MagistraliClient,
    enriched: EnrichedRow,
    forwarder_id: str,
    dry_run: bool,
) -> dict:
    """
    Создаёт рейс для заказа, у которого статус onExecute и нет рейса.
    Возвращает dict: success, flight_id, message.
    """
    row = enriched.row
    order_id = enriched.system_order_id  # UUID для API вызовов
    cargo_id = row.order_id              # человекочитаемый номер для логов

    result = {"success": False, "flight_id": None, "message": ""}

    if dry_run:
        result["success"] = True
        result["flight_id"] = "DRY_RUN"
        result["message"] = (
            f"[DRY RUN] Создан бы рейс '{row.flight_name}' для заказа {cargo_id} "
            f"с исполнителем '{row.company_name}' (ИНН: {row.inn})"
        )
        logger.info("[Строка %d] %s", row.row_number, result["message"])
        return result

    # --- Шаг 1: создаём рейс ---
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

    # --- Шаг 2: добавляем заказ в рейс (через системный UUID) ---
    try:
        client.add_orders_to_flight(flight_id=flight_id, order_ids=[order_id])
        logger.info("[Строка %d] Заказ %s добавлен в рейс %s", row.row_number, cargo_id, flight_id)
    except MagistraliAPIError as e:
        result["message"] = f"Рейс {flight_id} создан, но ошибка добавления заказа: {e}"
        result["flight_id"] = flight_id
        logger.error("[Строка %d] %s", row.row_number, result["message"])
        return result

    # --- Шаг 3: назначаем внешнего исполнителя ("Повезу не я") ---
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
        logger.error("[Строка %d] %s", row.row_number, result["message"])
        return result

    result["success"] = True
    result["flight_id"] = flight_id
    result["message"] = (
        f"Рейс '{row.flight_name}' (ID: {flight_id}) создан для заказа {cargo_id}. "
        f"Внешний исполнитель: {row.company_name} (ИНН: {row.inn})."
    )
    logger.info("[Строка %d] ✓ %s", row.row_number, result["message"])
    return result


# ----------------------------------------------------------------
# Пакетное обогащение данных из API (preview-шаг)
# ----------------------------------------------------------------
def enrich_from_api(
    client: MagistraliClient,
    valid_rows: list[OrderRow],
) -> list[EnrichedRow]:
    """
    Одним пакетным запросом получает данные по всем заказам из API.
    Строит EnrichedRow для каждой строки Excel.
    """
    cargo_ids = [r.order_id for r in valid_rows]

    # --- Пакетный запрос заказов ---
    logger.info("Запрашиваю %d заказов из API (пакетный запрос)...", len(cargo_ids))
    try:
        api_orders = client.get_orders(cargo_ids=cargo_ids)
    except MagistraliAPIError as e:
        logger.error("Критическая ошибка при получении заказов: %s", e)
        # Возвращаем все строки с ошибкой API
        return [EnrichedRow(row=r, api_error=str(e)) for r in valid_rows]

    # Строим индекс: cargoId → данные заказа
    # cargoId в ответе API — это поле "cargoId" (человекочитаемый номер)
    order_by_cargo_id: dict[str, dict] = {}
    for o in api_orders:
        cid = o.get("cargoId") or o.get("cargo_id") or ""
        if cid:
            order_by_cargo_id[cid.strip()] = o

    logger.info("Найдено в API: %d из %d заказов", len(order_by_cargo_id), len(cargo_ids))

    # --- Пакетный запрос существующих рейсов ---
    found_order_uuids = [o.get("id") for o in api_orders if o.get("id")]
    existing_flights_by_order: dict[str, list[str]] = {}

    if found_order_uuids:
        try:
            flights = client.get_flights_for_orders(order_ids=found_order_uuids)
            # Строим индекс: системный UUID заказа → список ID рейсов
            for flight in flights:
                tasks = flight.get("tasks") or flight.get("transferOrders") or []
                flight_id = flight.get("id", "")
                for task in tasks:
                    t_order_id = task.get("transferOrderId") or task.get("orderId") or ""
                    if t_order_id:
                        existing_flights_by_order.setdefault(t_order_id, []).append(flight_id)
        except MagistraliAPIError as e:
            logger.warning("Не удалось проверить существующие рейсы: %s. Продолжаю.", e)

    # --- Собираем EnrichedRow ---
    enriched_list = []
    for row in valid_rows:
        api_order = order_by_cargo_id.get(row.order_id.strip())

        if api_order is None:
            enriched_list.append(EnrichedRow(row=row))
            continue

        system_id = api_order.get("id", "")
        status = api_order.get("status", "")
        flight_ids = existing_flights_by_order.get(system_id, [])

        enriched_list.append(EnrichedRow(
            row=row,
            system_order_id=system_id,
            status=status,
            existing_flight_ids=flight_ids,
        ))

    return enriched_list


# ----------------------------------------------------------------
# Основная функция
# ----------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Создаёт рейсы в Яндекс Магистралях по данным из Excel файла"
    )
    parser.add_argument("--file", "-f", required=True, help="Путь к Excel файлу (.xlsx)")
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
            "Токен не настроен! Укажите YANDEX_OAUTH_TOKEN в config.py."
        )
        sys.exit(1)

    forwarder_id = args.forwarder_id
    if not forwarder_id:
        logger.error(
            "FORWARDER_ID не указан! Заполните его в config.py или передайте через --forwarder-id."
        )
        sys.exit(1)

    # ---- Читаем Excel ----
    try:
        rows = read_excel(args.file, sheet_name=args.sheet)
    except (FileNotFoundError, ValueError) as e:
        logger.error("Ошибка чтения файла: %s", e)
        sys.exit(1)

    if not rows:
        logger.warning("Файл не содержит строк с данными.")
        sys.exit(0)

    # ---- Разделяем невалидные строки Excel ----
    invalid_rows = [r for r in rows if not r.is_valid]
    valid_rows   = [r for r in rows if r.is_valid]

    if invalid_rows:
        print()
        print(f"  ⚠ Строк с ошибками в Excel ({len(invalid_rows)} шт.) — будут пропущены:")
        for row in invalid_rows:
            print(f"    Строка {row.row_number} ('{row.order_id}'): {'; '.join(row.errors)}")

    if not valid_rows:
        logger.error("Нет корректных строк для обработки. Выход.")
        sys.exit(1)

    # ---- Создаём клиент API ----
    client = MagistraliClient(
        token=config.YANDEX_OAUTH_TOKEN,
        base_url=config.BASE_URL,
    )

    # ---- Пакетно обогащаем данные из API ----
    print(f"\n  Проверяю {len(valid_rows)} заказов в Яндекс Магистралях...")
    enriched = enrich_from_api(client, valid_rows)

    # ---- Preview ----
    print_preview(enriched, dry_run=args.dry_run)

    # ---- Подтверждение ----
    ready = [e for e in enriched if e.ready]
    if not wait_for_confirmation(len(ready), dry_run=args.dry_run):
        sys.exit(0)

    # ---- Обрабатываем строки ----
    success_count = 0
    error_count   = 0

    for enriched_row in ready:
        result = process_row(
            client=client,
            enriched=enriched_row,
            forwarder_id=forwarder_id,
            dry_run=args.dry_run,
        )
        if result["success"]:
            success_count += 1
        else:
            error_count += 1

    # ---- Итог ----
    skip_count = len(enriched) - len(ready) - len(invalid_rows)
    print()
    print_sep()
    print("  ИТОГ:")
    print(f"  ✓ Успешно создано рейсов:        {success_count}")
    print(f"  ~ Пропущено (статус/рейс есть):  {skip_count}")
    print(f"  ✗ Ошибок при создании:           {error_count}")
    print(f"  ⚠ Строк с ошибками Excel:        {len(invalid_rows)}")
    print_sep()
    print()

    if error_count > 0:
        sys.exit(2)


if __name__ == "__main__":
    main()
