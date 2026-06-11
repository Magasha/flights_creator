"""
Узнаём допустимые значения orgType методом перебора.

Создаёт один тестовый рейс, пробует assign с разными orgType,
затем удаляет рейс через clear.

Запуск: python debug_orgtype.py --tin 7802355025 --name "ФТК СОТРАНС"
"""

import argparse
import json
import config
from api_client import MagistraliClient, MagistraliAPIError

# Все варианты которые могут быть приняты API
ORG_TYPE_CANDIDATES = [
    "OOO", "ООО", "LLC",
    "ИП", "IP", "IE",
    "AO", "АО", "JSC",
    "OAO", "ОАО", "PJSC",
    "ZAO", "ЗАО", "CJSC",
    "PAO", "ПАО",
    "NKO", "НКО",
    "INDIVIDUAL", "COMPANY",
    "1", "2", "3",
]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--tin",  required=True, help="ИНН, например 7802355025")
    parser.add_argument("--name", required=True, help="Название компании")
    args = parser.parse_args()

    client = MagistraliClient(token=config.YANDEX_OAUTH_TOKEN, base_url=config.BASE_URL)

    print(f"\n{'='*60}")
    print(f"  Подбор правильного значения orgType")
    print(f"  ИНН: {args.tin}  Компания: {args.name}")
    print(f"{'='*60}\n")

    # Создаём тестовый рейс
    print("  Создаю тестовый рейс...")
    try:
        flight_id = client.create_flight(
            executor_id=config.EXECUTOR_ID,
            flight_name="[TEST] debug_orgtype — удалить",
        )
        print(f"  Рейс создан: {flight_id}\n")
    except MagistraliAPIError as e:
        print(f"  ОШИБКА создания рейса: {e}")
        return

    # Пробуем каждый вариант
    working = []
    for org_type in ORG_TYPE_CANDIDATES:
        try:
            client.assign_external_executor(
                flight_id=flight_id,
                company_name=args.name,
                org_type=org_type,
                country="RU",
                tin=args.tin,
                currency="RUB",
                price=1000.0,
                vat=20.0,
                price_with_vat=1200.0,
                is_vat_payer=True,
            )
            print(f"  ✓ РАБОТАЕТ: orgType = '{org_type}'")
            working.append(org_type)
        except MagistraliAPIError as e:
            msg = str(e)[:80]
            print(f"  ✗ '{org_type}': {msg}")

    # Удаляем тестовый рейс
    print(f"\n  Удаляю тестовый рейс {flight_id}...")
    try:
        client.post("/api/flights/externalExecutors/clear/v0",
                    {"data": {"flightId": flight_id}})
        print("  Рейс очищен.")
    except MagistraliAPIError as e:
        print(f"  Не удалось очистить рейс: {e}")

    print(f"\n{'='*60}")
    if working:
        print(f"  Рабочие значения orgType: {working}")
        print(f"  Используй одно из них в Excel.")
    else:
        print("  Ни один вариант не сработал — возможно проблема в другом поле.")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
