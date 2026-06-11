"""
Диагностика API yamagistrali.ru.

Проверяет: getForFlightAddForExecutor — правильный эндпоинт для получения задач исполнителя.

Запуск: python debug_api.py --cargo-id 0-MPPEE-2606-11-15-001
"""

import argparse
import json
import config
from api_client import MagistraliClient, MagistraliAPIError


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--cargo-id", required=True, help="Номер заказа из Excel, например 0-MPPEE-2606-11-15-001")
    args = parser.parse_args()
    cid = args.cargo_id

    client = MagistraliClient(token=config.YANDEX_OAUTH_TOKEN, base_url=config.BASE_URL)

    print(f"\n{'='*60}")
    print(f"  Диагностика API yamagistrali.ru")
    print(f"  EXECUTOR_ID: {config.EXECUTOR_ID}")
    print(f"  Ищем заказ:  {cid}")
    print(f"{'='*60}\n")

    # --- Тест 1: getForFlightAddForExecutor ---
    print("  [1] Запрашиваю задачи через getForFlightAddForExecutor...")
    try:
        r = client.post(
            "/api/orders/v0/transferOrder/getForFlightAddForExecutor",
            {"data": {"executorId": config.EXECUTOR_ID}},
        )
        tasks = r.get("data", {}).get("tasks", [])
        print(f"      Всего задач доступно: {len(tasks)}")

        if tasks:
            # Показываем первые 3 задачи
            print("\n      Первые 3 задачи из API:")
            for t in tasks[:3]:
                print(f"        taskId:   {t.get('taskId')}")
                print(f"        taskName: {t.get('taskName')}")
                print(f"        orderId:  {t.get('orderId')}")
                print(f"        orderName:{t.get('orderName')}")
                print()

            # Ищем нужный cargoId по orderId или taskName
            found = None
            for t in tasks:
                for field in ("orderId", "taskName"):
                    if (t.get(field) or "").strip().upper() == cid.strip().upper():
                        found = t
                        break
                if found:
                    break

            print(f"{'='*60}")
            if found:
                print(f"  НАЙДЕНО: {cid}")
                print(f"  taskId:   {found['taskId']}")
                print(f"  taskName: {found.get('taskName')}")
                print(f"  orderId:  {found.get('orderId')}")
                print(f"\n  Всё готово — основной скрипт должен работать.")
            else:
                print(f"  НЕ НАЙДЕНО: {cid} среди {len(tasks)} задач")
                print()
                print(f"  Возможные причины:")
                print(f"    1. EXECUTOR_ID в config.py неверный (текущий: {config.EXECUTOR_ID})")
                print(f"    2. Номер заказа введён неправильно")
                print(f"    3. Заказ не в статусе 'В работе' или уже в рейсе")
                print()
                print(f"  Все доступные orderId (первые 20):")
                for t in tasks[:20]:
                    print(f"    orderId={t.get('orderId')}  taskName={t.get('taskName')}  taskId={t.get('taskId')}")
            print(f"{'='*60}")

        else:
            print(f"\n  Задач не найдено (пустой список).")
            print(f"  Проверьте EXECUTOR_ID в config.py: {config.EXECUTOR_ID}")
            print(f"\n  Полный ответ API:")
            print(json.dumps(r, ensure_ascii=False, indent=2)[:1000])

    except MagistraliAPIError as e:
        print(f"  ОШИБКА: {e}")
        if e.raw:
            print(f"  Raw: {json.dumps(e.raw, ensure_ascii=False)[:500]}")

    # --- Тест 2: проверяем EXECUTOR_ID через allowed forwarders ---
    print(f"\n{'='*60}")
    print("  [2] Проверяю доступные forwarder/executor ID для токена...")
    try:
        r = client.get("/api/orders/v0/transferOrder/forwarder/allow")
        items = r.get("data", {}).get("forwarders", [])
        print(f"      Доступных ID: {len(items)}")
        for item in items:
            print(f"        {item}")
    except MagistraliAPIError as e:
        print(f"  ОШИБКА: {e}")

    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
