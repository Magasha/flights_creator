"""
HTTP-клиент для API yamagistrali.ru
Обёртка над requests с автоматическими заголовками, логированием и обработкой ошибок.
"""

import logging
import requests
from typing import Any, Dict, Optional

logger = logging.getLogger(__name__)


class MagistraliAPIError(Exception):
    """Ошибка API Яндекс Магистралей."""

    def __init__(self, status_code: int, message: str, message_code: str = "", raw: dict = None):
        self.status_code = status_code
        self.message = message
        self.message_code = message_code
        self.raw = raw or {}
        super().__init__(f"[HTTP {status_code}] {message_code}: {message}")


class MagistraliClient:
    """
    Клиент для работы с API yamagistrali.ru.

    Использование:
        client = MagistraliClient(token="y0_Ag...", base_url="https://yamagistrali.ru")
        result = client.post("/api/flights/create/v1", {"data": {"executerId": "...", "flightName": "..."}})
    """

    def __init__(self, token: str, base_url: str = "https://yamagistrali.ru"):
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"OAuth {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        })

    def _handle_response(self, response: requests.Response) -> dict:
        """Проверяет ответ и возвращает данные или бросает MagistraliAPIError."""
        try:
            body = response.json()
        except ValueError:
            body = {}

        if not response.ok:
            error = body.get("error", {})
            raise MagistraliAPIError(
                status_code=response.status_code,
                message=error.get("message", response.text or "Неизвестная ошибка"),
                message_code=error.get("message_code", ""),
                raw=body,
            )

        return body

    def post(self, path: str, payload: Dict[str, Any]) -> dict:
        """Выполняет POST запрос."""
        url = f"{self.base_url}{path}"
        logger.debug("POST %s | payload: %s", url, payload)
        response = self.session.post(url, json=payload, timeout=30)
        result = self._handle_response(response)
        logger.debug("POST %s | response: %s", url, result)
        return result

    def get(self, path: str, params: Optional[Dict] = None) -> dict:
        """Выполняет GET запрос."""
        url = f"{self.base_url}{path}"
        logger.debug("GET %s | params: %s", url, params)
        response = self.session.get(url, params=params, timeout=30)
        result = self._handle_response(response)
        logger.debug("GET %s | response: %s", url, result)
        return result

    # ----------------------------------------------------------------
    # Вспомогательные методы для конкретных эндпоинтов
    # ----------------------------------------------------------------

    def get_allowed_forwarders(self) -> list:
        """Возвращает список доступных forwarder ID для текущего пользователя."""
        result = self.get("/api/orders/v0/transferOrder/forwarder/allow")
        return result.get("data", {}).get("forwarders", [])

    def get_orders(
        self,
        order_ids: list[str] = None,
        cargo_ids: list[str] = None,
        statuses: list[str] = None,
        limit: int = 100,
    ) -> list[dict]:
        """
        Получает заказы для заказчика.

        Args:
            order_ids:  Системные UUID заказов (exactOrderIds) — редко используется напрямую
            cargo_ids:  Человекочитаемые номера заказов (exactCargoIds), например '0-MPPEE-2606-11-15-001'
            statuses:   Список статусов (например ['onExecute'])
            limit:      Максимум записей (до 1000)

        Примечание:
            В интерфейсе Магистралей отображается cargoId (например '0-MPPEE-...'),
            а не системный UUID. Поэтому по умолчанию ищем через exactCargoIds.
        """
        payload = {
            "data": {
                "limit": limit,
            }
        }
        if cargo_ids:
            # Основной способ поиска — по отображаемому номеру заказа
            payload["data"]["exactCargoIds"] = cargo_ids
        if order_ids:
            # Резервный способ — по системному UUID
            payload["data"]["exactOrderIds"] = order_ids
        if statuses:
            payload["data"]["statuses"] = statuses

        result = self.post("/api/orders/v0/transferOrder/getFlatForCustomer", payload)
        return result.get("data", {}).get("items", [])

    def get_flights_for_orders(self, order_ids: list[str] = None, cargo_ids: list[str] = None) -> list[dict]:
        """
        Возвращает рейсы, связанные с указанными заказами.

        Args:
            order_ids:  Системные UUID заказов
            cargo_ids:  Человекочитаемые номера (cargoIds)
        """
        payload: dict = {"data": {}}
        if order_ids:
            payload["data"]["orderIds"] = order_ids
        if cargo_ids:
            payload["data"]["cargoIds"] = cargo_ids
        result = self.post("/api/flights/v1/listForForwarder", payload)
        return result.get("data", {}).get("items", [])

    def create_flight(self, executor_id: str, flight_name: str) -> str:
        """
        Создаёт новый пустой рейс.

        Args:
            executor_id: ID исполнителя (ваша организация-экспедитор)
            flight_name: Название рейса

        Returns:
            ID созданного рейса
        """
        payload = {
            "data": {
                "executerId": executor_id,
                "flightName": flight_name,
            }
        }
        result = self.post("/api/flights/create/v1", payload)
        flight_id = result.get("data", {}).get("id")
        if not flight_id:
            raise MagistraliAPIError(0, "Рейс создан, но ID не получен", raw=result)
        return flight_id

    def add_orders_to_flight(self, flight_id: str, order_ids: list[str]) -> dict:
        """
        Добавляет заказы (задачи) в рейс.

        Args:
            flight_id: ID рейса
            order_ids: Список ID заказов для добавления
        """
        payload = {
            "data": {
                "flightId": flight_id,
                "transferOrderIds": order_ids,
            }
        }
        return self.post("/api/flights/tasks/addBulk/v1", payload)

    def assign_external_executor(
        self,
        flight_id: str,
        company_name: str,
        org_type: str,
        country: str,
        tin: str,
        currency: str,
        price: float,
        vat: float,
        price_with_vat: float,
        is_vat_payer: bool = True,
        contact_name: str = None,
        contact_phone: str = None,
        contact_email: str = None,
        additional_comment: str = None,
    ) -> dict:
        """
        Назначает внешнего исполнителя на рейс ("Повезу не я").

        Args:
            flight_id:        ID рейса
            company_name:     Наименование компании
            org_type:         Тип компании (например 'ООО', 'ИП')
            country:          Код страны (например 'RU')
            tin:              ИНН
            currency:         Валюта (например 'RUB')
            price:            Стоимость без НДС
            vat:              Налоговая ставка НДС (например 20.0)
            price_with_vat:   Стоимость с НДС
            is_vat_payer:     Плательщик НДС (True/False)
            contact_name:     ФИО контактного лица (опционально)
            contact_phone:    Телефон контактного лица (опционально)
            contact_email:    Email контактного лица (опционально)
            additional_comment: Комментарий (опционально)
        """
        conditions = {
            "currency": currency,
            "price": price,
            "vat": vat,
            "priceWithVat": price_with_vat,
            "isVatPayer": is_vat_payer,
        }
        if additional_comment:
            conditions["additionalComment"] = additional_comment

        payload = {
            "data": {
                "flightId": flight_id,
                "name": company_name,
                "orgType": org_type,
                "country": country,
                "tin": tin,
                "conditions": conditions,
            }
        }
        if contact_name:
            payload["data"]["contactName"] = contact_name
        if contact_phone:
            payload["data"]["contactPhone"] = contact_phone
        if contact_email:
            payload["data"]["contactEmail"] = contact_email

        return self.post("/api/flights/externalExecutors/assign/v0", payload)

    def suggest_external_executors(self, executor_id: str, query: str) -> list[dict]:
        """Поиск внешних исполнителей по ИНН или названию (для подсказок)."""
        payload = {
            "data": {
                "execId": executor_id,
                "query": query,
            }
        }
        result = self.post("/api/flights/externalExecutors/suggest/v0", payload)
        return result.get("data", {}).get("externalExecutors", [])
