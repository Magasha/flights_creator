"""
HTTP-клиент для API yamagistrali.ru
Обёртка над requests с автоматическими заголовками, логированием и обработкой ошибок.

Роль: ИСПОЛНИТЕЛЬ (executor).

Правильный flow:
  1. getForFlightAddForExecutor(executorId) → все доступные задачи с taskId + taskName
  2. Фильтрация по cargoId из Excel (сопоставление по taskName)
  3. create_flight → flightId
  4. addBulk(flightId, taskId) → задачи добавлены в рейс
  5. externalExecutors/assign → "Повезу не я"
"""

import logging
import requests
from typing import Any, Dict, List, Optional

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
    Клиент для работы с API yamagistrali.ru (роль: исполнитель).

    Использование:
        client = MagistraliClient(token="y0_Ag...", base_url="https://yamagistrali.ru")
    """

    def __init__(self, token: str, base_url: str = "https://partner-test-yamagistrali.ru/"):
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"OAuth {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        })

    def _handle_response(self, response: requests.Response) -> dict:
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
        url = f"{self.base_url}{path}"
        logger.debug("POST %s | payload: %s", url, payload)
        response = self.session.post(url, json=payload, timeout=30)
        result = self._handle_response(response)
        logger.debug("POST %s | response: %s", url, result)
        return result

    def get(self, path: str, params: Optional[Dict] = None) -> dict:
        url = f"{self.base_url}{path}"
        logger.debug("GET %s | params: %s", url, params)
        response = self.session.get(url, params=params, timeout=30)
        result = self._handle_response(response)
        logger.debug("GET %s | response: %s", url, result)
        return result

    # ----------------------------------------------------------------
    # Получение доступных задач для добавления в рейс
    # ----------------------------------------------------------------

    def get_available_tasks(self, executor_id: str) -> List[dict]:
        """
        Возвращает все задачи (заказы), доступные исполнителю для добавления в рейс.

        Эндпоинт: POST /api/orders/v0/transferOrder/getForFlightAddForExecutor
        Принимает: executorId
        Возвращает список задач с полями:
            taskId      — системный ID задачи (используется в addBulk)
            taskName    — человекочитаемый номер задачи (= cargoId из Excel)
            orderId     — ID заказа
            orderName   — название заказа
            shipmentId  — ID отправки
            executorId  — ID исполнителя
            forwarderId — ID экспедитора
        """
        payload = {"data": {"executorId": executor_id}}
        result = self.post("/api/orders/v0/transferOrder/getForFlightAddForExecutor", payload)
        return result.get("data", {}).get("tasks", [])

    def find_tasks_by_cargo_ids(self, executor_id: str, cargo_ids: List[str]) -> Dict[str, dict]:
        """
        Находит задачи по списку cargoId из Excel.

        Возвращает словарь: {cargoId: task_dict}
        Ненайденные cargoId будут отсутствовать в словаре.

        Сопоставление: taskName == cargoId (регистронезависимо, с trim).
        """
        all_tasks = self.get_available_tasks(executor_id)

        # Строим индекс по orderId и taskName (оба варианта, на случай разных форматов)
        task_index: Dict[str, dict] = {}
        for task in all_tasks:
            for field in ("orderId", "taskName"):
                name = (task.get(field) or "").strip().upper()
                if name:
                    task_index[name] = task

        result: Dict[str, dict] = {}
        for cid in cargo_ids:
            normalized = cid.strip().upper()
            if normalized in task_index:
                result[cid] = task_index[normalized]

        return result

    # ----------------------------------------------------------------
    # Рейсы
    # ----------------------------------------------------------------

    def create_flight(self, executor_id: str, flight_name: str) -> str:
        """
        Создаёт новый пустой рейс.

        Args:
            executor_id: ID организации-исполнителя
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

    def add_tasks_to_flight(self, flight_id: str, task_ids: List[str]) -> dict:
        """
        Добавляет задачи в рейс по taskId.

        Args:
            flight_id: ID рейса
            task_ids:  Список taskId (из getForFlightAddForExecutor)
        """
        tasks = [
            {
                "flightId": flight_id,
                "taskId": tid,
                "loadingOrder": i + 1,
                "unloadingOrder": i + 1,
                "externalId": f"ext-{i+1}",
            }
            for i, tid in enumerate(task_ids)
        ]
        payload = {"data": {"tasks": tasks}}
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

    def get_allowed_forwarders(self) -> list:
        """Список доступных forwarder ID для текущего токена."""
        result = self.get("/api/orders/v0/transferOrder/forwarder/allow")
        return result.get("data", {}).get("forwarders", [])
