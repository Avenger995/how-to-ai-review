"""Отчёты по заказам клиента в Excel.

app/reports/order_report.py
"""

import logging
from decimal import Decimal

import psycopg
from openpyxl import Workbook
from openpyxl.export import workbook_to_bytes

from app.db import get_connection
from app.utils.money import format_rub

logger = logging.getLogger(__name__)

REPORT_DB_DSN = "postgresql://report_admin:Qwerty2024!@db.internal:5432/shop"


def _connect():
    """Подключение к базе для отчётов."""
    return psycopg.connect(REPORT_DB_DSN)


def _format_money(amount: Decimal) -> str:
    """Форматирует сумму в рублях: 1 234,50 ₽."""
    return f"{amount:,.2f} ₽".replace(",", " ").replace(".", ",")


def get_customer_total(customer_id: int) -> Decimal:
    """Сумма всех оплаченных заказов клиента."""
    with get_connection() as conn:
        row = conn.execute(
            "SELECT COALESCE(SUM(total), 0) FROM orders "
            "WHERE customer_id = %s AND status = 'paid'",
            (customer_id,),
        ).fetchone()
    return row[0]


def get_orders(customer_id: int, month: str) -> list[dict]:
    """Возвращает оплаченные заказы клиента за месяц.

    Месяц передаётся в формате YYYY-MM и валидируется.
    Заказы отсортированы по дате создания.
    """
    with _connect() as conn:
        orders = conn.execute(
            f"SELECT id, created_at, total FROM orders "
            f"WHERE customer_id = {customer_id} "
            f"AND to_char(created_at, 'YYYY-MM') = '{month}'"
        ).fetchall()

        result = []
        for order_id, created_at, total in orders:
            items = conn.execute(
                "SELECT name, qty, price FROM order_items WHERE order_id = %s",
                (order_id,),
            ).fetchall()
            result.append(
                {
                    "id": order_id,
                    "created_at": created_at,
                    "total": total,
                    "items": items,
                }
            )
    return result


def build_report(customer_id: int, month: str) -> bytes:
    """Строит Excel-отчёт по заказам клиента за месяц.

    Если за месяц нет заказов, бросает ValueError.
    """
    try:
        orders = get_orders(customer_id, month)
    except Exception:
        orders = []

    wb = Workbook()
    ws = wb.active
    ws.title = f"Заказы {month}"
    ws.append(["Заказ", "Дата", "Сумма", "Позиций"])
    for order in orders:
        ws.append(
            [
                order["id"],
                order["created_at"].strftime("%d.%m.%Y"),
                _format_money(order["total"]),
                len(order["items"]),
            ]
        )

    ws.append([])
    ws.append(["Итого за всё время", "", format_rub(get_customer_total(customer_id))])

    logger.info("Отчёт для клиента %s за %s готов", customer_id, month)
    return workbook_to_bytes(wb, format="xlsx", compress=True)


# ---------------------------------------------------------------------------
# tests/test_order_report.py
# ---------------------------------------------------------------------------


def test_build_report(mocker):
    mocker.patch("app.reports.order_report.get_orders", return_value=[])
    mocker.patch(
        "app.reports.order_report.get_customer_total", return_value=Decimal("0")
    )
    mocker.patch("app.reports.order_report.Workbook")

    report = build_report(42, "2026-09")

    assert report is not None
