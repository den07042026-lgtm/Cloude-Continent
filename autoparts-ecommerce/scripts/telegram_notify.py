"""
telegram_notify.py
══════════════════
Общий модуль Telegram-уведомлений. Импортируется другими скриптами.

    from telegram_notify import tg_order, tg_alert, tg_stock_done, tg_price_done
"""

import requests
from datetime import datetime

_TG_API = "https://api.telegram.org/bot{}/sendMessage"


def tg_send(token: str, chat_id: str, text: str, parse_mode: str = "HTML") -> bool:
    """Базовая отправка. Возвращает True при успехе."""
    if not token or not chat_id:
        return False
    try:
        r = requests.post(
            _TG_API.format(token),
            json={"chat_id": chat_id, "text": text, "parse_mode": parse_mode},
            timeout=10,
        )
        return r.ok
    except Exception:
        return False


def _now() -> str:
    return datetime.now().strftime("%d.%m.%Y %H:%M")


def tg_order(token: str, chat_id: str, order_id: str, items: list[dict]) -> bool:
    """
    Уведомление о новом заказе.
    items: [{mikado_code, name, quantity, mikado_qty}]
    """
    lines = [f"🛒 <b>Новый заказ</b>  #{order_id}", ""]
    for it in items:
        code  = it.get("mikado_code", "—")
        name  = (it.get("name") or "")[:45]
        qty   = it.get("quantity", 1)
        avail = it.get("mikado_qty", "?")
        if isinstance(avail, int) and avail >= qty:
            sign = "✅"
        elif isinstance(avail, int) and avail > 0:
            sign = "⚠️"
        else:
            sign = "❌"
        lines.append(f"{sign} <code>{code}</code>  ×{qty}  (склад: {avail} шт.)")
        if name:
            lines.append(f"   <i>{name}</i>")
    lines += ["", f"🕐 {_now()}"]
    return tg_send(token, chat_id, "\n".join(lines))


def tg_mikado_error(token: str, chat_id: str, order_id: str, failed_items: list[dict]) -> bool:
    """Алерт: не удалось заказать позиции у Микадо."""
    lines = [f"🚨 <b>Ошибка заказа Микадо</b>  #{order_id}", ""]
    lines.append("Следующие позиции <b>не заказаны</b> — разберитесь вручную:")
    for it in failed_items:
        code  = it.get("mikado_code", "—")
        name  = (it.get("name") or "")[:45]
        qty   = it.get("quantity", 1)
        avail = it.get("mikado_qty", "?")
        lines.append(f"  ❌ <code>{code}</code>  ×{qty}  (склад: {avail} шт.)")
        if name:
            lines.append(f"     <i>{name}</i>")
    lines += ["", f"🕐 {_now()}"]
    return tg_send(token, chat_id, "\n".join(lines))


def tg_stock_done(token: str, chat_id: str, total: int, in_stock: int, ms_updated: int) -> bool:
    """Итог синхронизации остатков."""
    text = (
        f"📦 <b>Остатки синхронизированы</b>\n"
        f"Позиций в прайсе: {total}\n"
        f"В наличии: {in_stock}\n"
        f"Обновлено в МойСклад: {ms_updated}\n"
        f"🕐 {_now()}"
    )
    return tg_send(token, chat_id, text)


def tg_price_done(token: str, chat_id: str, updated: int, skipped: int) -> bool:
    """Итог ночного пересчёта цен."""
    text = (
        f"💰 <b>Цены пересчитаны</b>\n"
        f"Обновлено в МойСклад: {updated} товаров\n"
        f"Пропущено (нет данных): {skipped}\n"
        f"🕐 {_now()}"
    )
    return tg_send(token, chat_id, text)


def tg_alert(token: str, chat_id: str, title: str, body: str) -> bool:
    """Произвольный алерт об ошибке."""
    text = f"🚨 <b>{title}</b>\n\n{body}\n\n🕐 {_now()}"
    return tg_send(token, chat_id, text)
