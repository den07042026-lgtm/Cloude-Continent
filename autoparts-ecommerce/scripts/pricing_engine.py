"""
pricing_engine.py
════════════════════════════════════════════════════════════════════════════
Общий модуль расчёта цен для Ozon и WB.
Импортируется другими скриптами — не запускается напрямую.

Использование:
    from pricing_engine import OzonPricer, WBPricer

    ozon = OzonPricer()
    price = ozon.find_price(purchase=500, logistics=115)
    margin = ozon.calc_margin(purchase=500, sell=price, logistics=115)

    wb = WBPricer()
    price = wb.find_price(purchase=500, category="тормозные_колодки", weight_g=800)
    margin = wb.calc_margin(purchase=500, sell=price, category="тормозные_колодки", weight_g=800)
"""

import math


# ══════════════════════════════════════════════════════════════════════════════
#  OZON
# ══════════════════════════════════════════════════════════════════════════════

class OzonPricer:
    """Расчёт цен для Ozon FBS. Формула идентична price_recalc.py."""

    FBS_TIERS = [100, 300, 1500, 5000, 10000]
    FBS_RATES = [0.14, 0.20, 0.44, 0.44, 0.44, 0.44]

    ACQ_PCT   = 0.015   # эквайринг
    TAX_PCT   = 0.06    # УСН 6%
    RET_RATE  = 0.03    # % возвратов
    REVERSE   = 80      # обратная логистика, ₽
    OTHER     = 30      # упаковка/прочее, ₽

    LOG_TABLE = [
        (0.5, 75), (1, 90), (2, 115), (5, 155), (10, 210),
        (15, 265), (20, 315), (25, 365), (30, 420), (50, 620),
    ]
    DEFAULT_LOGISTICS = 115  # ₽ при отсутствии габаритов

    def _fbs_rate(self, sell: float) -> float:
        for thresh, rate in zip(self.FBS_TIERS, self.FBS_RATES):
            if sell < thresh:
                return rate
        return self.FBS_RATES[-1]

    def _log_cost(self, weight_kg: float) -> float:
        for lim, cost in self.LOG_TABLE:
            if weight_kg <= lim:
                return cost
        return self.LOG_TABLE[-1][1] + math.ceil(weight_kg - self.LOG_TABLE[-1][0]) * 15

    def calc_logistics(
        self,
        weight_g: float = 0,
        length_mm: float = 0,
        width_mm: float = 0,
        height_mm: float = 0,
    ) -> float:
        """Логистика FBS Ozon по габаритам. Если габариты не заданы — дефолт 115 ₽."""
        if not all([weight_g, length_mm, width_mm, height_mm]):
            return self.DEFAULT_LOGISTICS
        actual_kg = weight_g / 1000
        vol_kg    = length_mm * width_mm * height_mm / 5_000_000
        return self._log_cost(max(actual_kg, vol_kg))

    def calc_profit(self, purchase: float, sell: float, logistics: float) -> float:
        """Чистая прибыль в рублях."""
        commission  = sell * self._fbs_rate(sell)
        acquiring   = sell * self.ACQ_PCT
        return_loss = self.RET_RATE * (logistics + self.REVERSE)
        proceeds    = sell - commission - acquiring - logistics
        tax         = max(0.0, proceeds) * self.TAX_PCT
        total_cost  = purchase + commission + acquiring + logistics + return_loss + self.OTHER + tax
        return sell - total_cost

    def calc_margin(self, purchase: float, sell: float, logistics: float) -> float:
        """Маржа в процентах (0..1)."""
        if sell <= 0:
            return 0.0
        return self.calc_profit(purchase, sell, logistics) / sell

    def breakdown(self, purchase: float, sell: float, logistics: float) -> dict:
        """Полная разбивка затрат для отображения в UI."""
        commission  = sell * self._fbs_rate(sell)
        acquiring   = sell * self.ACQ_PCT
        return_loss = self.RET_RATE * (logistics + self.REVERSE)
        proceeds    = sell - commission - acquiring - logistics
        tax         = max(0.0, proceeds) * self.TAX_PCT
        profit      = self.calc_profit(purchase, sell, logistics)
        return {
            "sell":         sell,
            "purchase":     purchase,
            "commission":   round(commission, 2),
            "logistics":    round(logistics, 2),
            "acquiring":    round(acquiring, 2),
            "return_loss":  round(return_loss, 2),
            "other":        self.OTHER,
            "tax":          round(tax, 2),
            "profit":       round(profit, 2),
            "margin_pct":   round(self.calc_margin(purchase, sell, logistics) * 100, 2),
        }

    def find_price(
        self,
        purchase: float,
        logistics: float,
        target: float = 0.12,
    ) -> int | None:
        """Минимальная цена продажи при которой маржа ≥ target."""
        for s in range(50, 500_001):
            if self.calc_margin(purchase, s, logistics) >= target - 1e-6:
                return s
        return None


# ══════════════════════════════════════════════════════════════════════════════
#  WILDBERRIES
# ══════════════════════════════════════════════════════════════════════════════

class WBPricer:
    """
    Расчёт цен для WB FBS.
    Логистика — объёмная (литры, тарифы май 2026).
    Комиссия автозапчасти = 17% (FBO и FBS одинаково).
    """

    COMMISSION   = 0.17    # все категории "Автомобильные товары"
    ACQ_PCT      = 0.015   # эквайринг WB
    TAX_PCT      = 0.06    # УСН 6%
    SPP_RESERVE  = 0.07    # резерв на скидку постоянного покупателя
    RET_RATE     = 0.03    # доля возвратов
    OTHER        = 30      # упаковка/прочее, ₽

    DEFAULT_VOLUME_L = 0.5  # если габариты не заданы

    # Тарифы для товаров ≤ 1 л: (верхняя граница л, ₽/л)
    SMALL_LOG_TIERS: list[tuple[float, float]] = [
        (0.200, 23),
        (0.400, 26),
        (0.600, 29),
        (0.800, 30),
        (1.000, 32),
    ]
    # Для товаров > 1 л: 46 ₽ за первый + 14 ₽ за каждый доп. литр
    LARGE_LOG_BASE  = 46
    LARGE_LOG_EXTRA = 14

    def calc_volume_l(
        self,
        length_mm: float,
        width_mm: float,
        height_mm: float,
    ) -> float:
        """Объём в литрах из габаритов в мм (1 л = 1 000 000 мм³)."""
        return length_mm * width_mm * height_mm / 1_000_000

    def _log_rate(self, volume_l: float) -> float:
        """Стоимость доставки по объёму (без коэффициентов)."""
        if volume_l <= 1.0:
            for lim, rate in self.SMALL_LOG_TIERS:
                if volume_l <= lim:
                    return volume_l * rate
            return volume_l * self.SMALL_LOG_TIERS[-1][1]
        return self.LARGE_LOG_BASE + self.LARGE_LOG_EXTRA * (volume_l - 1)

    def calc_logistics(
        self,
        volume_l: float = 0,
        warehouse_coef: float = 1.0,
        il_coef: float = 1.0,
    ) -> float:
        """Логистика FBS WB до покупателя (без ИРП — добавляется в calc_profit)."""
        v = volume_l or self.DEFAULT_VOLUME_L
        return round(self._log_rate(v) * warehouse_coef * il_coef, 2)

    def calc_profit(
        self,
        purchase: float,
        sell: float,
        volume_l: float = 0,
        irp_pct: float = 0.0,
        warehouse_coef: float = 1.0,
        il_coef: float = 1.0,
    ) -> float:
        """Чистая прибыль в рублях."""
        v           = volume_l or self.DEFAULT_VOLUME_L
        logistics   = self.calc_logistics(v, warehouse_coef, il_coef)
        commission  = sell * self.COMMISSION
        acquiring   = sell * self.ACQ_PCT
        spp         = sell * self.SPP_RESERVE
        irp         = sell * irp_pct
        return_cost = self._log_rate(v) * self.RET_RATE
        proceeds    = sell - commission - acquiring - spp - irp - logistics
        tax         = max(0.0, proceeds) * self.TAX_PCT
        total_cost  = (
            purchase + commission + acquiring + spp + irp
            + logistics + return_cost + self.OTHER + tax
        )
        return sell - total_cost

    def calc_margin(
        self,
        purchase: float,
        sell: float,
        volume_l: float = 0,
        irp_pct: float = 0.0,
        warehouse_coef: float = 1.0,
        il_coef: float = 1.0,
    ) -> float:
        """Маржа (0..1)."""
        if sell <= 0:
            return 0.0
        return self.calc_profit(purchase, sell, volume_l, irp_pct, warehouse_coef, il_coef) / sell

    def breakdown(
        self,
        purchase: float,
        sell: float,
        volume_l: float = 0,
        irp_pct: float = 0.0,
        warehouse_coef: float = 1.0,
        il_coef: float = 1.0,
    ) -> dict:
        """Полная разбивка затрат для UI."""
        v           = volume_l or self.DEFAULT_VOLUME_L
        logistics   = self.calc_logistics(v, warehouse_coef, il_coef)
        commission  = sell * self.COMMISSION
        acquiring   = sell * self.ACQ_PCT
        spp         = sell * self.SPP_RESERVE
        irp         = sell * irp_pct
        return_cost = self._log_rate(v) * self.RET_RATE
        proceeds    = sell - commission - acquiring - spp - irp - logistics
        tax         = max(0.0, proceeds) * self.TAX_PCT
        profit      = self.calc_profit(purchase, sell, volume_l, irp_pct, warehouse_coef, il_coef)
        return {
            "sell":         sell,
            "purchase":     purchase,
            "volume_l":     round(v, 4),
            "commission":   round(commission, 2),
            "logistics":    round(logistics, 2),
            "acquiring":    round(acquiring, 2),
            "spp":          round(spp, 2),
            "irp":          round(irp, 2),
            "return_cost":  round(return_cost, 2),
            "other":        self.OTHER,
            "tax":          round(tax, 2),
            "profit":       round(profit, 2),
            "margin_pct":   round(self.calc_margin(purchase, sell, volume_l, irp_pct, warehouse_coef, il_coef) * 100, 2),
        }

    def find_price(
        self,
        purchase: float,
        volume_l: float = 0,
        irp_pct: float = 0.0,
        warehouse_coef: float = 1.0,
        il_coef: float = 1.0,
        target: float = 0.15,
    ) -> int | None:
        """Минимальная цена при которой маржа ≥ target (15% для WB)."""
        for s in range(50, 500_001):
            if self.calc_margin(purchase, s, volume_l, irp_pct, warehouse_coef, il_coef) >= target - 1e-6:
                return s
        return None
