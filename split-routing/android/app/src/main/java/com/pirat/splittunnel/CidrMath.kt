package com.pirat.splittunnel

/**
 * Инвертирует список CIDR-диапазонов.
 * На входе: российские подсети (из antifilter.download).
 * На выходе: всё остальное пространство IPv4 — эти маршруты идут через VPN,
 *            а российский трафик автоматически идёт напрямую.
 */
object CidrMath {

    private data class Net(val ip: Long, val prefix: Int) {

        val mask:  Long = if (prefix == 0) 0L else (0xFFFFFFFFL shl (32 - prefix)) and 0xFFFFFFFFL
        val first: Long = ip and mask
        val last:  Long = first or (mask.inv() and 0xFFFFFFFFL)

        fun contains(o: Net) = prefix <= o.prefix && (o.ip and mask) == first
        fun overlaps(o: Net) = first <= o.last && last >= o.first

        /**
         * Вычитает [ex] из этого диапазона, возвращая оставшиеся куски.
         * Алгоритм: рекурсивно делим пополам, пока не найдём точное совпадение.
         */
        fun subtract(ex: Net): List<Net> {
            if (!overlaps(ex))   return listOf(this)
            if (ex.contains(this)) return emptyList()

            val result = mutableListOf<Net>()
            var cur = Net(first, prefix)

            while (cur.prefix < ex.prefix) {
                val p2   = cur.prefix + 1
                val half = 1L shl (32 - p2)
                val left  = Net(cur.first,        p2)
                val right = Net(cur.first + half, p2)
                if (ex.overlaps(left)) { result += right; cur = left  }
                else                   { result += left;  cur = right }
            }
            // cur == ex — его не добавляем (именно его исключаем)
            return result
        }

        override fun toString(): String {
            val a = (first shr 24) and 0xFF
            val b = (first shr 16) and 0xFF
            val c = (first shr  8) and 0xFF
            val d =  first         and 0xFF
            return "$a.$b.$c.$d/$prefix"
        }
    }

    private fun parse(cidr: String): Net? = runCatching {
        val slash = cidr.indexOf('/')
        if (slash < 0) return null
        val ipStr = cidr.substring(0, slash).trim()
        val prefix = cidr.substring(slash + 1).trim().toInt()
        val parts = ipStr.split(".")
        require(parts.size == 4)
        val ip = parts.fold(0L) { acc, s -> (acc shl 8) or s.trim().toLong() }
        Net(ip, prefix)
    }.getOrNull()

    /** Возвращает AllowedIPs = весь IPv4 минус [excludes]. */
    fun invertCidrs(excludes: List<String>): List<String> {
        val exNets = excludes.mapNotNull { parse(it) }
        var space  = listOf(Net(0L, 0))     // начинаем с 0.0.0.0/0
        for (ex in exNets) {
            space = space.flatMap { it.subtract(ex) }
        }
        return space.map { it.toString() }
    }
}
