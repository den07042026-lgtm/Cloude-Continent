package com.pirat.splittunnel

import android.content.Context
import com.wireguard.android.backend.GoBackend
import com.wireguard.android.backend.Tunnel
import com.wireguard.config.Config
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext

/**
 * Управляет WireGuard-туннелем со встроенным split-tunneling.
 *
 * Логика:
 *  - Скачивает список российских IP (~3000 подсетей)
 *  - Инвертирует его: AllowedIPs = всё IPv4 КРОМЕ российских
 *  - Поднимает WireGuard-туннель с этим AllowedIPs
 *  → Российский трафик идёт напрямую (минует VPN)
 *  → Всё остальное (Telegram, зарубежные сайты) идёт через VPN-сервер
 */
class TunnelManager(private val context: Context) : Tunnel {

    private val backend = GoBackend(context)
    private val prefs   = Prefs(context)

    @Volatile var isActive = false
        private set

    var onLog:         ((String)  -> Unit)? = null
    var onStateChange: ((Boolean) -> Unit)? = null

    // ── Tunnel interface ──────────────────────────────────────────────────────

    override fun getName(): String = "pirat"

    override fun onStateChange(newState: Tunnel.State) {
        isActive = newState == Tunnel.State.UP
        onStateChange?.invoke(isActive)
    }

    // ── Public API ────────────────────────────────────────────────────────────

    suspend fun start() = withContext(Dispatchers.IO) {
        log("Скачиваю список российских IP...")
        val ruCidrs = IpRepository.fetch(context)
        log("Получено ${ruCidrs.size} сетей  ⚓")

        log("Рассчитываю split-маршруты...")
        val allowedIps = CidrMath.invertCidrs(ruCidrs)
        log("Маршрутов для VPN: ${allowedIps.size}")

        val config = buildConfig(allowedIps)
        log("Поднимаю туннель...")
        backend.setState(this@TunnelManager, Tunnel.State.UP, config)
        log("★  Курс проложен! РУ-трафик идёт напрямую~")
    }

    suspend fun stop() = withContext(Dispatchers.IO) {
        log("Останавливаю туннель...")
        backend.setState(this@TunnelManager, Tunnel.State.DOWN, null)
        log("Якорь брошен. Туннель закрыт.")
    }

    // ── Private ───────────────────────────────────────────────────────────────

    /**
     * Берёт WireGuard-конфиг пользователя и заменяет AllowedIPs
     * на вычисленный split-tunnel список.
     */
    private fun buildConfig(allowedIps: List<String>): Config {
        val sb       = StringBuilder()
        var inPeer   = false
        var replaced = false

        for (raw in prefs.wgConfig.lines()) {
            val line = raw.trim()
            when {
                line.equals("[Peer]", ignoreCase = true) -> {
                    // Если предыдущий [Peer] не имел AllowedIPs — добавляем перед новым
                    if (inPeer && !replaced) {
                        sb.appendLine("AllowedIPs = ${allowedIps.joinToString(", ")}")
                    }
                    inPeer = true; replaced = false
                    sb.appendLine(raw)
                }
                inPeer && !replaced && line.startsWith("AllowedIPs", ignoreCase = true) -> {
                    sb.appendLine("AllowedIPs = ${allowedIps.joinToString(", ")}")
                    replaced = true
                }
                else -> sb.appendLine(raw)
            }
        }
        // Хвост: если последний [Peer] без AllowedIPs
        if (inPeer && !replaced) {
            sb.appendLine("AllowedIPs = ${allowedIps.joinToString(", ")}")
        }

        return Config.parse(sb.toString().trimEnd().reader().buffered())
    }

    private fun log(msg: String) = onLog?.invoke(msg)
}
