package com.pirat.splittunnel

import android.net.VpnService
import android.os.Bundle
import android.view.View
import android.widget.Button
import android.widget.EditText
import android.widget.ScrollView
import android.widget.TextView
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AlertDialog
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import kotlinx.coroutines.launch

class MainActivity : AppCompatActivity() {

    private lateinit var btnToggle: Button
    private lateinit var tvStatus:  TextView
    private lateinit var tvLog:     TextView
    private lateinit var scrollLog: ScrollView

    private val prefs   by lazy { Prefs(this) }
    private val tunnel  by lazy { TunnelManager(applicationContext) }

    // Запрашивает разрешение на VPN и стартует туннель при одобрении
    private val vpnPermLauncher = registerForActivityResult(
        ActivityResultContracts.StartActivityForResult()
    ) { result ->
        if (result.resultCode == RESULT_OK) {
            lifecycleScope.launch { startTunnel() }
        } else {
            appendLog("Разрешение VPN отклонено")
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        btnToggle = findViewById(R.id.btnToggle)
        tvStatus  = findViewById(R.id.tvStatus)
        tvLog     = findViewById(R.id.tvLog)
        scrollLog = findViewById(R.id.scrollLog)

        tunnel.onLog         = { msg -> runOnUiThread { appendLog(msg) } }
        tunnel.onStateChange = { on  -> runOnUiThread { refreshUi(on)  } }

        refreshUi(false)

        btnToggle.setOnClickListener {
            if (tunnel.isActive) {
                lifecycleScope.launch { stopTunnel() }
            } else {
                if (prefs.wgConfig.isBlank()) {
                    showConfigDialog { lifecycleScope.launch { requestVpnAndStart() } }
                } else {
                    lifecycleScope.launch { requestVpnAndStart() }
                }
            }
        }

        findViewById<Button>(R.id.btnConfig).setOnClickListener {
            showConfigDialog(null)
        }
    }

    // ── VPN lifecycle ──────────────────────────────────────────────────────────

    private suspend fun requestVpnAndStart() {
        val intent = VpnService.prepare(this)
        if (intent != null) {
            vpnPermLauncher.launch(intent)   // покажет системный диалог разрешения
        } else {
            startTunnel()                    // разрешение уже есть
        }
    }

    private suspend fun startTunnel() {
        btnToggle.isEnabled = false
        try {
            tunnel.start()
        } catch (e: Exception) {
            appendLog("Ошибка: ${e.message ?: e.javaClass.simpleName}")
            refreshUi(false)
        } finally {
            btnToggle.isEnabled = true
        }
    }

    private suspend fun stopTunnel() {
        btnToggle.isEnabled = false
        try {
            tunnel.stop()
        } catch (e: Exception) {
            appendLog("Ошибка остановки: ${e.message}")
        } finally {
            btnToggle.isEnabled = true
        }
    }

    // ── UI helpers ─────────────────────────────────────────────────────────────

    private fun refreshUi(active: Boolean) {
        if (active) {
            btnToggle.text = "⚓  БРОСИТЬ ЯКОРЬ"
            btnToggle.setBackgroundColor(getColor(R.color.coral))
            tvStatus.text = "★  РУ-трафик идёт напрямую!"
            tvStatus.setTextColor(getColor(R.color.mint))
        } else {
            btnToggle.text = "⚓  ПОДНЯТЬ ПАРУСА"
            btnToggle.setBackgroundColor(getColor(R.color.gold))
            tvStatus.text = "Введи конфиг и нажми кнопку"
            tvStatus.setTextColor(getColor(R.color.fg))
        }
    }

    private fun appendLog(msg: String) {
        tvLog.append("$msg\n")
        scrollLog.post { scrollLog.fullScroll(View.FOCUS_DOWN) }
    }

    private fun showConfigDialog(onSaved: (() -> Unit)?) {
        val input = EditText(this).apply {
            hint = "[Interface]\nPrivateKey = ...\nAddress = 10.x.x.x/24\n\n[Peer]\nPublicKey = ...\nEndpoint = host:51820\nAllowedIPs = 0.0.0.0/0"
            minLines  = 10
            typeface  = android.graphics.Typeface.MONOSPACE
            setTextColor(getColor(R.color.fg))
            setHintTextColor(getColor(R.color.muted))
            setBackgroundColor(getColor(R.color.surface))
            setPadding(24, 16, 24, 16)
            setText(prefs.wgConfig)
        }

        AlertDialog.Builder(this, R.style.Dialog_Dark)
            .setTitle("WireGuard конфиг")
            .setMessage("Экспортируй конфиг из Amnezia или любого WireGuard-сервера и вставь сюда.")
            .setView(input)
            .setPositiveButton("Сохранить") { _, _ ->
                prefs.wgConfig = input.text.toString().trim()
                appendLog("Конфиг сохранён ✓")
                onSaved?.invoke()
            }
            .setNegativeButton("Отмена", null)
            .show()
    }
}
