package com.pirat.splittunnel

import android.content.Context
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.File
import java.net.URL

object IpRepository {

    private const val PRIMARY  = "https://antifilter.download/list/subnet.lst"
    private const val FALLBACK = "https://www.ipdeny.com/ipblocks/data/countries/ru.zone"
    private const val CACHE    = "ru_cidrs.txt"
    private const val TTL_MS   = 7L * 24 * 60 * 60 * 1000   // 7 дней

    suspend fun fetch(context: Context): List<String> = withContext(Dispatchers.IO) {
        val file = File(context.cacheDir, CACHE)

        // Используем кеш, если он свежий
        if (file.exists() && System.currentTimeMillis() - file.lastModified() < TTL_MS) {
            return@withContext parse(file.readText())
        }

        // Скачиваем
        for (url in listOf(PRIMARY, FALLBACK)) {
            try {
                val text = URL(url).readText()
                file.writeText(text)
                return@withContext parse(text)
            } catch (_: Exception) { /* пробуем следующий источник */ }
        }

        // Если сеть недоступна — используем устаревший кеш
        if (file.exists()) return@withContext parse(file.readText())

        error("Не удалось загрузить список IP — нет сети и нет кеша")
    }

    private fun parse(text: String) = text.lines()
        .map  { it.trim() }
        .filter { it.isNotEmpty() && !it.startsWith("#") && '/' in it }
}
