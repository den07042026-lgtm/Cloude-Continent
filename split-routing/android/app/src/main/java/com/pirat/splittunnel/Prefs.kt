package com.pirat.splittunnel

import android.content.Context
import androidx.core.content.edit

class Prefs(context: Context) {

    private val sp = context.getSharedPreferences("pirat", Context.MODE_PRIVATE)

    var wgConfig: String
        get()      = sp.getString("wg_config", "") ?: ""
        set(value) = sp.edit { putString("wg_config", value) }
}
