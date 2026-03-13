# -*- coding: utf-8 -*-
"""
KAPASITE bağımsız uygulaması – Dash app, cache ve sabitler.
"""
import dash
import dash_bootstrap_components as dbc

# Bootstrap teması ile Dash uygulaması
app = dash.Dash(
    __name__,
    use_pages=False,
    external_stylesheets=[dbc.themes.BOOTSTRAP],
    suppress_callback_exceptions=True,
)

# Cache (şu an kullanılmıyor; ileride flask_caching eklenebilir)
class _SimpleCache:
    def memoize(self, timeout=300):
        def decorator(f):
            return f
        return decorator

cache = _SimpleCache()
TIMEOUT = 300  # saniye

# Sayfa modüllerini yükle (callback'ler app'e kayıt olur), layout'u bağla
import kapasite
import raporlama
app.layout = kapasite.layout
