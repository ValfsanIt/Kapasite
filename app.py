# -*- coding: utf-8 -*-
"""
KAPASITE bağımsız uygulaması – Dash app, cache ve sabitler.
"""
import os
import dash
import dash_bootstrap_components as dbc

# Proje kökü (PyCharm/Cursor fark etmez, her zaman app.py'nin bulunduğu klasör)
_ROOT = os.path.dirname(os.path.abspath(__file__))
_ASSETS = os.path.join(_ROOT, "assets")

# assets yoksa oluştur ve CSS kopyala (örn. GitHub clone sonrası)
if not os.path.isdir(_ASSETS):
    os.makedirs(_ASSETS, exist_ok=True)
for name in ("custom.css", "mes_styles.css"):
    src = os.path.join(_ROOT, name)
    dst = os.path.join(_ASSETS, name)
    if os.path.isfile(src) and (not os.path.isfile(dst) or os.path.getmtime(src) > os.path.getmtime(dst)):
        try:
            import shutil
            shutil.copy2(src, dst)
        except Exception:
            pass

# Bootstrap teması + assets klasörü mutlak yol (PyCharm'da da doğru yüklensin)
app = dash.Dash(
    __name__,
    use_pages=False,
    external_stylesheets=[dbc.themes.BOOTSTRAP],
    suppress_callback_exceptions=True,
    assets_folder=_ASSETS,
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
