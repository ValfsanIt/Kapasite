# -*- coding: utf-8 -*-

import os
import dash
import dash_bootstrap_components as dbc
from dash import html, dcc, Input, Output


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


app = dash.Dash(
    __name__,
    use_pages=False,
    external_stylesheets=[dbc.themes.BOOTSTRAP],
    suppress_callback_exceptions=True,
    assets_folder=_ASSETS,
)


class _SimpleCache:
    def memoize(self, timeout=300):
        def decorator(f):
            return f
        return decorator

cache = _SimpleCache()
TIMEOUT = 300  # saniye

# Sayfa modüllerini yükle (callback'ler app'e kayıt olur), layout'u bağla
import kapasite

app.layout = html.Div([
    dcc.Location(id="url", refresh=False),
    dbc.Nav(
        [
            dbc.NavLink("Kapasite Dashboard", href="/", active="exact"),
        ],
        pills=True,
        className="m-3",
    ),
    html.Div(id="page-content"),
])


@app.callback(Output("page-content", "children"), Input("url", "pathname"))
def _render_page(pathname):
    return kapasite.layout
