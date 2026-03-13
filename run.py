# -*- coding: utf-8 -*-

from app import app

if __name__ == "__main__":
    PORT = 8050
    print(f"Sunucu baslatiliyor: http://127.0.0.1:{PORT}")
    print("Tarayicida bu adresi acin. Durdurmak icin Ctrl+C.")
    app.run(
        host="127.0.0.1",
        port=PORT,
        debug=True,
        use_reloader=False,
    )
