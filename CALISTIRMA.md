# KAPASITE – Bağımsız çalıştırma

Proje artık **valfapp** veya başka bir dış projeye bağımlı değildir.

## 1. Sanal ortamı aç

```powershell
cd c:\Users\itstajyer\Desktop\KAPASITE
.\venv\Scripts\Activate.ps1
```

## 2. Bağımlılıkları kur (ilk seferde)

```powershell
pip install -r requirements.txt
```

*(pyodbc için bilgisayarda **ODBC Driver 17/18 for SQL Server** yüklü olmalı; veritabanına bağlanacaksanız.)*

## 3. Uygulamayı çalıştır

```powershell
python run.py
```

Tarayıcıda **http://localhost:8050** adresine gidin.

---

**Port değiştirmek:** `run.py` içinde `port=8050` değerini düzenleyin.
