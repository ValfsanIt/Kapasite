# KAPASITE - Operasyon Checklist

Bu dosya, raporlama ZIP + e-posta + Task Scheduler akışında ileride yapılacak değişiklikler için hızlı rehberdir.

## 1) Alıcı listesi değişikliği

Dosya: `config.py`

- `RAPORLAMA_NOTIFY_TO`
- `RAPORLAMA_NOTIFY_CC`
- `RAPORLAMA_NOTIFY_BCC`

Ornek:

```python
RAPORLAMA_NOTIFY_TO = ["a@valfsan.com.tr", "b@valfsan.com.tr"]
RAPORLAMA_NOTIFY_CC = ["c@valfsan.com.tr"]
RAPORLAMA_NOTIFY_BCC = []
```

## 2) Gonderen (From) degisikligi

Dosya: `config.py`

- `SMTP_USERNAME` alanini guncelle.
- Relay modunda da `From` bu degere gore set edilir.

Not: SMTP relay kurallari bu adresi izinli gonderen olarak kabul etmelidir.

## 3) Mail sunucu/kimlik dogrulama degisikligi

Dosya: `config.py`

- `MAIL_MODE` (`"relay"` veya `"o365"`)
- `mail_server`
- `mail_port`
- `SMTP_USE_STARTTLS`
- `SMTP_REQUIRE_AUTH`
- `SMTP_USERNAME`
- `SMTP_PASSWORD`

## 4) Mail konusu/icerigi degisikligi

Dosya: `raporlama.py`

- Fonksiyon: `_send_report_notification_email(...)`
- `subject` ve `body` burada olusturulur.

## 5) ZIP icindeki raporlar (3 Excel) degisikligi

Dosya: `raporlama.py`

- Degisken: `DATA_TYPES`
- Hangi veri tiplerinin ZIP icine girecegi burada tanimlidir.

## 6) Zamanlanmis gorev (03:00) degisikligi

Kod dosyasi: `report_job.py` (calisan job)
Saat/plan: Windows Task Scheduler

- Saat degisecekse Task Scheduler Trigger guncellenir.
- `report_job.py` tek calistirmada ZIP uretir ve mail gonderir.

## 7) Task Scheduler zorunlu alanlar

- Program/script:
  - `C:\Users\itstajyer\Desktop\KAPASITE\venv\Scripts\python.exe`
- Add arguments:
  - `report_job.py`
- Start in:
  - `C:\Users\itstajyer\Desktop\KAPASITE`
- If task is already running:
  - `Do not start a new instance`

## 8) Sunucu tasima / Python yolu degisirse

- Yeni sunucuda venv python yolunu bul:
  - `<proje>\venv\Scripts\python.exe`
- Task Scheduler `Program/script` alanini bu yeni yol ile guncelle.

## 9) Hizli test

Proje klasorunde manuel test:

```powershell
.\venv\Scripts\python.exe .\report_job.py
```

Beklenen:
- ZIP olusturma loglari
- Mail gonderim sonucu (SMTP/Outlook)

