#!/usr/bin/env python
# -*- coding: utf-8 -*-

from datetime import datetime
import os
import sys
import traceback

import raporlama

_RUN_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "run")
if _RUN_DIR not in sys.path:
    sys.path.append(_RUN_DIR)
from agent import ag  # type: ignore[reportMissingImports]


def main():
    started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[report_job] Basladi: {started_at}")

    if ag is None:
        ok, msg = raporlama._send_report_notification_email(
            status="BASARISIZ",
            detail="Veritabani baglantisi yok.",
            created_reports=[],
        )
        print(f"[report_job] Veritabani baglantisi yok. mail_ok={ok} msg={msg}")
        return 1

    try:
        zip_bytes, created = raporlama.build_raporlama_zip(ag)
    except Exception as exc:
        ok, msg = raporlama._send_report_notification_email(
            status="BASARISIZ",
            detail=f"Rapor olusturma hatasi: {exc!r}",
            created_reports=[],
        )
        print(f"[report_job] Hata: {exc!r}")
        traceback.print_exc()
        print(f"[report_job] mail_ok={ok} msg={msg}")
        return 1

    if not zip_bytes:
        ok, msg = raporlama._send_report_notification_email(
            status="BASARISIZ",
            detail="Hicbir veri tipi icin rapor olusturulamadi.",
            created_reports=[],
        )
        print(f"[report_job] ZIP olusmadi. mail_ok={ok} msg={msg}")
        return 1

    ok, msg = raporlama._send_report_notification_email(
        status="BASARILI",
        detail="ZIP dosyasi basariyla olusturuldu.",
        created_reports=created,
        attachment_bytes=zip_bytes,
        attachment_name="kapasite_raporlar.zip",
    )
    print(f"[report_job] Tamamlandi. created={created}")
    print(f"[report_job] mail_ok={ok} msg={msg}")
    return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main())
