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


def _save_zip_outputs(zip_bytes):
    """ZIP'i tarihli dosya + latest kopyası olarak diske kaydet."""
    root = os.path.dirname(os.path.abspath(__file__))
    out_dir = (os.environ.get("KAP_REPORT_OUT_DIR") or "").strip()
    if not out_dir:
        out_dir = os.path.join(root, "scheduled_reports")
    os.makedirs(out_dir, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dated_name = f"kapasite_raporlar_{ts}.zip"
    dated_path = os.path.join(out_dir, dated_name)
    latest_path = os.path.join(out_dir, "kapasite_raporlar_latest.zip")

    with open(dated_path, "wb") as f:
        f.write(zip_bytes)
    with open(latest_path, "wb") as f:
        f.write(zip_bytes)

    return dated_path, latest_path


def main():
    started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[report_job] Basladi: {started_at}")

    single_cc = (os.environ.get("KAP_REPORT_SINGLE_CC") or "").strip()
    if single_cc:
        raporlama.RAPORLAMA_SINGLE_COSTCENTER = "__AUTO__" if single_cc.lower() == "auto" else single_cc
        print(f"[report_job] Single costcenter modu aktif: {raporlama.RAPORLAMA_SINGLE_COSTCENTER}")

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

    try:
        dated_path, latest_path = _save_zip_outputs(zip_bytes)
        print(f"[report_job] ZIP kaydedildi: {dated_path}")
        print(f"[report_job] ZIP latest: {latest_path}")
    except Exception as exc:
        print(f"[report_job] ZIP diske yazilamadi: {exc!r}")
        traceback.print_exc()
        # Mail gonderimi yine denensin; gorev tamamen durmasin.

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
