#!/usr/bin/env python
# -*- coding: utf-8 -*-

from datetime import datetime
import os
import sys
import traceback
import time

import raporlama

_RUN_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "run")
if _RUN_DIR not in sys.path:
    sys.path.append(_RUN_DIR)
from agent import ag  # type: ignore[reportMissingImports]


def _save_zip_outputs(zip_bytes):
    """ZIP'i sadece tek güncel dosya olarak kaydet.

    Eski tarihli raporları temizleyerek klasör şişmesini önler.
    """
    root = os.path.dirname(os.path.abspath(__file__))
    out_dir = (os.environ.get("KAP_REPORT_OUT_DIR") or "").strip()
    if not out_dir:
        out_dir = os.path.join(root, "scheduled_reports")
    os.makedirs(out_dir, exist_ok=True)

    latest_path = os.path.join(out_dir, "kapasite_raporlar_latest.zip")

    with open(latest_path, "wb") as f:
        f.write(zip_bytes)

    # Eski tarihli birikmiş raporları temizle (yalnızca bizim pattern).
    for name in os.listdir(out_dir):
        if not name.lower().endswith(".zip"):
            continue
        if not name.startswith("kapasite_raporlar_"):
            continue
        if name == "kapasite_raporlar_latest.zip":
            continue
        old_path = os.path.join(out_dir, name)
        try:
            os.remove(old_path)
        except Exception:
            # Temizlik hatası job'ı düşürmesin.
            pass

    return latest_path


def main():
    t0_all = time.perf_counter()
    started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[report_job] Basladi: {started_at}", flush=True)

    single_cc = (os.environ.get("KAP_REPORT_SINGLE_CC") or "").strip()
    if single_cc:
        raporlama.RAPORLAMA_SINGLE_COSTCENTER = "__AUTO__" if single_cc.lower() == "auto" else single_cc
        print(f"[report_job] Single costcenter modu aktif: {raporlama.RAPORLAMA_SINGLE_COSTCENTER}", flush=True)

    if ag is None:
        ok, msg = raporlama._send_report_notification_email(
            status="BASARISIZ",
            detail="Veritabani baglantisi yok.",
            created_reports=[],
        )
        print(f"[report_job] Veritabani baglantisi yok. mail_ok={ok} msg={msg}", flush=True)
        return 1

    try:
        print("[report_job] Asama: ZIP olusturma basliyor...", flush=True)
        t_zip = time.perf_counter()
        zip_bytes, created = raporlama.build_raporlama_zip(ag)
        zip_sec = time.perf_counter() - t_zip
        print(f"[report_job] Asama: ZIP olusturma tamamlandi. sure={zip_sec:.2f}s", flush=True)
    except Exception as exc:
        ok, msg = raporlama._send_report_notification_email(
            status="BASARISIZ",
            detail=f"Rapor olusturma hatasi: {exc!r}",
            created_reports=[],
        )
        print(f"[report_job] Hata: {exc!r}", flush=True)
        traceback.print_exc()
        print(f"[report_job] mail_ok={ok} msg={msg}", flush=True)
        return 1

    if not zip_bytes:
        ok, msg = raporlama._send_report_notification_email(
            status="BASARISIZ",
            detail="Hicbir veri tipi icin rapor olusturulamadi.",
            created_reports=[],
        )
        print(f"[report_job] ZIP olusmadi. mail_ok={ok} msg={msg}", flush=True)
        return 1

    try:
        print("[report_job] Asama: ZIP diske kaydediliyor...", flush=True)
        t_save = time.perf_counter()
        latest_path = _save_zip_outputs(zip_bytes)
        save_sec = time.perf_counter() - t_save
        print(f"[report_job] ZIP kaydedildi: {latest_path}", flush=True)
        print(f"[report_job] ZIP latest: {latest_path}", flush=True)
        print(f"[report_job] Asama: ZIP diske kayit tamamlandi. sure={save_sec:.2f}s", flush=True)
    except Exception as exc:
        print(f"[report_job] ZIP diske yazilamadi: {exc!r}", flush=True)
        traceback.print_exc()
        # Mail gonderimi yine denensin; gorev tamamen durmasin.

    print("[report_job] Asama: E-posta gonderimi basliyor...", flush=True)
    t_mail = time.perf_counter()
    ok, msg = raporlama._send_report_notification_email(
        status="BASARILI",
        detail="ZIP dosyasi basariyla olusturuldu.",
        created_reports=created,
        attachment_bytes=zip_bytes,
        attachment_name="kapasite_raporlar.zip",
    )
    mail_sec = time.perf_counter() - t_mail
    all_sec = time.perf_counter() - t0_all
    print(f"[report_job] Tamamlandi. created={created}", flush=True)
    print(f"[report_job] mail_ok={ok} msg={msg}", flush=True)
    print(f"[report_job] Asama: E-posta tamamlandi. sure={mail_sec:.2f}s", flush=True)
    print(f"[report_job] Toplam sure={all_sec:.2f}s", flush=True)
    return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main())
