#!/usr/bin/env python
# -*- coding: utf-8 -*-

from datetime import datetime
import os
import sys
import traceback
import time

import config
import raporlama

_RUN_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "run")
if _RUN_DIR not in sys.path:
    sys.path.append(_RUN_DIR)
from agent import ag  # type: ignore[reportMissingImports]


def _scheduled_report_output_dir():
    """ZIP çıktı klasörü: KAP_REPORT_OUT_DIR (ortam) varsa o, yoksa config.SCHEDULED_REPORT_OUTPUT_DIR."""
    env = (os.environ.get("KAP_REPORT_OUT_DIR") or "").strip()
    if env:
        return os.path.normpath(os.path.abspath(os.path.expanduser(env)))
    return os.path.normpath(os.path.abspath(config.SCHEDULED_REPORT_OUTPUT_DIR))


def _scheduled_report_zip_fullpath():
    return os.path.join(
        _scheduled_report_output_dir(),
        getattr(config, "SCHEDULED_REPORT_ZIP_BASENAME", "kapasite_raporlar_latest.zip"),
    )


def _save_zip_outputs(zip_bytes):
    """ZIP'i tek güncel dosya olarak kaydet; aynı klasördeki diğer kapasite_raporlar*.zip dosyalarını sil.

    Biriken eski rapor kalmaz: yalnızca SCHEDULED_REPORT_ZIP_BASENAME korunur.
    """
    out_dir = _scheduled_report_output_dir()
    os.makedirs(out_dir, exist_ok=True)

    zip_basename = getattr(config, "SCHEDULED_REPORT_ZIP_BASENAME", "kapasite_raporlar_latest.zip")
    latest_path = os.path.join(out_dir, zip_basename)
    latest_path = os.path.normpath(os.path.abspath(latest_path))

    with open(latest_path, "wb") as f:
        f.write(zip_bytes)

    keep_lower = zip_basename.lower()
    for name in os.listdir(out_dir):
        if not name.lower().endswith(".zip"):
            continue
        if name.lower() == keep_lower:
            continue
        if not name.lower().startswith("kapasite_raporlar"):
            continue
        old_path = os.path.join(out_dir, name)
        try:
            os.remove(old_path)
        except Exception:
            pass

    return latest_path


def _zip_location_hint_success(saved_abs_path, zip_bytes):
    try:
        mb = len(zip_bytes) / (1024 * 1024)
        size_line = f"Dosya boyutu (yaklasik): {mb:.2f} MB"
    except Exception:
        size_line = ""
    lines = [saved_abs_path]
    if size_line:
        lines.append(size_line)
    return "\n".join(lines)


def main():
    t0_all = time.perf_counter()
    started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[report_job] Basladi: {started_at}", flush=True)

    expected_zip_path = _scheduled_report_zip_fullpath()
    out_dir = _scheduled_report_output_dir()
    print(f"[report_job] ZIP hedef klasor: {out_dir}", flush=True)
    print(f"[report_job] ZIP hedef dosya: {expected_zip_path}", flush=True)

    single_cc = (os.environ.get("KAP_REPORT_SINGLE_CC") or "").strip()
    if single_cc:
        raporlama.RAPORLAMA_SINGLE_COSTCENTER = "__AUTO__" if single_cc.lower() == "auto" else single_cc
        print(f"[report_job] Single costcenter modu aktif: {raporlama.RAPORLAMA_SINGLE_COSTCENTER}", flush=True)

    def _mail(status, detail, created, zip_info, exit_code_on_mail_fail):
        ok, msg = raporlama._send_report_notification_email(
            status=status,
            detail=detail,
            created_reports=created or [],
            attachment_bytes=None,
            zip_location_info=zip_info,
            is_scheduled_job_email=True,
        )
        print(f"[report_job] mail_ok={ok} msg={msg}", flush=True)
        return ok, msg, (0 if ok else exit_code_on_mail_fail)

    if ag is None:
        info = (
            "Bu calistirmada ZIP uretilemedi (veritabani yok).\n"
            f"Basarili calistirmalarda rapor su dosyaya yazilir:\n{expected_zip_path}"
        )
        _, _, code = _mail(
            "BASARISIZ",
            "Veritabani baglantisi yok.",
            [],
            info,
            1,
        )
        return code

    zip_bytes = None
    created = []
    try:
        print("[report_job] Asama: ZIP olusturma basliyor...", flush=True)
        t_zip = time.perf_counter()
        zip_bytes, created = raporlama.build_raporlama_zip(ag)
        zip_sec = time.perf_counter() - t_zip
        print(f"[report_job] Asama: ZIP olusturma tamamlandi. sure={zip_sec:.2f}s", flush=True)
    except Exception as exc:
        info = (
            f"Bu calistirmada ZIP uretilemedi (hata).\n"
            f"Hedef dosya (basarili kosuda): {expected_zip_path}"
        )
        _, _, code = _mail(
            "BASARISIZ",
            f"Rapor olusturma hatasi: {exc!r}",
            [],
            info,
            1,
        )
        print(f"[report_job] Hata: {exc!r}", flush=True)
        traceback.print_exc()
        return code

    if not zip_bytes:
        info = (
            "ZIP olusturulamadi (bos veri).\n"
            f"Hedef dosya (basarili kosuda): {expected_zip_path}"
        )
        _, _, code = _mail(
            "BASARISIZ",
            "Hicbir veri tipi icin rapor olusturulamadi.",
            [],
            info,
            1,
        )
        return code

    saved_abs_path = None
    try:
        print("[report_job] Asama: ZIP diske kaydediliyor...", flush=True)
        t_save = time.perf_counter()
        saved_abs_path = _save_zip_outputs(zip_bytes)
        save_sec = time.perf_counter() - t_save
        print(f"[report_job] ZIP kaydedildi: {saved_abs_path}", flush=True)
        print(f"[report_job] Asama: ZIP diske kayit tamamlandi. sure={save_sec:.2f}s", flush=True)
    except Exception as exc:
        print(f"[report_job] ZIP diske yazilamadi: {exc!r}", flush=True)
        traceback.print_exc()
        info = (
            "ZIP bellekte olusturuldu ancak diske yazilamadi.\n"
            f"Hedef yol: {expected_zip_path}\n"
            f"Hata: {exc!r}"
        )
        _, _, code = _mail(
            "BASARISIZ",
            "ZIP diske yazilamadi.",
            created,
            info,
            2,
        )
        return code

    print("[report_job] Asama: E-posta gonderimi basliyor...", flush=True)
    t_mail = time.perf_counter()
    zip_info = _zip_location_hint_success(saved_abs_path, zip_bytes)
    ok, msg = raporlama._send_report_notification_email(
        status="BASARILI",
        detail="ZIP dosyasi diske kaydedildi. (E-postada ek yok; tam yol asagida.)",
        created_reports=created,
        attachment_bytes=None,
        zip_location_info=zip_info,
        is_scheduled_job_email=True,
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
