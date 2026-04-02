#!/usr/bin/env python3
# -*- coding: utf-8 -*-


from __future__ import annotations

import argparse
import os
import sys
import time


def main(argv: list[str] | None = None) -> int:
    root = os.path.dirname(os.path.abspath(__file__))
    run_dir = os.path.join(root, "run")
    if run_dir not in sys.path:
        sys.path.insert(0, run_dir)

    from agent import ag  # noqa: E402  # type: ignore[reportMissingImports]
    import raporlama  # noqa: E402

    parser = argparse.ArgumentParser(description="Kapasite rapor ZIP üret (dashboard gerektirmez).")
    parser.add_argument(
        "output",
        nargs="?",
        default=None,
        help="ZIP dosya yolu (verilmezse KAP_RAPOR_OUT veya proje kökünde kapasite_raporlar.zip)",
    )
    parser.add_argument(
        "--single-cc",
        dest="single_cc",
        default=None,
        help="Sadece tek costcenter uret. Deger: AUTO veya costcenter adi (ornek: KURUTMA).",
    )
    parser.add_argument(
        "--all-cc",
        dest="all_cc",
        action="store_true",
        help="Tum costcenter'lari uret (single-cc ayarini gecersiz kil).",
    )
    args = parser.parse_args(argv)

    if ag is None:
        print("Hata: agent kullanılamıyor (run/agent veya pyodbc kontrol edin).", file=sys.stderr)
        return 1

    # Hizli test modu: tek costcenter (AUTO => ilk bulunan CC)
    env_single_cc = (os.environ.get("KAP_REPORT_SINGLE_CC") or "").strip()
    single_cc = (args.single_cc or env_single_cc).strip() if (args.single_cc or env_single_cc) else ""
    if args.all_cc:
        raporlama.RAPORLAMA_SINGLE_COSTCENTER = None
        print("[Kapasite] Mod: Tum costcenter", flush=True)
    elif single_cc:
        raporlama.RAPORLAMA_SINGLE_COSTCENTER = "__AUTO__" if single_cc.lower() == "auto" else single_cc
        print(f"[Kapasite] Mod: Tek costcenter ({raporlama.RAPORLAMA_SINGLE_COSTCENTER})", flush=True)
    else:
        raporlama.RAPORLAMA_SINGLE_COSTCENTER = None
        print("[Kapasite] Mod: Tum costcenter", flush=True)

    out = (args.output or os.environ.get("KAP_RAPOR_OUT") or "").strip()
    if not out:
        out = os.path.join(root, "kapasite_raporlar.zip")
    out_abs = os.path.abspath(out)

    print("[Kapasite] Rapor ZIP üretimi başladı.", flush=True)
    print(f"[Kapasite] Hedef: {out_abs}", flush=True)
    print(
        "[Kapasite] Veritabanı ve Excel dosyaları hazırlanıyor (bu süre uzun sürebilir)...",
        flush=True,
    )

    t0 = time.perf_counter()
    try:
        zip_bytes, created = raporlama.build_raporlama_zip(ag)
    except Exception as exc:
        print(f"Hata: rapor oluşturulamadı: {exc}", file=sys.stderr)
        return 1

    build_sec = time.perf_counter() - t0
    print(f"[Kapasite] ZIP belleği hazır ({build_sec:.1f} sn).", flush=True)

    if not zip_bytes:
        print(
            "Hata: ZIP boş (veritabanından cost center gelmedi veya tüm veri tipleri atlandı).",
            file=sys.stderr,
        )
        return 2

    parent = os.path.dirname(out_abs)
    if parent:
        os.makedirs(parent, exist_ok=True)
    print(f"[Kapasite] Diske yazılıyor: {out_abs}", flush=True)
    with open(out_abs, "wb") as f:
        f.write(zip_bytes)

    print(f"Tamam: {out_abs}", flush=True)
    print(f"Boyut: {len(zip_bytes):,} bayt", flush=True)
    print(f"İçerik: {', '.join(created) if created else '-'}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
