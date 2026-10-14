"""
Streamlit app (1 file): Perbandingan data Re-invent vs KRI
Jalankan:  streamlit run app.py

Isi file:  1) definisi kolom  2) helper  3) logika perbandingan  4) export Excel
           5) data contoh (dummy)  6) tampilan Streamlit
"""
import hashlib
import io
import re
from collections import defaultdict
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# DEFINISI KOLOM
# ----------------------------------------------------------------------------
REINVENT_COLS = [
    "No", "Hostname", "ip_oam", "ip_service", "Apps", "Apps Priority",
    "App PIC Email", "Environment", "VM/Baremetal", "Site", "rack", "floor",
    "contaiment", "Ansible Status", "NGPAM Status", "XDR Status", "OS Product",
    "Hardware Model", "Coverage Area", "hypervisor group", "Platform",
    "Monitoring tools", "Serial Numbers", "Parent Host", "Power State",
    "Lifecycle",
]

KRI_COLS = [
    "No", "Apps", "Hostname", "IP OAM", "Prod/Nonprod", "Apps Priority",
    "VM/Baremetal", "Brand/Platform", "OS Version", "OS EOL", "PIC TOWER",
    "Configured Ansible", "Ansible Status", "Configured NGPAM", "NGPAM Status",
    "Configured XDR", "XDR Status", "Monitoring Tools",
    "Remark From OSPLA Team", "Remark From NGPAM Team", "Last Update",
    "Timeline Integration OSPLA",
]

# alias nama kolom (typo / variasi penulisan)
REINVENT_ALIASES = {"containment": "contaiment", "VM/Baremeteal": "VM/Baremetal"}
KRI_ALIASES = {"VM/Baremeteal": "VM/Baremetal"}

# mapping kolom KRI -> kolom Re-invent
KRI_TO_REINVENT = {
    "Hostname": "Hostname",
    "IP OAM": "ip_oam",
    "Apps": "Apps",
    "Apps Priority": "Apps Priority",
    "Prod/Nonprod": "Environment",
    "VM/Baremetal": "VM/Baremetal",
    "Brand/Platform": "Platform",
    "OS Version": "OS Product",
    "Ansible Status": "Ansible Status",
    "NGPAM Status": "NGPAM Status",
    "XDR Status": "XDR Status",
    "Monitoring Tools": "Monitoring tools",
}

SYNC_CANDIDATES = [c for c in KRI_TO_REINVENT.values() if c != "Hostname"]
DEFAULT_SYNC = ["Ansible Status", "NGPAM Status", "XDR Status", "Monitoring tools"]
DEFAULT_REMARK_COLS = ["Remark From OSPLA Team", "Remark From NGPAM Team"]

DEFAULT_OPTIONS = dict(
    strategy="auto",              # auto | KRI | Re-invent
    reinvent_date=None,           # date / Timestamp / None
    sync=DEFAULT_SYNC,
    dismantle_kw="dismantle, decom",
    remark_cols=DEFAULT_REMARK_COLS,
    add_note=False,
    split_incomplete=False,
)

# Warna seluruh baris di file utama (jika ada beberapa masalah, dipakai urutan prioritas ini)
COLORS = {
    "ip": "FFE699",         # kuning  : hostname beda, IP sama
    "hip": "BDD7EE",        # biru    : hostname sama, IP beda
    "dup": "E4DFEC",        # ungu    : hostname duplikat
    "inc": "F4B084",        # oranye  : Hostname / VM-Baremetal kosong
    "dismantle": "D9D9D9",  # abu     : dismantle
    "new": "C6E0B4",        # hijau   : baris baru dari KRI
}
# Warna sel yang bermasalah (lebih pekat) - selalu ditampilkan walau satu baris punya banyak masalah
CELL_COLORS = {
    "ip": "FFC000",         # sel ip_oam
    "hip": "5B9BD5",        # sel ip_oam
    "dup": "8E7CC3",        # sel Hostname
    "inc": "ED7D31",        # sel Hostname / VM-Baremetal yang kosong
    "dismantle": "A6A6A6",  # sel Lifecycle
}


# ----------------------------------------------------------------------------
# HELPER
# ----------------------------------------------------------------------------
_INVISIBLE = re.compile(r"[\u200b\u200c\u200d\u2060\ufeff\u00ad]")  # zero-width, BOM, soft-hyphen
_SPACES = re.compile(r"\s+")  # spasi, tab, enter, NBSP, dan spasi unicode lain


def clean(v):
    """Bersihkan nilai sel: buang karakter tak terlihat, rapikan spasi ganda/tab/enter, trim kiri-kanan."""
    if v is None:
        return ""
    if not isinstance(v, str):
        try:
            if pd.isna(v):
                return ""
        except (TypeError, ValueError):
            pass
    s = _INVISIBLE.sub("", str(v))
    s = _SPACES.sub(" ", s).strip()
    return "" if s.lower() in ("nan", "nat", "none") else s


def nospace(v):
    """Hilangkan SEMUA spasi (juga tab / enter / NBSP) dari nilai. 'host1 / host2' -> 'host1/host2'."""
    return _SPACES.sub("", clean(v))


# kolom identitas (hostname & IP): semua spasinya dibuang saat data dibaca
NOSPACE_COLS = {
    "Hostname", "ip_oam", "ip_service", "IP OAM", "IP Services", "IP", "Hostname OS",
    "Hostname Re-invent", "Hostname KRI", "ip_oam Re-invent", "IP OAM KRI",
    "Validasi Final Hostname", "Validasi IP_OAM",
}


def norm_col(c):
    return re.sub(r"[^a-z0-9]", "", str(c).lower())


STATUS_COLS = ["Ansible Status", "NGPAM Status", "XDR Status"]


def norm_status(v):
    """
    Seragamkan penulisan status integrasi:
      integrated / Integrated / INTEGRATED                  -> INTEGRATED
      not integrated / Not-Integrated / non integrated ...  -> NOT_INTEGRATED
    Nilai lain (mis. Installed, Not Installed, kosong) dibiarkan apa adanya.
    """
    s = clean(v)
    key = re.sub(r"[^a-z]", "", s.lower())
    if key in ("integrated", "integrate"):
        return "INTEGRATED"
    if key in ("notintegrated", "notintegrate", "nonintegrated", "unintegrated", "belumintegrated"):
        return "NOT_INTEGRATED"
    return s


def host_key(v):
    """
    Kunci pencocokan hostname = bagian sebelum titik pertama, huruf kecil, tanpa spasi.
      hostname1.telkomsel.co.id -> hostname1
      HostName1.apa.saja.com    -> hostname1
      ' hostname1 '             -> hostname1
      hostname1.                -> hostname1
    Pengecualian: jika isinya alamat IP (10.1.1.1) dipakai utuh.
    """
    s = _SPACES.sub("", _INVISIBLE.sub("", clean(v))).lower().strip(".")
    if not s:
        return ""
    if re.fullmatch(r"\d{1,3}(\.\d{1,3}){3}", s):
        return s
    return s.split(".")[0]


def ip_set(v):
    return set(re.findall(r"\d{1,3}(?:\.\d{1,3}){3}", clean(v)))


def ip_tokens(v):
    """Daftar IP dengan urutan asli, tanpa duplikat."""
    return list(dict.fromkeys(re.findall(r"\d{1,3}(?:\.\d{1,3}){3}", clean(v))))


def split_slash(v):
    """'xxx / xxxx' atau 'xxx/xxxx' -> ['xxx', 'xxxx'] (spasi dibuang, bagian kosong dibuang)."""
    return [p for p in nospace(v).split("/") if p]


def resolve_slash_hostnames(re_df, kri_df):
    """
    Hostname berisi dua nilai dipisah tanda '/'  ( xxx / xxxx  atau  xxx/xxxx ).
    Tiap bagian (depan & belakang '/') dicocokkan dengan hostname di file lawannya
    (aturan titik tetap berlaku: host1.domain.com == host1). Bagian yang cocok dipertahankan,
    bagian lain dihapus. df diubah langsung (in-place).
    return: (DataFrame ringkasan, list log perubahan)
    """
    sides = {"Re-invent": re_df, "KRI": kri_df}
    other = {"Re-invent": "KRI", "KRI": "Re-invent"}
    keys = {n: {host_key(p) for v in d["Hostname"] for p in split_slash(v)} for n, d in sides.items()}
    full = {n: {nospace(v).lower() for v in d["Hostname"]} for n, d in sides.items()}
    rows, log = [], []
    for name, df in sides.items():
        oth = other[name]
        for i, v in df["Hostname"].items():
            if "/" not in v:
                continue
            parts = split_slash(v)
            hit = [p for p in parts if host_key(p) in keys[oth]]
            hit_keys = list(dict.fromkeys(host_key(p) for p in hit))
            if len(parts) <= 1:
                new, status, ket = (parts[0] if parts else ""), "Dirapikan", "Tanda / tanpa pasangan dibuang"
            elif v.lower() in full[oth]:
                new, status, ket = v, "Sama persis", f"Nilai yang sama (dengan /) juga ada di {oth}; dibiarkan"
            elif len(hit_keys) == 1:
                new, status = hit[0], "Dipilih"
                ket = f"Hanya bagian ini yang ada di {oth}; bagian lain dihapus"
            elif len(hit_keys) > 1:
                new, status = hit[0], "Perlu dicek"
                ket = f"Dua bagian sama-sama ada di {oth}; dipakai bagian pertama - cek manual"
            else:
                new, status = "/".join(parts), "Belum terselesaikan"
                ket = f"Tidak ada bagian yang cocok di {oth}; nilai dibiarkan"
            rows.append({"Sumber": name, "Hostname asli": v, "Hasil": new, "Status": status, "Keterangan": ket})
            if new != v:
                df.at[i, "Hostname"] = new
                log.append({"Hostname": new, "Field": "Hostname (tanda /)", "Nilai Re-invent (lama)": v,
                            "Nilai baru": new, "Sumber": name, "Alasan": ket})
    return pd.DataFrame(rows, columns=["Sumber", "Hostname asli", "Hasil", "Status", "Keterangan"]), log


def parse_dates(series):
    """
    Format ISO (2026-10-01 / 2026-10-01 08:00:00) dibaca apa adanya.
    Format lain (mis. 01/10/2026) dibaca day-first (dd/mm/yyyy) sesuai kebiasaan lokal.
    Catatan: dayfirst=True pada string ISO akan salah membalik bulan/hari, jadi dipisah.
    """
    s = series.astype(str).str.strip()
    iso = s.str.match(r"^\d{4}-\d{1,2}-\d{1,2}")
    parts = []
    if iso.any():
        parts.append(pd.to_datetime(s[iso], errors="coerce", format="ISO8601"))
    if (~iso).any():
        rest = s[~iso]
        try:
            parts.append(pd.to_datetime(rest, errors="coerce", dayfirst=True, format="mixed"))
        except (TypeError, ValueError):
            parts.append(pd.to_datetime(rest, errors="coerce", dayfirst=True))
    if not parts:
        return pd.Series(dtype="datetime64[ns]", index=series.index)
    return pd.concat(parts).reindex(series.index)


def canon_map(columns, canon, aliases):
    """{nama kolom asli: nama kolom kanonik} (tanpa peduli huruf besar/kecil, spasi, simbol)."""
    lookup = {norm_col(c): c for c in canon}
    lookup.update({norm_col(k): v for k, v in aliases.items()})
    rename = {}
    for c in columns:
        n = norm_col(c)
        if n in lookup and lookup[n] not in rename.values():
            rename[c] = lookup[n]
    return rename


def standardize(df, canon, aliases):
    """Samakan nama kolom ke daftar kanonik; kolom yang hilang diisi kosong."""
    df = df.rename(columns=canon_map(df.columns, canon, aliases))
    missing = [c for c in canon if c not in df.columns]
    for c in missing:
        df[c] = ""
    df = df[canon].copy()
    for c in df.columns:
        df[c] = df[c].map(nospace if c in NOSPACE_COLS else clean)
    return df.reset_index(drop=True), missing


# ----------------------------------------------------------------------------
# PROSES UTAMA
# ----------------------------------------------------------------------------
def compute(re_df, kri_df, opts=None):
    """
    re_df  : DataFrame Re-invent yang sudah di-standardize()
    kri_df : DataFrame KRI yang sudah di-standardize()
    return : dict berisi main, colors, kri_only, re_only, conf, hip, dismantle,
             incomplete, dup, log, summary
    """
    o = {**DEFAULT_OPTIONS, **(opts or {})}
    re_df = re_df.copy()
    kri_df = kri_df.copy()

    # seragamkan penulisan status (INTEGRATED / NOT_INTEGRATED) di kedua sumber sebelum dibandingkan
    for df in (re_df, kri_df):
        for col in STATUS_COLS:
            df[col] = df[col].map(norm_status)

    # hostname "xxx / xxxx": ambil bagian yang cocok dengan file lawan, bagian lain dihapus
    slash_df, slash_log = resolve_slash_hostnames(re_df, kri_df)

    re_df["_key"] = re_df["Hostname"].map(host_key)
    kri_df["_key"] = kri_df["Hostname"].map(host_key)
    # baris tanpa hostname diberi key unik supaya tidak saling cocok
    for df, tag in ((re_df, "re"), (kri_df, "kri")):
        empty = df["_key"] == ""
        df.loc[empty, "_key"] = [f"__nohost_{tag}_{i}" for i in df.index[empty]]

    kri_df["_dt"] = parse_dates(kri_df["Last Update"])
    reinvent_dt = pd.Timestamp(o["reinvent_date"]) if o["reinvent_date"] else None

    # --- duplikat -----------------------------------------------------------
    kri_sorted = kri_df.sort_values("_dt", kind="stable", na_position="first")
    dup_kri = kri_sorted[kri_sorted.duplicated("_key", keep=False)].copy()
    dup_kri = dup_kri[~dup_kri["_key"].str.startswith("__nohost")]
    kri_u = kri_sorted.drop_duplicates("_key", keep="last").sort_index()  # simpan yang terbaru

    dup_re = re_df[re_df.duplicated("_key", keep=False)].copy()
    dup_re = dup_re[~dup_re["_key"].str.startswith("__nohost")]

    dup_report = pd.concat(
        [
            dup_re[["Hostname", "ip_oam", "Apps", "VM/Baremetal"]].assign(Sumber="Re-invent"),
            dup_kri.rename(columns={"IP OAM": "ip_oam"})[
                ["Hostname", "ip_oam", "Apps", "VM/Baremetal"]
            ].assign(Sumber="KRI (dipakai: Last Update terbaru)"),
        ],
        ignore_index=True,
    )

    re_keys = set(re_df["_key"])
    kri_keys = set(kri_u["_key"])
    kri_only = kri_u[~kri_u["_key"].isin(re_keys)]
    re_only = re_df[~re_df["_key"].isin(kri_keys)]
    kri_idx = kri_u.set_index("_key")

    # --- keputusan nilai "terupdate" ----------------------------------------
    def decide(kv, rv, kdt):
        if not kv:
            return rv, "Re-invent", "KRI kosong"
        if not rv:
            return kv, "KRI", "Re-invent kosong"
        if o["strategy"] == "KRI":
            return kv, "KRI", "Prioritas KRI"
        if o["strategy"] == "Re-invent":
            return rv, "Re-invent", "Prioritas Re-invent"
        # otomatis berdasarkan tanggal
        if pd.notna(kdt) and reinvent_dt is not None:
            if kdt >= reinvent_dt:
                return kv, "KRI", f"KRI Last Update {kdt:%Y-%m-%d} >= data Re-invent {reinvent_dt:%Y-%m-%d}"
            return rv, "Re-invent", f"Re-invent {reinvent_dt:%Y-%m-%d} lebih baru dari KRI {kdt:%Y-%m-%d}"
        return kv, "KRI", "Tanggal tidak tersedia, default KRI"

    log = list(slash_log)
    res = re_df.copy()
    res["_src"] = "Re-invent saja"

    kws = [k.strip().lower() for k in o["dismantle_kw"].split(",") if k.strip()]

    def is_dismantle(*texts):
        t = " ".join(clean(x).lower() for x in texts)
        return any(k in t for k in kws)

    # --- update baris Re-invent yang ada di KRI ------------------------------
    for i, row in res.iterrows():
        k = row["_key"]
        remark_texts = []
        if k in kri_idx.index:
            kr = kri_idx.loc[k]
            res.at[i, "_src"] = "Re-invent + KRI"
            remark_texts = [kr[c] for c in o["remark_cols"]]
            # ip_oam berisi beberapa IP (a/b): pertahankan yang sama dengan KRI saja
            a_tok, b_set = ip_tokens(row["ip_oam"]), ip_set(kr["IP OAM"])
            keep = [t for t in a_tok if t in b_set]
            if len(a_tok) > 1 and 0 < len(keep) < len(a_tok):
                res.at[i, "ip_oam"] = "/".join(keep)
                log.append({"Hostname": row["Hostname"], "Field": "ip_oam (tanda /)",
                            "Nilai Re-invent (lama)": row["ip_oam"], "Nilai baru": "/".join(keep),
                            "Sumber": "KRI", "Alasan": "Hanya IP yang sama dengan KRI dipertahankan"})
            for kcol, rcol in KRI_TO_REINVENT.items():
                if rcol == "Hostname":
                    continue
                kv, rv = kr[kcol], row[rcol]
                if rcol in o["sync"]:
                    new, src, why = decide(kv, rv, kr["_dt"])
                else:  # kolom lain: hanya isi jika Re-invent kosong
                    new, src, why = (rv, "Re-invent", "") if rv else (kv, "KRI", "Re-invent kosong")
                if new.casefold() != rv.casefold():
                    res.at[i, rcol] = new
                    log.append({"Hostname": row["Hostname"], "Field": rcol, "Nilai Re-invent (lama)": rv,
                                "Nilai baru": new, "Sumber": src, "Alasan": why})
        # dismantle
        if is_dismantle(row["Lifecycle"], *remark_texts) and res.at[i, "Lifecycle"] != "dismantle":
            log.append({"Hostname": row["Hostname"], "Field": "Lifecycle", "Nilai Re-invent (lama)": row["Lifecycle"],
                        "Nilai baru": "dismantle", "Sumber": "KRI/Re-invent", "Alasan": "Keyword dismantle ditemukan"})
            res.at[i, "Lifecycle"] = "dismantle"

    # --- tambahkan baris KRI yang tidak ada di Re-invent ---------------------
    n_orig = len(res)
    new_rows = []
    for _, kr in kri_only.iterrows():
        r = {c: "" for c in REINVENT_COLS}
        for kcol, rcol in KRI_TO_REINVENT.items():
            r[rcol] = kr[kcol]
        if is_dismantle(*[kr[c] for c in o["remark_cols"]]):
            r["Lifecycle"] = "dismantle"
        r["_key"] = kr["_key"]
        r["_src"] = "KRI saja (ditambahkan)"
        new_rows.append(r)
    if new_rows:
        res = pd.concat([res, pd.DataFrame(new_rows)], ignore_index=True)

    key_to_idx = defaultdict(list)
    for idx, k in res["_key"].items():
        key_to_idx[k].append(idx)

    flags = defaultdict(set)
    for idx in range(n_orig, len(res)):
        flags[idx].add("new")

    # --- hostname beda, IP sama ---------------------------------------------
    ip_to_re = defaultdict(list)
    for i, row in re_df.iterrows():
        for ip in ip_set(row["ip_oam"]):
            ip_to_re[ip].append(i)

    conf, seen = [], set()
    for _, kr in kri_u.iterrows():
        for ip in ip_set(kr["IP OAM"]):
            for i in ip_to_re.get(ip, []):
                if re_df.at[i, "_key"] == kr["_key"]:
                    continue
                sig = (ip, i, kr["_key"])
                if sig in seen:
                    continue
                seen.add(sig)
                conf.append({
                    "IP": ip,
                    "Hostname Re-invent": re_df.at[i, "Hostname"],
                    "Hostname KRI": kr["Hostname"],
                    "VM/Baremetal Re-invent": re_df.at[i, "VM/Baremetal"],
                    "VM/Baremetal KRI": kr["VM/Baremetal"],
                    "Validasi Final Hostname": "",
                })
                flags[i].add("ip")
                for j in key_to_idx.get(kr["_key"], []):
                    flags[j].add("ip")
    conf_df = pd.DataFrame(conf, columns=[
        "IP", "Hostname Re-invent", "Hostname KRI", "VM/Baremetal Re-invent",
        "VM/Baremetal KRI", "Validasi Final Hostname"])

    # --- hostname sama, IP beda ---------------------------------------------
    hip = []
    for i, row in re_df.iterrows():
        k = row["_key"]
        if k in kri_idx.index:
            a, b = ip_set(row["ip_oam"]), ip_set(kri_idx.at[k, "IP OAM"])
            if a and b and a.isdisjoint(b):
                vm_re, vm_kri = row["VM/Baremetal"], kri_idx.at[k, "VM/Baremetal"]
                ket = "Hostname sama tetapi IP berbeda - cek IP mana yang benar (file utama mempertahankan IP Re-invent)"
                if vm_re and vm_kri and vm_re.casefold() != vm_kri.casefold():
                    ket += "; VM/Baremetal juga berbeda"
                hip.append({"Hostname Re-invent": row["Hostname"], "Hostname KRI": kri_idx.at[k, "Hostname"],
                            "ip_oam Re-invent": row["ip_oam"], "IP OAM KRI": kri_idx.at[k, "IP OAM"],
                            "VM/Baremetal Re-invent": vm_re, "VM/Baremetal KRI": vm_kri,
                            "Keterangan": ket, "Validasi IP_OAM": ""})
                flags[i].add("hip")
    hip_df = pd.DataFrame(hip, columns=["Hostname Re-invent", "Hostname KRI", "ip_oam Re-invent", "IP OAM KRI",
                                        "VM/Baremetal Re-invent", "VM/Baremetal KRI", "Keterangan",
                                        "Validasi IP_OAM"])

    # --- hostname duplikat (di Re-invent maupun KRI) -> tandai baris terkait ---
    for k in set(dup_re["_key"]) | set(dup_kri["_key"]):
        for j in key_to_idx.get(k, []):
            flags[j].add("dup")

    # --- kelengkapan wajib: Hostname & VM/Baremetal --------------------------
    for idx, row in res.iterrows():
        if not row["Hostname"] or not row["VM/Baremetal"]:
            flags[idx].add("inc")

    inc_idx = [i for i in res.index if "inc" in flags[i]]
    excluded = set(inc_idx) if o["split_incomplete"] else set()

    # --- bentuk file utama ---------------------------------------------------
    # nomor, warna & catatan dihitung di finalize_main() supaya bisa diulang setelah duplikat dihapus
    main_res = res[~res.index.isin(excluded)]
    base = main_res[REINVENT_COLS].copy().reset_index(drop=True)
    dup_re_keys, dup_kri_keys = set(dup_re["_key"]), set(dup_kri["_key"])
    meta = []
    for idx in main_res.index:
        k = res.at[idx, "_key"]
        meta.append(dict(
            key=k, flags=set(flags[idx]) - {"dup", "inc"},
            dup_re=k in dup_re_keys, dup_kri=k in dup_kri_keys,
            host_empty=not res.at[idx, "Hostname"], vm_empty=not res.at[idx, "VM/Baremetal"],
            dismantle=res.at[idx, "Lifecycle"] == "dismantle",
        ))
    M = dict(base=base, meta=meta, add_note=o["add_note"])
    main, colors, cells = finalize_main(M)

    incomplete_df = res.loc[inc_idx, REINVENT_COLS].copy()
    incomplete_df.insert(0, "Masalah", [
        ", ".join(x for x, c in (("Hostname kosong", not res.at[i, "Hostname"]),
                                 ("VM/Baremetal kosong", not res.at[i, "VM/Baremetal"])) if c) for i in inc_idx])
    dismantle_df = res[res["Lifecycle"] == "dismantle"][REINVENT_COLS].copy()

    kri_only_df = kri_only[KRI_COLS].copy()
    re_only_df = re_only[REINVENT_COLS].copy()
    log_df = pd.DataFrame(log, columns=["Hostname", "Field", "Nilai Re-invent (lama)", "Nilai baru", "Sumber", "Alasan"])

    summary = pd.DataFrame([
        ("Total baris Re-invent", len(re_df)),
        ("Total baris KRI", len(kri_df)),
        ("Hostname ada di keduanya", len(re_keys & kri_keys)),
        ("Ada di KRI, tidak ada di Re-invent", len(kri_only)),
        ("Ada di Re-invent, tidak ada di KRI", len(re_only)),
        ("Hostname beda, IP sama", len(conf_df)),
        ("Hostname sama, IP beda", len(hip_df)),
        ("Lifecycle = dismantle", len(dismantle_df)),
        ("Hostname / VM-Baremetal kosong", len(inc_idx)),
        ("Duplikat hostname", len(dup_report)),
        ("Hostname bertanda / (diproses)", len(slash_df)),
        ("Jumlah perubahan nilai", len(log_df)),
        ("Total baris file utama", len(main)),
    ], columns=["Keterangan", "Jumlah"])

    return dict(main=main, colors=colors, cells=cells, kri_only=kri_only_df, re_only=re_only_df, conf=conf_df,
                hip=hip_df, dismantle=dismantle_df, incomplete=incomplete_df, dup=dup_report,
                log=log_df, summary=summary, slash=slash_df, M=M)


def finalize_main(M, deleted=()):
    """
    Bentuk file utama dari hasil compute(): buang baris yang dipilih untuk dihapus (posisi di M["base"]),
    nomori ulang, lalu hitung warna & catatan. Warna "duplikat" otomatis hilang kalau grupnya
    sudah tinggal 1 baris.
    return: (DataFrame, {pos: warna_baris}, [(pos, kolom, warna_sel)])
    """
    deleted = set(deleted)
    base, meta = M["base"], M["meta"]
    keep = [p for p in range(len(base)) if p not in deleted]
    left = defaultdict(int)
    for p in keep:
        if meta[p]["dup_re"]:
            left[meta[p]["key"]] += 1
    main = base.iloc[keep].reset_index(drop=True)
    main["No"] = list(range(1, len(main) + 1))

    colors, cells, notes = {}, [], []
    for pos, p in enumerate(keep):
        m = meta[p]
        f = set(m["flags"])
        if m["dup_kri"] or (m["dup_re"] and left[m["key"]] >= 2):
            f.add("dup")
        if m["host_empty"] or m["vm_empty"]:
            f.add("inc")
        for key in ("ip", "hip", "dup", "inc"):          # urutan = prioritas warna baris
            if key in f:
                colors[pos] = COLORS[key]
                break
        else:
            if m["dismantle"]:
                colors[pos] = COLORS["dismantle"]
            elif "new" in f:
                colors[pos] = COLORS["new"]
        if "ip" in f:
            cells.append((pos, "ip_oam", CELL_COLORS["ip"]))
        if "hip" in f:
            cells.append((pos, "ip_oam", CELL_COLORS["hip"]))
        if "dup" in f:
            cells.append((pos, "Hostname", CELL_COLORS["dup"]))
        if m["host_empty"]:
            cells.append((pos, "Hostname", CELL_COLORS["inc"]))
        if m["vm_empty"]:
            cells.append((pos, "VM/Baremetal", CELL_COLORS["inc"]))
        if m["dismantle"]:
            cells.append((pos, "Lifecycle", CELL_COLORS["dismantle"]))
        out = []
        if "ip" in f:
            out.append("Hostname beda/IP sama")
        if "hip" in f:
            out.append("Hostname sama/IP beda")
        if "dup" in f:
            out.append("Hostname duplikat")
        if "inc" in f:
            out.append("Hostname/VM-Baremetal kosong")
        if m["dismantle"]:
            out.append("Dismantle")
        if "new" in f:
            out.append("Baru dari KRI")
        notes.append("; ".join(out))
    if M["add_note"]:
        main["Catatan"] = notes
    return main, colors, cells


DUP_VIEW_COLS = ["Hostname", "ip_oam", "Apps", "VM/Baremetal", "Environment", "Site", "Lifecycle"]


def dup_rows_df(M):
    """Baris file utama yang hostname-nya duplikat. Index = posisi baris di M["base"]."""
    base, meta = M["base"], M["meta"]
    pos = [p for p, m in enumerate(meta) if m["dup_re"] or m["dup_kri"]]
    df = base.iloc[pos][["No"] + DUP_VIEW_COLS].rename(columns={"No": "No asal"}).copy()
    df.insert(0, "Grup", [meta[p]["key"] for p in pos])
    df.insert(1, "Sumber", ["Re-invent (duplikat)" if meta[p]["dup_re"] else "KRI duplikat (dipakai Last Update terbaru)"
                            for p in pos])
    return df.sort_values("Grup", kind="stable")


# ----------------------------------------------------------------------------
# EXPORT EXCEL
# ----------------------------------------------------------------------------
MAIN_SHEET = "Reinvent_Updated"


def to_excel(sheets, highlights=None, cell_highlights=None):
    """
    sheets          : {nama_sheet: DataFrame}
    highlights      : {nama_sheet: {posisi_baris: hex_warna}}          -> warna satu baris penuh
    cell_highlights : {nama_sheet: [(posisi_baris, nama_kolom, hex)]}  -> warna sel tertentu
    """
    highlights = highlights or {}
    cell_highlights = cell_highlights or {}
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, df in sheets.items():
            name = name[:31]
            df.to_excel(w, sheet_name=name, index=False)
            ws = w.sheets[name]
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.font = Font(bold=True)
            for ci, col in enumerate(df.columns, 1):
                longest = max([len(str(col))] + [len(str(v)) for v in df[col].head(500)])
                ws.column_dimensions[get_column_letter(ci)].width = min(longest + 2, 50)
            for pos, color in highlights.get(name, {}).items():
                fill = PatternFill("solid", start_color=color, end_color=color)
                for c in ws[pos + 2]:
                    c.fill = fill
            cols = list(df.columns)
            for pos, col, color in cell_highlights.get(name, []):
                if col in cols:
                    ws.cell(row=pos + 2, column=cols.index(col) + 1).fill = PatternFill(
                        "solid", start_color=color, end_color=color)
    return buf.getvalue()


def build_outputs(result, deleted=()):
    """Bangun 2 file Excel (bytes): file utama & file laporan. deleted = posisi baris duplikat yang dihapus."""
    deleted = set(deleted)
    main, colors, cells = finalize_main(result["M"], deleted)
    main_bytes = to_excel({MAIN_SHEET: main}, {MAIN_SHEET: colors}, {MAIN_SHEET: cells})
    summary = result["summary"].copy()
    summary.loc[summary["Keterangan"] == "Total baris file utama", "Jumlah"] = len(main)
    summary.loc[len(summary)] = ("Baris duplikat dihapus (pilihan manual)", len(deleted))
    deleted_df = result["M"]["base"].iloc[sorted(deleted)][REINVENT_COLS].copy()
    report_sheets = {
        "Ringkasan": summary,
        "KRI_Tidak_Ada_di_Reinvent": result["kri_only"],
        "Reinvent_Tidak_Ada_di_KRI": result["re_only"],
        "IP_Sama_Hostname_Beda": result["conf"],
        "Hostname_Sama_IP_Beda": result["hip"],
        "Dismantle": result["dismantle"],
        "Perlu_Dilengkapi": result["incomplete"],
        "Duplikat": result["dup"],
        "Duplikat_Dihapus": deleted_df,
        "Hostname_Slash": result["slash"],
        "Log_Perubahan": result["log"],
    }
    return main_bytes, to_excel(report_sheets)


# ----------------------------------------------------------------------------
# DATA CONTOH (DUMMY) - dibuat di memori, tidak butuh file
# ----------------------------------------------------------------------------
def _re_row(no, host, ip, apps="App-A", vm="VM", **kw):
    r = {c: "" for c in REINVENT_COLS}
    r.update({"No": no, "Hostname": host, "ip_oam": ip, "Apps": apps, "VM/Baremetal": vm,
              "Environment": "Prod", "Site": "JKT-01", "Ansible Status": "Not Integrated",
              "NGPAM Status": "Not Integrated", "XDR Status": "Not Installed",
              "Monitoring tools": "Zabbix", "Lifecycle": "Active"})
    r.update(kw)
    return r


def _kri_row(no, host, ip, apps="App-A", vm="VM", **kw):
    r = {c: "" for c in KRI_COLS}
    r.update({"No": no, "Hostname": host, "IP OAM": ip, "Apps": apps, "VM/Baremetal": vm,
              "Prod/Nonprod": "Prod", "Brand/Platform": "VMware", "OS Version": "RHEL 8",
              "Ansible Status": "Integrated", "NGPAM Status": "Integrated", "XDR Status": "Installed",
              "Monitoring Tools": "Zabbix", "Last Update": "2026-09-30"})
    r.update(kw)
    return r


def sample_data():
    reinvent = [
        _re_row(1, "srv-app01", "10.10.1.1"),                       # FQDN di KRI -> match
        _re_row(2, "SRV-DB01", "10.10.1.2", apps="DB"),             # beda huruf besar/kecil
        _re_row(3, "srv-web01", "10.10.1.3", apps="Web"),           # hanya di Re-invent
        _re_row(4, "srv-old01", "10.10.1.4", apps="Legacy"),        # remark KRI = dismantle
        _re_row(5, "srv-conflict-a", "10.10.1.50", apps="Billing"), # IP sama, hostname beda
        _re_row(6, "srv-nokri", "10.10.1.6", vm=""),                # VM/Baremetal kosong
        _re_row(7, "srv-ipdiff", "10.10.1.7"),                      # hostname sama, IP beda
    ]
    kri = [
        _kri_row(1, "srv-app01.telkomsel.co.id", "10.10.1.1", **{"Last Update": "2026-10-01"}),
        _kri_row(2, "srv-db01.telkomsel.co.id", "10.10.1.2", apps="DB",
                 **{"Last Update": "2026-08-01", "Ansible Status": "Not Integrated"}),
        _kri_row(3, "srv-old01", "10.10.1.4", apps="Legacy",
                 **{"Remark From OSPLA Team": "Server sudah DISMANTLE Sept 2026"}),
        _kri_row(4, "srv-new01", "10.10.1.100", apps="Portal"),     # hanya di KRI
        _kri_row(5, "srv-conflict-b", "10.10.1.50", apps="Billing-DR"),
        _kri_row(6, "srv-ipdiff", "10.10.9.9"),
        _kri_row(7, "srv-novm", "10.10.1.101", vm=""),
        _kri_row(8, "srv-app01", "10.10.1.1", **{"Last Update": "2026-05-01"}),  # duplikat (lebih lama)
        _kri_row(9, "srv-web01 / srv-zzz", "10.10.1.3", apps="Web"),  # tanda / -> srv-web01 dipilih
    ]
    return pd.DataFrame(reinvent, columns=REINVENT_COLS), pd.DataFrame(kri, columns=KRI_COLS)


# ----------------------------------------------------------------------------
# TAMPILAN STREAMLIT
# ----------------------------------------------------------------------------
DATA_DIR = Path(__file__).parent / "data"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
P1_KEY = "p1_result"   # hasil Halaman 1 (setelah hapus duplikat) -> dipakai otomatis di Halaman 2


def read_any(src, label, prefer_sheet=None):
    """src: UploadedFile / Path / DataFrame. Mendukung xlsx, xls, csv; xlsx multi-sheet bisa dipilih."""
    if isinstance(src, pd.DataFrame):
        return src
    name = (src.name if hasattr(src, "name") else str(src)).lower()
    if name.endswith(".csv"):
        return pd.read_csv(src, dtype=str, keep_default_na=False)
    xl = pd.ExcelFile(src)
    sheet = prefer_sheet if prefer_sheet in xl.sheet_names else xl.sheet_names[0]
    if len(xl.sheet_names) > 1 and prefer_sheet not in xl.sheet_names:
        sheet = st.sidebar.selectbox(f"Sheet {label}", xl.sheet_names, key=f"sheet_{label}")
    return xl.parse(sheet, dtype=str)


def find_file(folder, keyword):
    if not folder.exists():
        return None
    for p in sorted(folder.iterdir()):
        if keyword in p.name.lower() and p.suffix.lower() in (".xlsx", ".xls", ".csv") and not p.name.startswith("~$"):
            return p
    return None


# ----------------------------------------------------------------------------
# MENU 2: TERAPKAN VALIDASI (PERBAIKAN) DARI FILE EXCEL HASIL VALIDASI
#
# Kedua menu saling berkorelasi:
#   - Jika baris sudah dihapus / hostname sudah diganti lewat validasi "IP sama, hostname beda",
#     maka entri "Hostname sama, IP beda" untuk hostname itu otomatis GUGUR (tidak diproses, tidak ditandai).
#   - Jika IP sudah diganti lewat validasi "Hostname sama, IP beda" sehingga konflik IP sudah tidak ada,
#     maka entri "IP sama, hostname beda" untuk IP itu otomatis GUGUR.
#   Berlaku dua arah, baik dalam satu sesi (urutan bisa dipilih) maupun lintas sesi
#   (upload Laporan_Validasi.xlsx sebelumnya sebagai riwayat).
# ----------------------------------------------------------------------------
MENU_IP = "IP sama, hostname beda"
MENU_HOST = "Hostname sama, IP beda"

# kolom template file validasi
VAL_IP_COLS = ["IP", "Hostname Re-invent", "Hostname KRI", "VM/Baremetal Re-invent",
               "VM/Baremetal KRI", "Validasi Final Hostname"]
VAL_IP_ALIASES = {"VM/Baremeteal Re-invent": "VM/Baremetal Re-invent",
                  "VM/Baremeteal KRI": "VM/Baremetal KRI",
                  "Validasi Final": "Validasi Final Hostname"}
VAL_HOST_COLS = ["Hostname Re-invent", "Hostname KRI", "ip_oam Re-invent", "IP OAM KRI",
                 "VM/Baremetal Re-invent", "VM/Baremetal KRI", "Keterangan", "Validasi IP_OAM"]
VAL_HOST_ALIASES = {"VM/Baremeteal Re-invent": "VM/Baremetal Re-invent",
                    "VM/Baremeteal KRI": "VM/Baremetal KRI",
                    "Validasi IP": "Validasi IP_OAM", "Validasi OAM": "Validasi IP_OAM"}
LOG_COLS = ["Menu", "Hostname", "Field", "Nilai lama", "Nilai baru", "Keterangan"]

# warna penanda "validasi masih kosong / belum divalidasi"
MARK_ROW = "FFC7CE"
MARK_CELL = "FF5050"

IP_FILLS = {COLORS["ip"], CELL_COLORS["ip"]}
HIP_FILLS = {COLORS["hip"], CELL_COLORS["hip"]}
INC_FILLS = {COLORS["inc"], CELL_COLORS["inc"]}


def cell_rgb(c):
    """Warna isian (solid) sebuah sel openpyxl -> 'RRGGBB' atau None."""
    f = c.fill
    if f is None or f.fill_type != "solid":
        return None
    v = getattr(f.start_color, "rgb", None)
    if isinstance(v, str) and len(v) in (6, 8):
        return v[-6:].upper()
    return None


def read_prev_main(src):
    """
    Baca file hasil olahan sebelumnya (Reinvent_Updated.xlsx) beserta warna selnya.
    Return: (DataFrame kolom Re-invent, {nomor_baris: {kolom: warna}}, kolom_hilang)
    """
    name = (src.name if hasattr(src, "name") else str(src)).lower()
    fills = {}
    if name.endswith(".xlsx"):
        raw_bytes = src.getvalue() if hasattr(src, "getvalue") else Path(src).read_bytes()
        wb = load_workbook(io.BytesIO(raw_bytes), data_only=True)
        ws = wb[MAIN_SHEET] if MAIN_SHEET in wb.sheetnames else wb.worksheets[0]
        rows = list(ws.iter_rows())
        header = [clean(c.value) for c in rows[0]] if rows else []
        cmap = canon_map(header, REINVENT_COLS, REINVENT_ALIASES)
        data = []
        for r in rows[1:]:
            vals = [clean(c.value) for c in r]
            if not any(vals):
                continue
            fl = {}
            for i, c in enumerate(r):
                col = cmap.get(header[i]) if i < len(header) else None
                rgb = cell_rgb(c)
                if col and rgb:
                    fl[col] = rgb
            fills[len(data)] = fl
            data.append(vals)
        raw = pd.DataFrame(data, columns=header)
    else:
        raw = read_any(src, "File sebelumnya")
    df, missing = standardize(raw, REINVENT_COLS, REINVENT_ALIASES)
    return df, fills, missing


def read_history(src):
    """
    Riwayat validasi sebelumnya dari Laporan_Validasi.xlsx (sheet Log_Perubahan):
      gone    : {kunci hostname: alasan}  untuk baris yang sudah dihapus / hostname yang sudah diganti
      ip_hist : {kunci hostname: {IP lama}} untuk IP yang sudah diganti
    """
    df = pd.read_excel(src, sheet_name="Log_Perubahan", dtype=str)
    df, _ = standardize(df, LOG_COLS, {})
    gone, ip_hist = {}, defaultdict(set)
    for _, r in df.iterrows():
        f = r["Field"]
        if f == "(baris dihapus)":
            gone[host_key(r["Hostname"])] = f"sudah dihapus lewat validasi sebelumnya ({r['Menu']})"
        elif f == "Hostname" and host_key(r["Nilai lama"]) != host_key(r["Nilai baru"]):
            gone[host_key(r["Nilai lama"])] = f"sudah diganti menjadi {r['Nilai baru']} lewat validasi sebelumnya"
        elif f == "ip_oam":
            ip_hist[host_key(r["Hostname"])] |= ip_set(r["Nilai lama"]) - ip_set(r["Nilai baru"])
    return dict(gone=gone, ip_hist=ip_hist)


def _new_state(df, fills, history=None):
    df = df.copy()
    df["_key"] = df["Hostname"].map(host_key)
    S = dict(df=df, fills=dict(fills), ips={i: ip_set(v) for i, v in df["ip_oam"].items()},
             marks={}, deleted=set(), gone={}, ip_hist=defaultdict(set), new_ips={},
             log=[], notfound=[], gugur=[], warns=[])
    if history:
        S["gone"].update(history["gone"])
        for k, v in history["ip_hist"].items():
            S["ip_hist"][k] |= v
    return S


def _log(S, menu, host, field, old, new, ket=""):
    S["log"].append({"Menu": menu, "Hostname": host, "Field": field,
                     "Nilai lama": old, "Nilai baru": new, "Keterangan": ket})


def _gugur(S, menu, ip, host, reason):
    S["gugur"].append({"Menu": menu, "IP": ip, "Hostname": host, "Alasan": reason})


def _mark(S, idx, reason, col, menu, kind, ip="", keys=()):
    """Tandai baris (merah muda). kind: 'ip' = konflik IP sama; 'hostip' = hostname sama IP beda."""
    if idx in S["marks"]:
        return
    S["marks"][idx] = dict(reason=reason, col=col, menu=menu, kind=kind, ip=ip, keys=set(keys),
                           host=S["df"].at[idx, "Hostname"], fills=S["fills"].pop(idx, {}))


def _unmark(S, idx, remove):
    m = S["marks"].pop(idx, None)
    if m:
        S["fills"][idx] = {k: v for k, v in m["fills"].items() if v not in remove}


def _strip(S, idx, remove):
    f = S["fills"].get(idx)
    if f:
        S["fills"][idx] = {k: v for k, v in f.items() if v not in remove}


def _complete_cleanup(S, idx):
    """Jika Hostname & VM/Baremetal sudah terisi, hapus warna 'kosong' (oranye)."""
    df = S["df"]
    if df.at[idx, "Hostname"] and df.at[idx, "VM/Baremetal"]:
        _strip(S, idx, INC_FILLS)


def _fill_vm(S, idx, options, menu):
    df = S["df"]
    if df.at[idx, "VM/Baremetal"]:
        return
    for v in options:
        v = clean(v)
        if v:
            df.at[idx, "VM/Baremetal"] = v
            _log(S, menu, df.at[idx, "Hostname"], "VM/Baremetal", "", v, "Diisi dari data validasi / baris lain")
            return


def apply_ip_validation(S, val):
    """
    IP sama, hostname beda.
    Baris yang bentrok dihapus dan diganti SATU baris berhostname 'Validasi Final Hostname'.
    Validasi kosong -> ditandai. Konflik yang sudah tidak ada (mis. IP sudah diganti lewat menu
    'Hostname sama, IP beda') -> entri GUGUR.
    """
    menu = MENU_IP
    df = S["df"]
    flagged = {i for i, f in S["fills"].items() if set(f.values()) & IP_FILLS}  # bentrok dari hasil lama

    ip_idx = defaultdict(list)
    for i in df.index:
        for ip in S["ips"].get(i, ()):
            ip_idx[ip].append(i)

    groups = {}
    for _, r in val.iterrows():
        ips = sorted(ip_set(r["IP"]))
        if not ips:
            if any(clean(r[c]) for c in val.columns):
                S["warns"].append({"Menu": menu, "Detail": "IP kosong/tidak valid di file validasi "
                                   f"(Hostname Re-invent={r['Hostname Re-invent']}, KRI={r['Hostname KRI']})"})
            continue
        groups.setdefault(ips[0], []).append(r)
    validated_ips = set(groups)

    for ip, rows in groups.items():
        keys = set()
        for r in rows:
            keys |= {host_key(r["Hostname Re-invent"]), host_key(r["Hostname KRI"])}
        keys.discard("")
        re_keys = {host_key(r["Hostname Re-invent"]) for r in rows} - {""}
        names = " | ".join(sorted({clean(r[c]) for r in rows for c in ("Hostname Re-invent", "Hostname KRI")
                                   if clean(r[c])}))
        cand = [i for i in ip_idx.get(ip, []) if i not in S["deleted"] and df.at[i, "_key"] in keys]
        distinct = {df.at[i, "_key"] for i in cand}
        hist_keys = sorted(k for k in keys if ip in S["ip_hist"].get(k, set()))
        gone_keys = [k for k in sorted(keys) if k in S["gone"]]

        if not cand:
            if hist_keys:
                _gugur(S, menu, ip, names, f"GUGUR: IP hostname {', '.join(hist_keys)} sudah diganti lewat validasi "
                       f"{MENU_HOST}")
            elif gone_keys:
                _gugur(S, menu, ip, names, f"GUGUR: hostname {gone_keys[0]} {S['gone'][gone_keys[0]]}")
            else:
                S["notfound"].append({"Menu": menu, "IP": ip, "Hostname": names,
                                      "Alasan": "IP + hostname tidak ditemukan di file hasil sebelumnya "
                                                "(salah ketik, atau sudah diselesaikan - upload Laporan_Validasi "
                                                "sebelumnya sebagai riwayat)"})
            continue

        if len(distinct) < 2:   # hanya satu hostname yang masih memakai IP ini -> konflik sudah tidak ada
            if hist_keys:
                why = f"IP hostname {', '.join(hist_keys)} sudah diganti lewat validasi {MENU_HOST}"
            else:
                why = "konflik sudah tidak ada di file (hanya satu hostname yang memakai IP ini)"
            _gugur(S, menu, ip, names, f"GUGUR: {why}")
            for i in cand:
                _strip(S, i, IP_FILLS)
                _complete_cleanup(S, i)
            continue

        finals = [clean(r["Validasi Final Hostname"]) for r in rows if clean(r["Validasi Final Hostname"])]
        if not finals:
            for i in cand:
                _mark(S, i, "Validasi Final Hostname masih kosong", "Hostname", menu, "ip", ip, distinct)
            continue
        final = finals[0]
        fk = host_key(final)
        if len({host_key(f) for f in finals}) > 1:
            S["warns"].append({"Menu": menu, "Detail": f"IP {ip}: validasi final berbeda-beda "
                               f"({', '.join(sorted(set(finals)))}); dipakai: {final}"})

        survivor = next((i for i in cand if df.at[i, "_key"] == fk), None)
        if survivor is None:
            survivor = next((i for i in cand if df.at[i, "_key"] in re_keys), cand[0])
        others = [i for i in cand if i != survivor]

        vm_options = [df.at[o, "VM/Baremetal"] for o in others]
        for r in rows:
            for hc, vc in (("Hostname Re-invent", "VM/Baremetal Re-invent"), ("Hostname KRI", "VM/Baremetal KRI")):
                if host_key(r[hc]) == fk:
                    vm_options.insert(0, r[vc])
            vm_options += [r["VM/Baremetal Re-invent"], r["VM/Baremetal KRI"]]

        old_host, old_key = df.at[survivor, "Hostname"], df.at[survivor, "_key"]
        if old_host != final:
            df.at[survivor, "Hostname"] = final
            _log(S, menu, final, "Hostname", old_host, final, f"IP {ip}: diganti sesuai Validasi Final Hostname")
        if old_key != fk:
            S["gone"][old_key] = f"sudah diganti menjadi {final} lewat validasi {menu}"
        df.at[survivor, "_key"] = fk
        _fill_vm(S, survivor, vm_options, menu)
        for o in others:
            ok = df.at[o, "_key"]
            _log(S, menu, df.at[o, "Hostname"], "(baris dihapus)", df.at[o, "Hostname"], "",
                 f"IP {ip}: baris bentrok dihapus, diganti hostname final {final}")
            if ok != fk:
                S["gone"][ok] = f"sudah dihapus lewat validasi {menu} (diganti {final})"
            S["deleted"].add(o)
            S["fills"].pop(o, None)
            m = S["marks"].pop(o, None)
            if m:
                _gugur(S, m["menu"], m["ip"], m["host"], f"GUGUR: baris dihapus lewat validasi {menu}")
        _strip(S, survivor, IP_FILLS)
        _complete_cleanup(S, survivor)

        if ((df["_key"] == fk) & ~df.index.isin(cand) & ~df.index.isin(list(S["deleted"]))).any():
            S["warns"].append({"Menu": menu, "Detail": f"IP {ip}: hostname final '{final}' ternyata sudah ada di "
                               "baris lain pada file sebelumnya (kemungkinan duplikat)"})

    S["df"] = df.drop(index=list(S["deleted"]))
    df = S["df"]

    # bandingkan dengan data sebelumnya: konflik IP yang masih tersisa tapi tidak ada di file validasi
    ip_rows = defaultdict(list)
    for i in df.index:
        for ip in S["ips"].get(i, ()):
            ip_rows[ip].append(i)
    for ip, idxs in ip_rows.items():
        if ip in validated_ips:
            continue
        use = [i for i in idxs if i in flagged] if flagged else idxs
        distinct = {df.at[i, "_key"] for i in use}
        if len(distinct) > 1:
            for i in use:
                _mark(S, i, f"IP {ip} sama, hostname beda - belum ada di file validasi", "Hostname",
                      menu, "ip", ip, distinct)


def apply_ip_oam_validation(S, val):
    """
    Hostname sama, IP beda.
    ip_oam diganti dengan 'Validasi IP_OAM'. Validasi kosong -> ditandai. Hostname yang sudah
    dihapus/diganti lewat menu 'IP sama, hostname beda' -> entri GUGUR.
    """
    menu = MENU_HOST
    df = S["df"]
    key_idx = defaultdict(list)
    for i in df.index:
        if df.at[i, "_key"]:
            key_idx[df.at[i, "_key"]].append(i)

    for _, r in val.iterrows():
        keys = {host_key(r["Hostname Re-invent"]), host_key(r["Hostname KRI"])} - {""}
        if not keys:
            continue
        names = " | ".join(sorted({clean(r[c]) for c in ("Hostname Re-invent", "Hostname KRI") if clean(r[c])}))
        cand = list(dict.fromkeys(i for k in keys for i in key_idx.get(k, []) if i not in S["deleted"]))
        if not cand:
            gone_keys = [k for k in sorted(keys) if k in S["gone"]]
            ip_txt = clean(r["ip_oam Re-invent"])
            if gone_keys:
                _gugur(S, menu, ip_txt, names, f"GUGUR: hostname {gone_keys[0]} {S['gone'][gone_keys[0]]}")
            else:
                S["notfound"].append({"Menu": menu, "IP": ip_txt, "Hostname": names,
                                      "Alasan": "Hostname tidak ditemukan di file hasil sebelumnya "
                                                "(salah ketik, atau sudah diselesaikan - upload Laporan_Validasi "
                                                "sebelumnya sebagai riwayat)"})
            continue
        new_ip = clean(r["Validasi IP_OAM"])
        if not new_ip:
            for i in cand:
                _mark(S, i, "Validasi IP_OAM masih kosong", "ip_oam", menu, "hostip", "", keys)
            continue
        if not ip_set(new_ip):
            S["warns"].append({"Menu": menu, "Detail": f"{names}: Validasi IP_OAM '{new_ip}' bukan format IP, "
                               "tidak diterapkan"})
            for i in cand:
                _mark(S, i, f"Validasi IP_OAM tidak valid: {new_ip}", "ip_oam", menu, "hostip", "", keys)
            continue
        file_ip = ip_set(r["ip_oam Re-invent"])
        new_set = ip_set(new_ip)
        for i in cand:
            old = df.at[i, "ip_oam"]
            old_set = S["ips"].get(i, set())
            if file_ip and old_set and old_set.isdisjoint(file_ip):
                S["warns"].append({"Menu": menu, "Detail": f"{names}: ip_oam di file sebelumnya ({old}) berbeda "
                                   f"dengan ip_oam Re-invent di file validasi ({clean(r['ip_oam Re-invent'])}); "
                                   "tetap diganti"})
            if old != new_ip:
                df.at[i, "ip_oam"] = new_ip
                S["ips"][i] = new_set
                S["ip_hist"][df.at[i, "_key"]] |= (old_set - new_set)
                S["new_ips"][i] = new_set
                _log(S, menu, df.at[i, "Hostname"], "ip_oam", old, new_ip, "Diganti sesuai Validasi IP_OAM")
            _fill_vm(S, i, [r["VM/Baremetal Re-invent"], r["VM/Baremetal KRI"]], menu)
            _strip(S, i, HIP_FILLS | IP_FILLS)   # IP sudah divalidasi -> konflik IP lama tidak berlaku lagi
            _complete_cleanup(S, i)
            m = S["marks"].get(i)
            if m and m["kind"] == "ip":
                _gugur(S, m["menu"], m["ip"], m["host"], f"GUGUR: IP hostname sudah diganti lewat validasi {menu}")
            _unmark(S, i, HIP_FILLS | IP_FILLS)


def _finalize(S):
    """Sinkronkan kedua menu: hapus tanda yang sudah tidak relevan + cek IP baru yang bentrok."""
    df = S["df"]
    ip_rows = defaultdict(list)
    for i in df.index:
        for ip in S["ips"].get(i, ()):
            ip_rows[ip].append(i)

    for idx, m in list(S["marks"].items()):
        if idx not in df.index:      # baris sudah dihapus menu lain -> tanda gugur
            S["marks"].pop(idx)
            _gugur(S, m["menu"], m["ip"], m["host"], "GUGUR: baris sudah dihapus lewat validasi " + MENU_IP)
            continue
        if m["kind"] == "ip":
            cur = {df.at[j, "_key"] for j in ip_rows.get(m["ip"], []) if df.at[j, "_key"] in m["keys"]}
            if len(cur) < 2:         # konflik IP sudah tidak ada (IP diganti menu lain)
                _unmark(S, idx, IP_FILLS)
                _complete_cleanup(S, idx)
                _gugur(S, m["menu"], m["ip"], m["host"],
                       f"GUGUR: konflik IP sudah tidak ada (IP hostname terkait sudah diganti lewat {MENU_HOST})")

    seen = set()
    for idx, new_set in S["new_ips"].items():
        if idx not in df.index:
            continue
        for ip in new_set:
            for j in ip_rows.get(ip, []):
                if j != idx and df.at[j, "_key"] != df.at[idx, "_key"] and (idx, j) not in seen:
                    seen.add((idx, j)); seen.add((j, idx))
                    S["warns"].append({"Menu": MENU_HOST, "Detail": f"IP baru {ip} untuk {df.at[idx, 'Hostname']} "
                                       f"sama dengan hostname lain: {df.at[j, 'Hostname']} (perlu validasi lagi)"})


def run_validation(prev_df, prev_fills, jobs, first=MENU_IP, history=None):
    """
    jobs  : list of (menu, DataFrame validasi)
    first : menu yang diterapkan lebih dulu (memengaruhi menu berikutnya)
    """
    S = _new_state(prev_df, prev_fills, history)
    for menu, val in sorted(jobs, key=lambda j: 0 if j[0] == first else 1):
        (apply_ip_validation if menu == MENU_IP else apply_ip_oam_validation)(S, val)
    _finalize(S)
    return S


def build_validation_outputs(S, rows_before):
    df = S["df"].reset_index()
    old_idx = df["index"].tolist()
    main = df[REINVENT_COLS].copy()
    main["No"] = list(range(1, len(main) + 1))

    colors, cells, marked = {}, [], []
    for pos, oi in enumerate(old_idx):
        if oi in S["marks"]:
            m = S["marks"][oi]
            colors[pos] = MARK_ROW
            cells.append((pos, m["col"], MARK_CELL))
            marked.append({"Menu": m["menu"], "Hostname": df.at[pos, "Hostname"], "ip_oam": df.at[pos, "ip_oam"],
                           "Alasan": m["reason"]})
        else:
            for col, rgb in S["fills"].get(oi, {}).items():
                cells.append((pos, col, rgb))

    log_df = pd.DataFrame(S["log"], columns=LOG_COLS)
    marked_df = pd.DataFrame(marked, columns=["Menu", "Hostname", "ip_oam", "Alasan"])
    gugur_df = pd.DataFrame(S["gugur"], columns=["Menu", "IP", "Hostname", "Alasan"]).drop_duplicates(
        ignore_index=True)
    nf_df = pd.DataFrame(S["notfound"], columns=["Menu", "IP", "Hostname", "Alasan"])
    warn_df = pd.DataFrame(S["warns"], columns=["Menu", "Detail"]).drop_duplicates(ignore_index=True)
    summary = pd.DataFrame([
        ("Baris file hasil sebelumnya", rows_before),
        ("Baris setelah validasi", len(main)),
        ("Baris dihapus (bentrok IP)", len(S["deleted"])),
        ("Jumlah perubahan nilai", len(log_df)),
        ("Baris DITANDAI (validasi kosong / belum divalidasi)", len(marked_df)),
        ("Entri GUGUR (sudah terselesaikan oleh menu lain)", len(gugur_df)),
        ("Data validasi tidak ditemukan di file sebelumnya", len(nf_df)),
        ("Peringatan", len(warn_df)),
    ], columns=["Keterangan", "Jumlah"])

    main_bytes = to_excel({MAIN_SHEET: main}, {MAIN_SHEET: colors}, {MAIN_SHEET: cells})
    report_bytes = to_excel({"Ringkasan": summary, "Log_Perubahan": log_df, "Ditandai": marked_df,
                             "Gugur_Terselesaikan": gugur_df, "Tidak_Ditemukan": nf_df, "Peringatan": warn_df})
    return dict(main=main, main_bytes=main_bytes, report_bytes=report_bytes, summary=summary,
                log=log_df, marked=marked_df, gugur=gugur_df, notfound=nf_df, warns=warn_df)


VAL_SPECS = {
    MENU_IP: dict(
        cols=VAL_IP_COLS, aliases=VAL_IP_ALIASES, must="Validasi Final Hostname",
        sheet="IP_Sama_Hostname_Beda", key="val_ip",
        label="File validasi 'IP sama, hostname beda' (kolom: IP, Hostname Re-invent, Hostname KRI, "
              "VM/Baremetal Re-invent, VM/Baremetal KRI, Validasi Final Hostname)"),
    MENU_HOST: dict(
        cols=VAL_HOST_COLS, aliases=VAL_HOST_ALIASES, must="Validasi IP_OAM",
        sheet="Hostname_Sama_IP_Beda", key="val_host",
        label="File validasi 'Hostname sama, IP beda' (kolom: Hostname Re-invent, Hostname KRI, ip_oam Re-invent, "
              "IP OAM KRI, VM/Baremetal Re-invent, VM/Baremetal KRI, Keterangan, Validasi IP_OAM)"),
}
WORK_KEY = "val_work"   # kunci st.session_state: hasil langkah yang sedang berjalan


def load_validation(menu, f):
    """Baca + validasi file validasi sebuah menu. Raise ValueError jika kolom wajib tidak ada."""
    sp = VAL_SPECS[menu]
    raw = read_any(f, f"Validasi {menu}", prefer_sheet=sp["sheet"])
    val, miss_v = standardize(raw, sp["cols"], sp["aliases"])
    need = [c for c in (sp["must"], "Hostname Re-invent") if c in miss_v]
    if menu == MENU_IP and "IP" in miss_v:
        need.append("IP")
    if need:
        raise ValueError(f"File validasi '{menu}' wajib punya kolom: {', '.join(need)}")
    return val


def render_validation_results(O, done):
    st.subheader("Hasil setelah: " + "  ➜  ".join(done))
    sm = dict(zip(O["summary"]["Keterangan"], O["summary"]["Jumlah"]))
    c = st.columns(4)
    c[0].metric("Baris sebelum", sm["Baris file hasil sebelumnya"])
    c[1].metric("Baris sesudah", sm["Baris setelah validasi"])
    c[2].metric("Perubahan nilai", sm["Jumlah perubahan nilai"])
    c[3].metric("Baris dihapus", sm["Baris dihapus (bentrok IP)"])
    c = st.columns(4)
    c[0].metric("🟥 Ditandai (validasi kosong)", sm["Baris DITANDAI (validasi kosong / belum divalidasi)"])
    c[1].metric("Entri gugur (terkait menu lain)", sm["Entri GUGUR (sudah terselesaikan oleh menu lain)"])
    c[2].metric("Tidak ditemukan", sm["Data validasi tidak ditemukan di file sebelumnya"])
    c[3].metric("Peringatan", sm["Peringatan"])

    st.markdown("**Warna di file hasil:** 🟥 baris ditandai - Validasi Final Hostname / Validasi IP_OAM masih kosong "
                "atau konflik belum ada di file validasi. Warna lain dari hasil sebelumnya dipertahankan.")

    d1, d2 = st.columns(2)
    d1.download_button("⬇️ Download FILE HASIL VALIDASI (format Re-invent)", O["main_bytes"],
                       "Reinvent_Updated_Validasi.xlsx", XLSX, type="primary")
    d2.download_button("⬇️ Download LAPORAN VALIDASI (log, ditandai, gugur, dll)", O["report_bytes"],
                       "Laporan_Validasi.xlsx", XLSX)

    tabs = st.tabs(["File hasil", "Log perubahan", "Ditandai (kosong)", "Gugur (terselesaikan)",
                    "Tidak ditemukan", "Peringatan"])
    for tab, df in zip(tabs, [O["main"], O["log"], O["marked"], O["gugur"], O["notfound"], O["warns"]]):
        with tab:
            st.write(f"{len(df)} baris")
            st.dataframe(df, hide_index=True)


def validation_page():
    st.title("✅ Terapkan Validasi (Perbaikan)")
    st.caption("Proses satu menu dulu. Setelah selesai, baru upload file validasi menu lainnya - hasil langkah 1 "
               "otomatis dipakai sebagai dasar langkah 2 (kedua menu saling berkorelasi).")

    work = st.session_state.get(WORK_KEY)

    # ---------------- LANGKAH 1 ----------------
    if work is None:
        st.subheader("Langkah 1 - pilih menu dan upload file")
        menu1 = st.radio("Menu yang diproses lebih dulu", [MENU_IP, MENU_HOST], horizontal=True, key="start_menu")
        p1 = st.session_state.get(P1_KEY)
        use_auto = False
        if p1 is not None:
            pick = st.radio(
                "1) File hasil olahan sebelumnya",
                [f"Otomatis dari Halaman 1 ({len(p1['df'])} baris, sudah dikurangi {p1['deleted']} duplikat)",
                 "Upload file sendiri"], key="prev_src")
            use_auto = pick.startswith("Otomatis")
        else:
            st.info("Belum ada hasil Halaman 1 di sesi ini. Proses dulu di Halaman 1 (otomatis terbawa ke sini), "
                    "atau upload file hasil olahan sebelumnya di bawah.")
        f_prev = None
        if not use_auto:
            f_prev = st.file_uploader("1) File hasil olahan sebelumnya (Reinvent_Updated.xlsx)",
                                      type=["xlsx", "xls", "csv"], key="prev_main")
        f_val = st.file_uploader("2) " + VAL_SPECS[menu1]["label"], type=["xlsx", "xls", "csv"],
                                 key=VAL_SPECS[menu1]["key"])
        f_hist = st.file_uploader(
            "3) (Opsional) Laporan_Validasi.xlsx dari validasi sebelumnya", type=["xlsx"], key="hist",
            help="Hanya jika menu lain sudah dikerjakan di sesi/hari sebelumnya, supaya entri yang sudah "
                 "terselesaikan otomatis gugur.")
        go = st.button("▶️ Proses langkah 1", type="primary", key="go1",
                       disabled=((not use_auto and f_prev is None) or f_val is None))
        if not go:
            st.info("Upload file validasi menu terpilih" + ("" if use_auto else " dan file hasil olahan sebelumnya")
                    + ", lalu klik **Proses langkah 1**.")
            return

        if use_auto:
            prev_df, miss = standardize(p1["df"], REINVENT_COLS, REINVENT_ALIASES)
            prev_fills = {k: dict(v) for k, v in p1["fills"].items()}
        else:
            try:
                prev_df, prev_fills, miss = read_prev_main(f_prev)
            except Exception as e:
                st.error(f"Gagal membaca file hasil sebelumnya: {e}")
                return
        if "Hostname" in miss:
            st.error("Kolom Hostname tidak ditemukan di file hasil sebelumnya.")
            return
        if miss:
            st.warning(f"Kolom tidak ditemukan di file sebelumnya (dikosongkan): {', '.join(miss)}")
        history = None
        if f_hist is not None:
            try:
                history = read_history(f_hist)
            except Exception as e:
                st.warning(f"Riwayat tidak terbaca (butuh sheet Log_Perubahan): {e}")
        try:
            val = load_validation(menu1, f_val)
        except Exception as e:
            st.error(str(e))
            return

        S = run_validation(prev_df, prev_fills, [(menu1, val)], first=menu1, history=history)
        st.session_state[WORK_KEY] = dict(S=S, done=[menu1], rows_before=len(prev_df),
                                          O=build_validation_outputs(S, len(prev_df)))
        st.rerun()
        return

    # ---------------- HASIL + LANGKAH 2 ----------------
    done = work["done"]
    st.success("✅ Langkah 1 selesai: " + done[0] + (f"  |  ✅ Langkah 2 selesai: {done[1]}" if len(done) > 1 else ""))
    render_validation_results(work["O"], done)
    st.divider()

    remaining = [m for m in (MENU_IP, MENU_HOST) if m not in done]
    if remaining:
        m2 = remaining[0]
        st.subheader(f"Langkah 2 - {m2}")
        st.caption("Tidak perlu upload ulang file hasil sebelumnya: hasil langkah 1 otomatis dipakai. "
                   "Entri yang sudah terselesaikan di langkah 1 akan gugur otomatis.")
        f2 = st.file_uploader(VAL_SPECS[m2]["label"], type=["xlsx", "xls", "csv"], key=VAL_SPECS[m2]["key"])
        if st.button("▶️ Proses langkah 2", type="primary", key="go2", disabled=(f2 is None)):
            try:
                val2 = load_validation(m2, f2)
            except Exception as e:
                st.error(str(e))
                return
            S = work["S"]
            (apply_ip_validation if m2 == MENU_IP else apply_ip_oam_validation)(S, val2)
            _finalize(S)
            work["done"].append(m2)
            work["O"] = build_validation_outputs(S, work["rows_before"])
            st.session_state[WORK_KEY] = work
            st.rerun()
            return
    else:
        st.success("Kedua menu sudah diproses ✅  Download file hasil di atas.")

    if st.button("🔄 Mulai ulang dari awal", key="reset"):
        st.session_state.pop(WORK_KEY, None)
        st.rerun()


# ----------------------------------------------------------------------------
# MENU 3: BANDINGKAN DENGAN INVENTORY BAREMETAL
#
# Alur: Inventory baremetal (file baru) dicocokkan ke data Re-invent (boleh Reinvent_Updated.xlsx
# hasil halaman 1/2) lewat IP OAM / IP Service. Isi kolom di inventory baremetal lalu DISAMAKAN
# dengan Re-invent (hostname, apps, priority, PIC, environment, site/rack/floor, status Ansible,
# NGPAM, XDR, dll). Nilai dari file baremetal tidak dipakai untuk kolom-kolom tersebut.
# ----------------------------------------------------------------------------
BM_COLS = [
    "No", "Apps", "PIC APPS", "Priority", "Prod/Nonprod", "Brand", "Hardware Type", "EOS Hardware",
    "Serial Number", "Site", "Rack", "Floor", "Hostname MGMT", "Hostname OS", "IP OAM", "IP Services",
    "IP MGMT", "User MGMT", "Password MGMT", "Method access OS", "User OS", "Password OS", "EOS OS",
    "Operating System", "CPU (Core)", "Memory (GB)", "Status", "Scope", "Environment", "Server Type",
    "Zabbix", "Dynatrace", "Ansible", "Notes OS", "Notes MGMT", "Remark Hardware Faulty",
    "CRQ number Dismantle / Poweroff", "Kolom1", "Kolom12", "Telnet & ping",
]
BM_ALIASES = {"IP Service": "IP Services", "PIC Apps": "PIC APPS", "Serial Numbers": "Serial Number"}

# kolom tambahan di sisi kanan (diisi dari Re-invent)
BM_EXTRA = ["VM/Baremetal", "NGPAM Status", "XDR Status", "Lifecycle", "Match By", "Catatan Perbandingan"]

# kolom inventory baremetal -> kolom Re-invent. Nilai Re-invent MENIMPA nilai baremetal (jika Re-invent terisi)
BM_FOLLOW = {
    "Hostname OS": "Hostname", "Apps": "Apps", "PIC APPS": "App PIC Email", "Priority": "Apps Priority",
    "Prod/Nonprod": "Environment", "Site": "Site", "Rack": "rack", "Floor": "floor",
    "IP OAM": "ip_oam", "IP Services": "ip_service", "Ansible": "Ansible Status",
}
# kolom yang hanya diisi dari Re-invent jika di baremetal kosong
BM_FILL = {"Hardware Type": "Hardware Model", "Serial Number": "Serial Numbers", "Operating System": "OS Product"}
# kolom tambahan: selalu ikut Re-invent (termasuk jika kosong) untuk baris yang cocok
BM_EXTRA_FROM_RE = {"VM/Baremetal": "VM/Baremetal", "NGPAM Status": "NGPAM Status",
                    "XDR Status": "XDR Status", "Lifecycle": "Lifecycle"}
BM_STATUS = {"Ansible", "NGPAM Status", "XDR Status"}

BM_ROW = {"multi": "E4DFEC", "nomatch": "FCE4D6", "hostdiff": "FFF2CC", "dismantle": "D9D9D9"}
BM_CHANGED = "C6E0B4"   # sel yang nilainya berubah mengikuti Re-invent
BM_MATCH_LABEL = {0: "IP OAM", 1: "IP OAM baremetal = IP Service Re-invent",
                  2: "IP Service", 3: "IP Service baremetal = IP OAM Re-invent"}


def compare_baremetal(bm_df, re_df):
    """
    bm_df : inventory baremetal (sudah standardize(BM_COLS))
    re_df : Re-invent (sudah standardize(REINVENT_COLS))
    """
    bm = bm_df.copy().reset_index(drop=True)
    re_df = re_df.copy().reset_index(drop=True)
    for col in STATUS_COLS:
        re_df[col] = re_df[col].map(norm_status)
    re_df["_key"] = re_df["Hostname"].map(host_key)

    re_ip = defaultdict(list)          # ip -> [(idx, "oam"/"svc")]
    for i, r in re_df.iterrows():
        for ip in ip_set(r["ip_oam"]):
            re_ip[ip].append((i, "oam"))
        for ip in ip_set(r["ip_service"]):
            re_ip[ip].append((i, "svc"))

    for c in BM_EXTRA:
        bm[c] = ""
    log, changed_cells, row_flags = [], [], {}
    nomatch, multi, hostdiff = [], [], []
    matched_re = set()

    for pos, row in bm.iterrows():
        oam, svc = ip_set(row["IP OAM"]), ip_set(row["IP Services"])
        score = {}   # idx re-invent -> skor terbaik (kecil = lebih kuat)
        for ip in oam:
            for i, kind in re_ip.get(ip, []):
                score[i] = min(score.get(i, 9), 0 if kind == "oam" else 1)
        for ip in svc:
            for i, kind in re_ip.get(ip, []):
                score[i] = min(score.get(i, 9), 2 if kind == "svc" else 3)
        flags = set()
        if not score:
            nomatch.append(pos)
            bm.at[pos, "Catatan Perbandingan"] = "Tidak ditemukan di Re-invent (IP OAM / IP Service)"
            row_flags[pos] = {"nomatch"}
            continue

        best = min(score.values())
        cands = [i for i, s in score.items() if s == best]
        notes = []
        if len(cands) > 1:
            same_host = [i for i in cands if re_df.at[i, "_key"] == host_key(row["Hostname OS"])]
            pick = same_host[0] if same_host else cands[0]
            flags.add("multi")
            notes.append("Cocok ke " + str(len(cands)) + " baris Re-invent (" +
                         ", ".join(re_df.at[i, "Hostname"] or "-" for i in cands) + "); dipakai: " +
                         (re_df.at[pick, "Hostname"] or "-"))
            multi.append({"Hostname OS baremetal": row["Hostname OS"], "IP OAM baremetal": row["IP OAM"],
                          "Kandidat Re-invent": ", ".join(re_df.at[i, "Hostname"] or "-" for i in cands),
                          "Dipakai": re_df.at[pick, "Hostname"]})
        else:
            pick = cands[0]
        matched_re.add(pick)
        rr = re_df.loc[pick]
        bm.at[pos, "Match By"] = BM_MATCH_LABEL[best]

        # --- kolom yang mengikuti Re-invent -----------------------------------
        for bcol, rcol in BM_FOLLOW.items():
            new = rr[rcol]
            if bcol in BM_STATUS:
                new = norm_status(new)
            if not new:        # Re-invent kosong -> nilai lama dibiarkan
                continue
            if new.casefold() != row[bcol].casefold():
                bm.at[pos, bcol] = new
                changed_cells.append((pos, bcol, BM_CHANGED))
                log.append({"Hostname OS": rr["Hostname"], "IP OAM": rr["ip_oam"], "Field": bcol,
                            "Nilai inventory baremetal (lama)": row[bcol], "Nilai baru (Re-invent)": new})
        # --- kolom yang hanya diisi jika kosong -------------------------------
        for bcol, rcol in BM_FILL.items():
            if not row[bcol] and rr[rcol]:
                bm.at[pos, bcol] = rr[rcol]
                changed_cells.append((pos, bcol, BM_CHANGED))
                log.append({"Hostname OS": rr["Hostname"], "IP OAM": rr["ip_oam"], "Field": bcol,
                            "Nilai inventory baremetal (lama)": "", "Nilai baru (Re-invent)": rr[rcol]})
        # --- kolom tambahan ----------------------------------------------------
        for bcol, rcol in BM_EXTRA_FROM_RE.items():
            bm.at[pos, bcol] = rr[rcol]
        for bcol in ("NGPAM Status", "XDR Status"):
            changed_cells.append((pos, bcol, BM_CHANGED))

        # --- catatan ------------------------------------------------------------
        old_host = row["Hostname OS"]
        if old_host and rr["Hostname"] and host_key(old_host) != host_key(rr["Hostname"]):
            flags.add("hostdiff")
            notes.append(f"Hostname OS berbeda: baremetal '{old_host}' -> Re-invent '{rr['Hostname']}'")
            hostdiff.append({"IP OAM": row["IP OAM"], "Hostname OS baremetal": old_host,
                             "Hostname Re-invent": rr["Hostname"], "Match By": BM_MATCH_LABEL[best]})
        if rr["VM/Baremetal"] and "bare" not in rr["VM/Baremetal"].lower():
            notes.append(f"Di Re-invent tercatat {rr['VM/Baremetal']} (bukan Baremetal)")
        if rr["Lifecycle"].lower() == "dismantle":
            flags.add("dismantle")
            notes.append("Lifecycle Re-invent = dismantle")
        if best >= 1:
            notes.append("Match lewat " + BM_MATCH_LABEL[best])
        bm.at[pos, "Catatan Perbandingan"] = "; ".join(notes)
        row_flags[pos] = flags

    bm["No"] = list(range(1, len(bm) + 1))
    colors = {}
    for pos, f in row_flags.items():
        for key in ("multi", "nomatch", "hostdiff", "dismantle"):
            if key in f:
                colors[pos] = BM_ROW[key]
                break

    # Re-invent baremetal yang belum ada di inventory baremetal
    is_bare = re_df["VM/Baremetal"].str.lower().str.contains("bare", na=False)
    re_only = re_df[is_bare & ~re_df.index.isin(matched_re)][REINVENT_COLS].copy()
    nomatch_df = bm.loc[nomatch, BM_COLS + BM_EXTRA].copy()
    log_df = pd.DataFrame(log, columns=["Hostname OS", "IP OAM", "Field", "Nilai inventory baremetal (lama)",
                                        "Nilai baru (Re-invent)"])
    host_df = pd.DataFrame(hostdiff, columns=["IP OAM", "Hostname OS baremetal", "Hostname Re-invent", "Match By"])
    multi_df = pd.DataFrame(multi, columns=["Hostname OS baremetal", "IP OAM baremetal", "Kandidat Re-invent", "Dipakai"])
    summary = pd.DataFrame([
        ("Total baris inventory baremetal", len(bm)),
        ("Total baris Re-invent", len(re_df)),
        ("Baremetal yang cocok dengan Re-invent", len(bm) - len(nomatch)),
        ("Baremetal TIDAK ada di Re-invent", len(nomatch)),
        ("Re-invent (Baremetal) TIDAK ada di inventory baremetal", len(re_only)),
        ("Hostname OS berbeda (ditimpa Re-invent)", len(host_df)),
        ("Cocok ke lebih dari 1 baris Re-invent", len(multi_df)),
        ("Jumlah perubahan nilai", len(log_df)),
    ], columns=["Keterangan", "Jumlah"])
    return dict(main=bm[BM_COLS + BM_EXTRA], colors=colors, cells=changed_cells, nomatch=nomatch_df,
                re_only=re_only, hostdiff=host_df, multi=multi_df, log=log_df, summary=summary)


def build_baremetal_outputs(R):
    main = to_excel({"Baremetal_Updated": R["main"]}, {"Baremetal_Updated": R["colors"]},
                    {"Baremetal_Updated": R["cells"]})
    report = to_excel({
        "Ringkasan": R["summary"],
        "Baremetal_Tidak_Ada_di_Reinvent": R["nomatch"],
        "Reinvent_Tidak_Ada_di_Baremetal": R["re_only"],
        "Hostname_Beda": R["hostdiff"],
        "Multi_Match": R["multi"],
        "Log_Perubahan": R["log"],
    })
    return main, report


def sample_baremetal():
    def r(no, host, oam, svc="", apps="Lama", **kw):
        d = {c: "" for c in BM_COLS}
        d.update({"No": no, "Hostname OS": host, "IP OAM": oam, "IP Services": svc, "Apps": apps,
                  "Priority": "P3", "Site": "OLD-SITE"})
        d.update(kw)
        return d
    rows = [r(1, "srv-app01", "10.10.1.1"), r(2, "srv-db01-old", "10.10.1.2"),
            r(3, "srv-lonely", "10.10.77.7"), r(4, "srv-svc", "", "10.20.0.5")]
    return pd.DataFrame(rows, columns=BM_COLS)


def baremetal_page():
    st.title("🖥️ Bandingkan dengan Inventory Baremetal")
    st.caption("Inventory baremetal dicocokkan ke Re-invent lewat IP OAM / IP Service. Hostname, Apps, Priority, "
               "PIC, Environment, Site/Rack/Floor, status Ansible/NGPAM/XDR, dll mengikuti data Re-invent.")
    use_dummy = st.sidebar.checkbox("Pakai data contoh (dummy)", value=False, key="bm_dummy")
    if use_dummy:
        f_bm = sample_baremetal()
        f_re, _ = sample_data()
        f_re = f_re.copy()
        f_re["ip_service"] = ["", "", "", "", "", "", ""]
        f_re.loc[2, "ip_service"] = "10.20.0.5"
        f_re["VM/Baremetal"] = "Baremetal"
        f_re["Hardware Model"] = "Dell R740"
        f_re["App PIC Email"] = "pic@telkomsel.co.id"
    else:
        f_bm = st.file_uploader("1) Inventory baremetal (file baru)", type=["xlsx", "xls", "csv"], key="bm_file")
        f_re = st.file_uploader("2) Data Re-invent (boleh Reinvent_Updated.xlsx hasil halaman 1/2)",
                                type=["xlsx", "xls", "csv"], key="bm_re")
    if f_bm is None or f_re is None:
        st.info("Upload inventory baremetal dan data Re-invent, atau centang *Pakai data contoh* di sidebar.")
        return
    try:
        bm_df, miss_bm = standardize(read_any(f_bm, "Baremetal"), BM_COLS, BM_ALIASES)
        re_df, miss_re = standardize(read_any(f_re, "Re-invent (baremetal)"), REINVENT_COLS, REINVENT_ALIASES)
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        return
    if miss_bm:
        st.warning(f"Kolom tidak ditemukan di inventory baremetal (dikosongkan): {', '.join(miss_bm)}")
    if miss_re:
        st.warning(f"Kolom tidak ditemukan di Re-invent (dikosongkan): {', '.join(miss_re)}")
    if all(c in miss_bm for c in ("IP OAM", "IP Services")):
        st.error("Inventory baremetal harus punya kolom IP OAM atau IP Services.")
        return

    R = compare_baremetal(bm_df, re_df)
    main_bytes, report_bytes = build_baremetal_outputs(R)
    s = dict(zip(R["summary"]["Keterangan"], R["summary"]["Jumlah"]))
    c = st.columns(4)
    c[0].metric("Baremetal", s["Total baris inventory baremetal"])
    c[1].metric("Cocok dgn Re-invent", s["Baremetal yang cocok dengan Re-invent"])
    c[2].metric("Tidak ada di Re-invent", s["Baremetal TIDAK ada di Re-invent"])
    c[3].metric("Re-invent tdk ada di baremetal", s["Re-invent (Baremetal) TIDAK ada di inventory baremetal"])
    c = st.columns(3)
    c[0].metric("Hostname beda", s["Hostname OS berbeda (ditimpa Re-invent)"])
    c[1].metric("Multi match", s["Cocok ke lebih dari 1 baris Re-invent"])
    c[2].metric("Perubahan nilai", s["Jumlah perubahan nilai"])
    st.markdown("**Warna:** 🟩 sel diubah mengikuti Re-invent · 🟨 hostname berbeda · 🟪 cocok ke >1 baris Re-invent · "
                "🟧 tidak ada di Re-invent · ⬜ dismantle")
    d1, d2 = st.columns(2)
    d1.download_button("⬇️ Download INVENTORY BAREMETAL (sudah disesuaikan)", main_bytes, "Baremetal_Updated.xlsx",
                       XLSX, type="primary")
    d2.download_button("⬇️ Download LAPORAN PERBANDINGAN", report_bytes, "Laporan_Baremetal.xlsx", XLSX)
    tabs = st.tabs(["Inventory hasil", "Baremetal tidak ada di Re-invent", "Re-invent tidak ada di baremetal",
                    "Hostname beda", "Multi match", "Log perubahan"])
    for tab, df in zip(tabs, [R["main"], R["nomatch"], R["re_only"], R["hostdiff"], R["multi"], R["log"]]):
        with tab:
            st.write(f"{len(df)} baris")
            st.dataframe(df, hide_index=True)


def dup_selection_ui(R):
    """Tampilkan baris duplikat dengan kotak centang 'Hapus'. Return: set posisi baris yang dihapus."""
    dups = dup_rows_df(R["M"])
    if dups.empty:
        return set()
    st.subheader("🗑️ Duplikat hostname - pilih yang dihapus")
    st.caption(f"{len(dups)} baris dalam {dups['Grup'].nunique()} grup hostname duplikat. Centang **Hapus** pada baris "
               "yang mau dibuang; file Excel di bawah baru dibuat setelah baris yang dicentang dikeluarkan.")
    modes = ["Manual (belum ada yang dicentang)", "Hapus semua kecuali baris pertama tiap grup",
             "Hapus semua kecuali baris terakhir tiap grup"]
    mode = st.radio("Centang awal", modes, horizontal=True, key="dup_mode")
    view = dups.copy()
    if mode == modes[1]:
        init = view["Grup"].duplicated(keep="first")
    elif mode == modes[2]:
        init = view["Grup"].duplicated(keep="last")
    else:
        init = pd.Series(False, index=view.index)
    view.insert(0, "Hapus", init.values)
    sig = hashlib.md5("|".join(f"{i}:{h}" for i, h in zip(view.index, view["Hostname"])).encode()).hexdigest()[:8]
    edited = st.data_editor(
        view, hide_index=True, key=f"dup_ed_{sig}_{modes.index(mode)}",
        column_config={"Hapus": st.column_config.CheckboxColumn("Hapus", help="Centang = baris dibuang dari file utama")},
        disabled=[c for c in view.columns if c != "Hapus"],
    )
    deleted = set(edited.index[edited["Hapus"].astype(bool)])
    emptied = [g for g, d in edited.groupby("Grup") if d["Hapus"].all()]
    if emptied:
        st.warning("Semua baris untuk hostname ini ikut dihapus (hostname hilang dari file): " + ", ".join(emptied[:10])
                   + (" ..." if len(emptied) > 10 else ""))
    st.info(f"{len(deleted)} baris dipilih untuk dihapus → file utama menjadi "
            f"{len(R['M']['base']) - len(deleted)} baris.")
    return deleted


def main():
    st.set_page_config(page_title="Re-invent vs KRI", page_icon="🔍", layout="wide")

    page = st.sidebar.radio("Halaman", ["1. Bandingkan data", "2. Terapkan validasi (perbaikan)",
                                        "3. Bandingkan dengan inventory baremetal"])
    if page.startswith("2"):
        validation_page()
        return
    if page.startswith("3"):
        baremetal_page()
        return

    with st.sidebar:
        st.header("1. Sumber data")
        modes = ["Upload file"]
        if find_file(DATA_DIR, "reinvent") and find_file(DATA_DIR, "kri"):
            modes.append("File di folder data/ (repo)")
        modes.append("Data contoh (dummy)")
        mode = st.radio("Ambil data dari", modes)

        if mode == "Upload file":
            f_re = st.file_uploader("Data Re-invent", type=["xlsx", "xls", "csv"])
            f_kri = st.file_uploader("Data KRI", type=["xlsx", "xls", "csv"])
        elif mode.startswith("File di folder"):
            f_re, f_kri = find_file(DATA_DIR, "reinvent"), find_file(DATA_DIR, "kri")
            st.caption(f"Re-invent: `{f_re.name}`  \nKRI: `{f_kri.name}`")
        else:
            f_re, f_kri = sample_data()
            st.caption("Memakai data dummy bawaan aplikasi.")

        st.header("2. Aturan status terupdate")
        strategy_label = st.radio(
            "Jika nilai Ansible / NGPAM / XDR berbeda, pakai:",
            ["Otomatis (bandingkan tanggal)", "Selalu KRI", "Selalu Re-invent"],
            help="Re-invent tidak punya kolom tanggal, jadi tanggal update data Re-invent diisi manual di bawah. "
                 "Jika nilai di salah satu sisi kosong, nilai dari sisi lain otomatis dipakai.",
        )
        strategy = {"Otomatis (bandingkan tanggal)": "auto", "Selalu KRI": "KRI",
                    "Selalu Re-invent": "Re-invent"}[strategy_label]
        use_date = st.checkbox("Tentukan tanggal update data Re-invent", value=True) if strategy == "auto" else False
        reinvent_date = st.date_input("Tanggal data Re-invent terakhir diupdate", value=date.today()) if use_date else None
        sync = st.multiselect(
            "Kolom yang disamakan dengan data terupdate", SYNC_CANDIDATES, default=DEFAULT_SYNC,
            help="Kolom lain hanya diisi dari KRI jika di Re-invent kosong.",
        )

        st.header("3. Dismantle")
        dismantle_kw = st.text_input("Keyword dismantle (pisahkan koma)", DEFAULT_OPTIONS["dismantle_kw"])
        remark_cols = st.multiselect(
            "Kolom KRI yang dicek",
            ["Remark From OSPLA Team", "Remark From NGPAM Team", "Timeline Integration OSPLA", "OS EOL"],
            default=DEFAULT_REMARK_COLS,
        )

        st.header("4. Output")
        add_note = st.checkbox(
            "Tambah kolom 'Catatan' di file utama", value=False,
            help="Default mati agar kolom file utama sama persis dengan Re-invent. Baris ditandai dengan warna.",
        )
        split_incomplete = st.checkbox("Pisahkan baris tanpa Hostname/VM-Baremetal dari file utama", value=False)

    st.title("🔍 Perbandingan Data Re-invent vs KRI")
    st.caption("Tujuan: semua yang ada di KRI harus ada di Re-invent. File utama memakai kolom persis seperti Re-invent.")

    if not (f_re is not None and f_kri is not None):
        st.info("Silakan upload kedua file (Re-invent & KRI) di sidebar, atau pilih *Data contoh* untuk mencoba.")
        st.stop()

    try:
        raw_re = read_any(f_re, "Re-invent")
        raw_kri = read_any(f_kri, "KRI")
    except Exception as e:  # file rusak / format tidak didukung
        st.error(f"Gagal membaca file: {e}")
        st.stop()

    re_df, miss_re = standardize(raw_re, REINVENT_COLS, REINVENT_ALIASES)
    kri_df, miss_kri = standardize(raw_kri, KRI_COLS, KRI_ALIASES)

    if miss_re:
        st.warning(f"Kolom tidak ditemukan di Re-invent (dikosongkan): {', '.join(miss_re)}")
    if miss_kri:
        st.warning(f"Kolom tidak ditemukan di KRI (dikosongkan): {', '.join(miss_kri)}")
    for nm, miss in (("Re-invent", miss_re), ("KRI", miss_kri)):
        if "Hostname" in miss:
            st.error(f"Kolom Hostname wajib ada di file {nm}.")
            st.stop()

    opts = dict(strategy=strategy, reinvent_date=reinvent_date, sync=sync, dismantle_kw=dismantle_kw,
                remark_cols=remark_cols, add_note=add_note, split_incomplete=split_incomplete)
    R = compute(re_df, kri_df, opts)

    s = dict(zip(R["summary"]["Keterangan"], R["summary"]["Jumlah"]))
    c = st.columns(5)
    c[0].metric("Re-invent", s["Total baris Re-invent"])
    c[1].metric("KRI", s["Total baris KRI"])
    c[2].metric("Di KRI saja", s["Ada di KRI, tidak ada di Re-invent"])
    c[3].metric("Di Re-invent saja", s["Ada di Re-invent, tidak ada di KRI"])
    c[4].metric("Dismantle", s["Lifecycle = dismantle"])
    c = st.columns(4)
    c[0].metric("Hostname beda, IP sama", s["Hostname beda, IP sama"])
    c[1].metric("Hostname sama, IP beda", s["Hostname sama, IP beda"])
    c[2].metric("Hostname/VM kosong", s["Hostname / VM-Baremetal kosong"])
    c[3].metric("Perubahan nilai", s["Jumlah perubahan nilai"])

    st.markdown("**Warna di file utama:** 🟨 hostname beda IP sama · 🟦 hostname sama IP beda · "
                "🟪 hostname duplikat · 🟧 Hostname/VM-Baremetal kosong · ⬜ dismantle · 🟩 baris baru dari KRI  \n"
                "<small>Seluruh baris diwarnai sesuai masalah utama, sel yang bermasalah (ip_oam / Hostname / "
                "VM-Baremetal / Lifecycle) diwarnai lebih pekat.</small>", unsafe_allow_html=True)

    deleted = dup_selection_ui(R)
    main_bytes, report_bytes = build_outputs(R, deleted)

    # simpan hasil (setelah hapus duplikat) -> Halaman 2 otomatis memakainya, tidak perlu upload ulang
    final_main, f_colors, f_cells = finalize_main(R["M"], deleted)
    fills = {pos: ({c: f_colors[pos] for c in REINVENT_COLS} if pos in f_colors else {})
             for pos in range(len(final_main))}
    for pos, col, color in f_cells:
        fills[pos][col] = color
    st.session_state[P1_KEY] = dict(df=final_main[REINVENT_COLS].copy(), fills=fills, deleted=len(deleted))
    st.success(f"✅ Hasil ini ({len(final_main)} baris) otomatis dipakai di Halaman 2 - tidak perlu upload ulang.")

    d1, d2 = st.columns(2)
    d1.download_button("⬇️ Download FILE UTAMA (format Re-invent)", main_bytes, "Reinvent_Updated.xlsx",
                       XLSX, type="primary")
    d2.download_button("⬇️ Download FILE LAPORAN (IP sama, dismantle, log, dll)", report_bytes,
                       "Laporan_Perbandingan.xlsx", XLSX)

    with st.expander("🔎 Cek kunci pencocokan hostname (hostname1.domain apa saja → hostname1)"):
        chk = pd.concat([
            pd.DataFrame({"Sumber": "Re-invent", "Hostname asli": re_df["Hostname"], "Kunci": re_df["Hostname"].map(host_key)}),
            pd.DataFrame({"Sumber": "KRI", "Hostname asli": kri_df["Hostname"], "Kunci": kri_df["Hostname"].map(host_key)}),
        ], ignore_index=True)
        st.dataframe(chk, hide_index=True)

    tabs = st.tabs(["File utama", "KRI tidak ada di Re-invent", "Re-invent tidak ada di KRI",
                    "IP sama, hostname beda", "Hostname sama, IP beda", "Dismantle",
                    "Perlu dilengkapi", "Duplikat", "Log perubahan", "Hostname bertanda /"])
    frames = [R["main"], R["kri_only"], R["re_only"], R["conf"], R["hip"], R["dismantle"],
              R["incomplete"], R["dup"], R["log"], R["slash"]]
    hints = {3: "Isi kolom **Validasi Final Hostname** di sheet `IP_Sama_Hostname_Beda` (file laporan), "
                "lalu upload di halaman 2 (Terapkan validasi).",
             4: "Isi kolom **Validasi IP_OAM** di sheet `Hostname_Sama_IP_Beda` (file laporan), "
                "lalu upload di halaman 2 (Terapkan validasi).",
             9: "Hostname berisi dua nilai (xxx / xxxx): bagian yang cocok dengan file lawan dipertahankan, "
                "bagian lain dihapus. Status 'Belum terselesaikan' / 'Perlu dicek' sebaiknya dicek manual."}
    for n, (tab, df) in enumerate(zip(tabs, frames)):
        with tab:
            if n in hints:
                st.caption(hints[n])
            st.write(f"{len(df)} baris")
            st.dataframe(df, hide_index=True)


if __name__ == "__main__":
    main()
