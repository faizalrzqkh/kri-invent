"""
Streamlit app (1 file): Perbandingan data Re-invent vs KRI
Jalankan:  streamlit run app.py

Isi file:  1) definisi kolom  2) helper  3) logika perbandingan  4) export Excel
           5) data contoh (dummy)  6) tampilan Streamlit
"""
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
        df[c] = df[c].map(clean)
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

    log = []
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
    def label(idx):
        out = []
        if "ip" in flags[idx]:
            out.append("Hostname beda/IP sama")
        if "hip" in flags[idx]:
            out.append("Hostname sama/IP beda")
        if "dup" in flags[idx]:
            out.append("Hostname duplikat")
        if "inc" in flags[idx]:
            out.append("Hostname/VM-Baremetal kosong")
        if res.at[idx, "Lifecycle"] == "dismantle":
            out.append("Dismantle")
        if "new" in flags[idx]:
            out.append("Baru dari KRI")
        return "; ".join(out)

    main_res = res[~res.index.isin(excluded)]
    main = main_res[REINVENT_COLS].copy().reset_index(drop=True)
    main["No"] = list(range(1, len(main) + 1))
    if o["add_note"]:
        main["Catatan"] = [label(i) for i in main_res.index]

    colors, cells = {}, []
    for pos, idx in enumerate(main_res.index):
        f = flags[idx]
        is_dis = res.at[idx, "Lifecycle"] == "dismantle"
        for key in ("ip", "hip", "dup", "inc"):          # urutan = prioritas warna baris
            if key in f:
                colors[pos] = COLORS[key]
                break
        else:
            if is_dis:
                colors[pos] = COLORS["dismantle"]
            elif "new" in f:
                colors[pos] = COLORS["new"]
        # warna sel yang bermasalah
        if "ip" in f:
            cells.append((pos, "ip_oam", CELL_COLORS["ip"]))
        if "hip" in f:
            cells.append((pos, "ip_oam", CELL_COLORS["hip"]))
        if "dup" in f:
            cells.append((pos, "Hostname", CELL_COLORS["dup"]))
        if "inc" in f:
            if not res.at[idx, "Hostname"]:
                cells.append((pos, "Hostname", CELL_COLORS["inc"]))
            if not res.at[idx, "VM/Baremetal"]:
                cells.append((pos, "VM/Baremetal", CELL_COLORS["inc"]))
        if is_dis:
            cells.append((pos, "Lifecycle", CELL_COLORS["dismantle"]))

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
        ("Jumlah perubahan nilai", len(log_df)),
        ("Total baris file utama", len(main)),
    ], columns=["Keterangan", "Jumlah"])

    return dict(main=main, colors=colors, cells=cells, kri_only=kri_only_df, re_only=re_only_df, conf=conf_df,
                hip=hip_df, dismantle=dismantle_df, incomplete=incomplete_df, dup=dup_report,
                log=log_df, summary=summary)


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


def build_outputs(result):
    """Bangun 2 file Excel (bytes): file utama & file laporan."""
    main_bytes = to_excel({MAIN_SHEET: result["main"]}, {MAIN_SHEET: result["colors"]},
                          {MAIN_SHEET: result["cells"]})
    report_sheets = {
        "Ringkasan": result["summary"],
        "KRI_Tidak_Ada_di_Reinvent": result["kri_only"],
        "Reinvent_Tidak_Ada_di_KRI": result["re_only"],
        "IP_Sama_Hostname_Beda": result["conf"],
        "Hostname_Sama_IP_Beda": result["hip"],
        "Dismantle": result["dismantle"],
        "Perlu_Dilengkapi": result["incomplete"],
        "Duplikat": result["dup"],
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
    ]
    return pd.DataFrame(reinvent, columns=REINVENT_COLS), pd.DataFrame(kri, columns=KRI_COLS)


# ----------------------------------------------------------------------------
# TAMPILAN STREAMLIT
# ----------------------------------------------------------------------------
DATA_DIR = Path(__file__).parent / "data"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


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


def _new_state(df, fills):
    df = df.copy()
    df["_key"] = df["Hostname"].map(host_key)
    df["_ips"] = df["ip_oam"].map(ip_set)
    return dict(df=df, fills=dict(fills), marks={}, marked=[], deleted=set(),
                log=[], notfound=[], warns=[])


def _log(S, menu, host, field, old, new, ket=""):
    S["log"].append({"Menu": menu, "Hostname": host, "Field": field,
                     "Nilai lama": old, "Nilai baru": new, "Keterangan": ket})


def _mark(S, idx, reason, col, menu):
    if idx in S["marks"]:
        return
    S["marks"][idx] = (reason, col)
    S["fills"].pop(idx, None)
    S["marked"].append({"Menu": menu, "Hostname": S["df"].at[idx, "Hostname"],
                        "ip_oam": S["df"].at[idx, "ip_oam"], "Alasan": reason})


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
    Untuk tiap IP: baris-baris yang bentrok di file sebelumnya dihapus lalu diganti SATU baris
    dengan Hostname = 'Validasi Final Hostname'. Jika validasi kosong -> baris ditandai.
    """
    menu = MENU_IP
    df = S["df"]
    flagged = {i for i, f in S["fills"].items() if set(f.values()) & IP_FILLS}  # bentrok dari hasil lama

    ip_idx = defaultdict(list)
    for i, ips in zip(df.index, df["_ips"]):
        for ip in ips:
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
        names = " | ".join(sorted({clean(r[c]) for r in rows for c in ("Hostname Re-invent", "Hostname KRI") if clean(r[c])}))
        cand = [i for i in ip_idx.get(ip, []) if i not in S["deleted"] and df.at[i, "_key"] in keys]
        if not cand:
            S["notfound"].append({"Menu": menu, "IP": ip, "Hostname": names,
                                  "Alasan": "IP + hostname tidak ditemukan di file hasil sebelumnya"})
            continue

        finals = [clean(r["Validasi Final Hostname"]) for r in rows if clean(r["Validasi Final Hostname"])]
        if not finals:
            for i in cand:
                _mark(S, i, "Validasi Final Hostname masih kosong", "Hostname", menu)
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

        old_host = df.at[survivor, "Hostname"]
        if old_host != final:
            df.at[survivor, "Hostname"] = final
            _log(S, menu, final, "Hostname", old_host, final, f"IP {ip}: diganti sesuai Validasi Final Hostname")
        df.at[survivor, "_key"] = fk
        _fill_vm(S, survivor, vm_options, menu)
        for o in others:
            _log(S, menu, df.at[o, "Hostname"], "(baris dihapus)", df.at[o, "Hostname"], "",
                 f"IP {ip}: baris bentrok dihapus, diganti hostname final {final}")
            S["deleted"].add(o)
            S["fills"].pop(o, None)
        _strip(S, survivor, IP_FILLS)
        _complete_cleanup(S, survivor)

        if ((df["_key"] == fk) & ~df.index.isin(cand) & ~df.index.isin(list(S["deleted"]))).any():
            S["warns"].append({"Menu": menu, "Detail": f"IP {ip}: hostname final '{final}' ternyata sudah ada di "
                               "baris lain pada file sebelumnya (kemungkinan duplikat)"})

    S["df"] = df.drop(index=list(S["deleted"]))
    df = S["df"]

    # bandingkan dengan data sebelumnya: konflik IP yang masih tersisa tapi tidak ada di file validasi
    ip_rows = defaultdict(list)
    for i, ips in zip(df.index, df["_ips"]):
        for ip in ips:
            ip_rows[ip].append(i)
    for ip, idxs in ip_rows.items():
        if ip in validated_ips:
            continue
        use = [i for i in idxs if i in flagged] if flagged else idxs
        if len({df.at[i, "_key"] for i in use}) > 1:
            for i in use:
                _mark(S, i, f"IP {ip} sama, hostname beda - belum ada di file validasi", "Hostname", menu)


def apply_ip_oam_validation(S, val):
    """
    Hostname sama, IP beda.
    ip_oam di file sebelumnya diganti dengan 'Validasi IP_OAM'. Jika validasi kosong -> baris ditandai.
    """
    menu = MENU_HOST
    df = S["df"]
    key_idx = defaultdict(list)
    for i, k in zip(df.index, df["_key"]):
        if k:
            key_idx[k].append(i)

    for _, r in val.iterrows():
        keys = {host_key(r["Hostname Re-invent"]), host_key(r["Hostname KRI"])} - {""}
        if not keys:
            continue
        names = " | ".join(sorted({clean(r[c]) for c in ("Hostname Re-invent", "Hostname KRI") if clean(r[c])}))
        cand = [i for k in keys for i in key_idx.get(k, []) if i in df.index]
        if not cand:
            S["notfound"].append({"Menu": menu, "IP": clean(r["ip_oam Re-invent"]), "Hostname": names,
                                  "Alasan": "Hostname tidak ditemukan di file hasil sebelumnya"})
            continue
        new_ip = clean(r["Validasi IP_OAM"])
        if not new_ip:
            for i in cand:
                _mark(S, i, "Validasi IP_OAM masih kosong", "ip_oam", menu)
            continue
        if not ip_set(new_ip):
            S["warns"].append({"Menu": menu, "Detail": f"{names}: Validasi IP_OAM '{new_ip}' bukan format IP, tidak diterapkan"})
            for i in cand:
                _mark(S, i, f"Validasi IP_OAM tidak valid: {new_ip}", "ip_oam", menu)
            continue
        file_ip = ip_set(r["ip_oam Re-invent"])
        for i in cand:
            old = df.at[i, "ip_oam"]
            if file_ip and ip_set(old) and ip_set(old).isdisjoint(file_ip):
                S["warns"].append({"Menu": menu, "Detail": f"{names}: ip_oam di file sebelumnya ({old}) berbeda dengan "
                                   f"ip_oam Re-invent di file validasi ({clean(r['ip_oam Re-invent'])}); tetap diganti"})
            if old != new_ip:
                df.at[i, "ip_oam"] = new_ip
                _log(S, menu, df.at[i, "Hostname"], "ip_oam", old, new_ip, "Diganti sesuai Validasi IP_OAM")
            _fill_vm(S, i, [r["VM/Baremetal Re-invent"], r["VM/Baremetal KRI"]], menu)
            _strip(S, i, HIP_FILLS)
            _complete_cleanup(S, i)
            S["marks"].pop(i, None)


def run_validation(prev_df, prev_fills, jobs):
    """jobs: list of (menu, DataFrame validasi). Urutan: IP sama dulu, lalu hostname sama."""
    S = _new_state(prev_df, prev_fills)
    for menu, val in sorted(jobs, key=lambda j: 0 if j[0] == MENU_IP else 1):
        (apply_ip_validation if menu == MENU_IP else apply_ip_oam_validation)(S, val)
    return S


def build_validation_outputs(S, rows_before):
    df = S["df"].reset_index()
    old_idx = df["index"].tolist()
    main = df[REINVENT_COLS].copy()
    main["No"] = list(range(1, len(main) + 1))

    colors, cells = {}, []
    for pos, oi in enumerate(old_idx):
        if oi in S["marks"]:
            colors[pos] = MARK_ROW
            cells.append((pos, S["marks"][oi][1], MARK_CELL))
        else:
            for col, rgb in S["fills"].get(oi, {}).items():
                cells.append((pos, col, rgb))

    log_df = pd.DataFrame(S["log"], columns=["Menu", "Hostname", "Field", "Nilai lama", "Nilai baru", "Keterangan"])
    marked_df = pd.DataFrame(S["marked"], columns=["Menu", "Hostname", "ip_oam", "Alasan"])
    nf_df = pd.DataFrame(S["notfound"], columns=["Menu", "IP", "Hostname", "Alasan"])
    warn_df = pd.DataFrame(S["warns"], columns=["Menu", "Detail"])
    summary = pd.DataFrame([
        ("Baris file hasil sebelumnya", rows_before),
        ("Baris setelah validasi", len(main)),
        ("Baris dihapus (bentrok IP)", len(S["deleted"])),
        ("Jumlah perubahan nilai", len(log_df)),
        ("Baris DITANDAI (validasi kosong / belum divalidasi)", len(marked_df)),
        ("Data validasi tidak ditemukan di file sebelumnya", len(nf_df)),
        ("Peringatan", len(warn_df)),
    ], columns=["Keterangan", "Jumlah"])

    main_bytes = to_excel({MAIN_SHEET: main}, {MAIN_SHEET: colors}, {MAIN_SHEET: cells})
    report_bytes = to_excel({"Ringkasan": summary, "Log_Perubahan": log_df, "Ditandai": marked_df,
                             "Tidak_Ditemukan": nf_df, "Peringatan": warn_df})
    return dict(main=main, main_bytes=main_bytes, report_bytes=report_bytes, summary=summary,
                log=log_df, marked=marked_df, notfound=nf_df, warns=warn_df)


def validation_page():
    st.title("✅ Terapkan Validasi (Perbaikan)")
    st.caption("Upload file hasil olahan sebelumnya + file Excel validasi yang sudah Anda isi. "
               "Aplikasi menerapkan perbaikan dan menandai baris yang validasinya masih kosong.")

    with st.sidebar:
        st.header("Menu yang diperbaiki")
        menus = st.multiselect("Pilih menu", [MENU_IP, MENU_HOST], default=[MENU_IP])
        st.header("Upload file")
        f_prev = st.file_uploader("File hasil olahan sebelumnya (Reinvent_Updated.xlsx)",
                                  type=["xlsx", "xls", "csv"], key="prev_main")
        f_val = {}
        if MENU_IP in menus:
            f_val[MENU_IP] = st.file_uploader(
                "Validasi 'IP sama, hostname beda' (kolom: IP, Hostname Re-invent, Hostname KRI, "
                "VM/Baremetal Re-invent, VM/Baremetal KRI, Validasi Final Hostname)",
                type=["xlsx", "xls", "csv"], key="val_ip")
        if MENU_HOST in menus:
            f_val[MENU_HOST] = st.file_uploader(
                "Validasi 'Hostname sama, IP beda' (kolom: Hostname Re-invent, Hostname KRI, ip_oam Re-invent, "
                "IP OAM KRI, VM/Baremetal Re-invent, VM/Baremetal KRI, Keterangan, Validasi IP_OAM)",
                type=["xlsx", "xls", "csv"], key="val_host")

    if not menus:
        st.info("Pilih minimal satu menu di sidebar.")
        st.stop()
    if f_prev is None or any(f is None for f in f_val.values()):
        st.info("Upload file hasil olahan sebelumnya dan file validasi untuk setiap menu yang dipilih.")
        st.stop()

    try:
        prev_df, prev_fills, miss = read_prev_main(f_prev)
    except Exception as e:
        st.error(f"Gagal membaca file hasil sebelumnya: {e}")
        st.stop()
    if "Hostname" in miss:
        st.error("Kolom Hostname tidak ditemukan di file hasil sebelumnya.")
        st.stop()
    if miss:
        st.warning(f"Kolom tidak ditemukan di file sebelumnya (dikosongkan): {', '.join(miss)}")

    specs = {MENU_IP: (VAL_IP_COLS, VAL_IP_ALIASES, "Validasi Final Hostname", "sheet IP_Sama_Hostname_Beda"),
             MENU_HOST: (VAL_HOST_COLS, VAL_HOST_ALIASES, "Validasi IP_OAM", "sheet Hostname_Sama_IP_Beda")}
    sheet_pref = {MENU_IP: "IP_Sama_Hostname_Beda", MENU_HOST: "Hostname_Sama_IP_Beda"}
    jobs = []
    for menu, f in f_val.items():
        cols, aliases, must, _ = specs[menu]
        try:
            raw = read_any(f, f"Validasi {menu}", prefer_sheet=sheet_pref[menu])
        except Exception as e:
            st.error(f"Gagal membaca file validasi '{menu}': {e}")
            st.stop()
        val, miss_v = standardize(raw, cols, aliases)
        need = [c for c in (must, "Hostname Re-invent") if c in miss_v]
        if "IP" in miss_v and menu == MENU_IP:
            need.append("IP")
        if need:
            st.error(f"File validasi '{menu}' wajib punya kolom: {', '.join(need)}")
            st.stop()
        jobs.append((menu, val))

    S = run_validation(prev_df, prev_fills, jobs)
    O = build_validation_outputs(S, len(prev_df))

    sm = dict(zip(O["summary"]["Keterangan"], O["summary"]["Jumlah"]))
    c = st.columns(4)
    c[0].metric("Baris sebelum", sm["Baris file hasil sebelumnya"])
    c[1].metric("Baris sesudah", sm["Baris setelah validasi"])
    c[2].metric("Perubahan nilai", sm["Jumlah perubahan nilai"])
    c[3].metric("Baris dihapus", sm["Baris dihapus (bentrok IP)"])
    c = st.columns(3)
    c[0].metric("🟥 Ditandai (validasi kosong)", sm["Baris DITANDAI (validasi kosong / belum divalidasi)"])
    c[1].metric("Tidak ditemukan", sm["Data validasi tidak ditemukan di file sebelumnya"])
    c[2].metric("Peringatan", sm["Peringatan"])

    st.markdown("**Warna di file hasil:** 🟥 baris ditandai - Validasi Final Hostname / Validasi IP_OAM masih kosong "
                "atau konflik belum ada di file validasi. Warna lain dari hasil sebelumnya dipertahankan.")

    d1, d2 = st.columns(2)
    d1.download_button("⬇️ Download FILE HASIL VALIDASI (format Re-invent)", O["main_bytes"],
                       "Reinvent_Updated_Validasi.xlsx", XLSX, type="primary")
    d2.download_button("⬇️ Download LAPORAN VALIDASI (log, ditandai, tidak ditemukan)", O["report_bytes"],
                       "Laporan_Validasi.xlsx", XLSX)

    tabs = st.tabs(["File hasil", "Log perubahan", "Ditandai (kosong)", "Tidak ditemukan", "Peringatan"])
    for tab, df in zip(tabs, [O["main"], O["log"], O["marked"], O["notfound"], O["warns"]]):
        with tab:
            st.write(f"{len(df)} baris")
            st.dataframe(df, hide_index=True)


def main():
    st.set_page_config(page_title="Re-invent vs KRI", page_icon="🔍", layout="wide")

    page = st.sidebar.radio("Halaman", ["1. Bandingkan data", "2. Terapkan validasi (perbaikan)"])
    if page.startswith("2"):
        validation_page()
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
    main_bytes, report_bytes = build_outputs(R)

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
                    "Perlu dilengkapi", "Duplikat", "Log perubahan"])
    frames = [R["main"], R["kri_only"], R["re_only"], R["conf"], R["hip"], R["dismantle"],
              R["incomplete"], R["dup"], R["log"]]
    hints = {3: "Isi kolom **Validasi Final Hostname** di sheet `IP_Sama_Hostname_Beda` (file laporan), "
                "lalu upload di halaman 2 (Terapkan validasi).",
             4: "Isi kolom **Validasi IP_OAM** di sheet `Hostname_Sama_IP_Beda` (file laporan), "
                "lalu upload di halaman 2 (Terapkan validasi)."}
    for n, (tab, df) in enumerate(zip(tabs, frames)):
        with tab:
            if n in hints:
                st.caption(hints[n])
            st.write(f"{len(df)} baris")
            st.dataframe(df, hide_index=True)


if __name__ == "__main__":
    main()
