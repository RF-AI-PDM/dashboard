"""Tests for data_loader.py: known Excel totals, formatters, and date parsing."""

from __future__ import annotations

from datetime import date

import pytest

import data_loader as dl

# data/Data.xlsx is a live file the user edits, so its totals change over time.
# Assert internal consistency (KPI realisasi == sum of the source column) instead
# of a hardcoded snapshot, which would go stale the next time the user edits it.


def test_load_all_known_totals(test_excel):
    d = dl.load_all(test_excel)
    expected_mon = float(d["monitoring"]["Nilai Kontrak Setelah PPN (IPS)"].sum())
    expected_ak = float(d["monitoring_akumulatif"]["Nilai Kontrak Mitra Setelah PPN"].sum())
    assert abs(d["pencapaian"]["realisasi_tahun"] - expected_mon) < 1
    assert abs(d["akumulatif_kpi"]["realisasi_tahun"] - expected_ak) < 1


def test_load_all_shape(test_excel):
    d = dl.load_all(test_excel)
    for key in ("monitoring", "pencapaian", "akumulatif_kpi", "monitoring_akumulatif",
                "potensi", "penagihan", "file_mtime"):
        assert key in d
    assert not d["monitoring"].empty
    assert "Judul Pekerjaan" in d["monitoring"].columns
    # Template/marker rows (Excel rows 2-4) must be filtered out.
    assert d["monitoring"]["Judul Pekerjaan"].str.strip().eq("").sum() == 0


def test_load_all_status_normalized(test_excel):
    d = dl.load_all(test_excel)
    mon = d["monitoring"]
    assert set(mon["STATUS"].unique()) <= {"SELESAI", "PROSES", "CANCEL", "BELUM ADA STATUS"}
    # Originals are preserved alongside the normalized column.
    assert "STATUS (Asli)" in mon.columns


@pytest.mark.parametrize("value, expected", [
    (0, "Rp 0"),
    (16618722076, "Rp 16.618.722.076"),
    (-6844328390, "-Rp 6.844.328.390"),
    (1500000.4, "Rp 1.500.000"),
    (None, "-"),
    (float("nan"), "-"),
    ("not a number", "-"),
])
def test_fmt_rp(value, expected):
    assert dl.fmt_rp(value) == expected


@pytest.mark.parametrize("value, expected", [
    (0.6924, "69,2%"),
    (0.0, "0,0%"),
    (1.0, "100,0%"),
    (None, "-"),
    (float("nan"), "-"),
])
def test_fmt_pct(value, expected):
    assert dl.fmt_pct(value) == expected


@pytest.mark.parametrize("text, expected", [
    ("04 September 2026", date(2026, 9, 4)),
    ("4 September 2026", date(2026, 9, 4)),
    ("15/03/2026", date(2026, 3, 15)),
    ("15-03-2026", date(2026, 3, 15)),
    ("2026-03-15", date(2026, 3, 15)),
    ("", None),
    (None, None),
    ("tidak ada tanggal", None),
])
def test_parse_id_date(text, expected):
    assert dl.parse_id_date(text) == expected


@pytest.mark.parametrize("text, expected", [
    ("14 hari kalender", 14),
    ("30 Hari", 30),
    ("20 (Dua Puluh) hari kalender", 20),
    ("", None),
    (None, None),
    ("tidak ada durasi", None),
])
def test_parse_duration_days(text, expected):
    assert dl.parse_duration_days(text) == expected
