"""Headless AppTest smoke tests. Runs against the test_excel throwaway copy only.

AppTest.from_file needs an absolute path (it resolves relative paths against the
calling script, not the CWD) -- see CLAUDE.md.
"""

from __future__ import annotations

import os

from streamlit.testing.v1 import AppTest

APP_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app.py")


def _run_app(monkeypatch, test_excel, **env) -> AppTest:
    monkeypatch.setenv("DASHBOARD_EXCEL", os.path.abspath(test_excel))
    for key in ("DASHBOARD_READONLY", "APP_PASSWORD"):
        monkeypatch.delenv(key, raising=False)
    for key, value in env.items():
        monkeypatch.setenv(key, value)
    at = AppTest.from_file(APP_PATH, default_timeout=180)
    at.run()
    return at


def test_app_loads_without_exception(monkeypatch, test_excel):
    at = _run_app(monkeypatch, test_excel)
    assert at.exception == []
    assert len(at.tabs) == 14


def test_app_read_only_disables_save_buttons(monkeypatch, test_excel):
    at = _run_app(monkeypatch, test_excel, DASHBOARD_READONLY="1")
    assert at.exception == []
    save_keys = ("btn_save_mon", "btn_save_ak", "btn_save_pot", "btn_save_tgt")
    save_buttons = [b for b in at.button if b.key in save_keys]
    assert len(save_buttons) == len(save_keys)
    assert all(b.disabled for b in save_buttons)
    notices = [i.value for i in at.info] + [w.value for w in at.warning]
    assert any("baca-saja" in n.lower() for n in notices)


def test_app_password_gate_blocks_until_correct(monkeypatch, test_excel):
    at = _run_app(monkeypatch, test_excel, APP_PASSWORD="rahasia-test")
    assert at.exception == []
    assert len(at.tabs) == 0  # locked behind the login form

    at.text_input[0].input("password-salah")
    at.button[0].click()
    at.run()
    assert [e.value for e in at.error] == ["Password salah."]
    assert len(at.tabs) == 0

    at.text_input[0].input("rahasia-test")
    at.button[0].click()
    at.run()
    assert at.exception == []
    assert len(at.tabs) == 14
