# tests/test_history.py
import sys
import os
import json
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from history import save_period, load_last_period, get_previous_ids


def _tmp_path():
    f = tempfile.NamedTemporaryFile(suffix='.json', delete=False)
    f.close()
    os.unlink(f.name)
    return f.name


def test_save_and_load_round_trip():
    path = _tmp_path()
    try:
        save_period(path, '2026-05/06', ['1001', '1002', '1003'])
        last = load_last_period(path)
        assert last['period'] == '2026-05/06'
        assert set(last['ids']) == {'1001', '1002', '1003'}
    finally:
        if os.path.exists(path):
            os.unlink(path)


def test_load_returns_none_when_no_history():
    path = _tmp_path()
    last = load_last_period(path)
    assert last is None


def test_get_previous_ids_returns_set():
    path = _tmp_path()
    try:
        save_period(path, '2026-05/06', ['1001', '1002'])
        ids = get_previous_ids(path)
        assert ids == {'1001', '1002'}
    finally:
        if os.path.exists(path):
            os.unlink(path)


def test_get_previous_ids_returns_empty_when_no_history():
    path = _tmp_path()
    ids = get_previous_ids(path)
    assert ids == set()


def test_save_appends_history():
    path = _tmp_path()
    try:
        save_period(path, '2026-03/04', ['900'])
        save_period(path, '2026-05/06', ['1001', '1002'])
        with open(path) as f:
            data = json.load(f)
        assert len(data['history']) == 2
        assert data['history'][-1]['period'] == '2026-05/06'
    finally:
        if os.path.exists(path):
            os.unlink(path)
