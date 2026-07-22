# history.py
import json
import os
from typing import List, Optional, Set


def _load_data(path: str) -> dict:
    if not os.path.exists(path):
        return {'history': []}
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_period(path: str, period: str, ids: List[str]) -> None:
    """保存当期名单到历史文件。"""
    data = _load_data(path)
    data['history'].append({'period': period, 'ids': list(ids)})
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_last_period(path: str) -> Optional[dict]:
    """返回最近一期历史记录，无历史时返回 None。"""
    data = _load_data(path)
    if not data['history']:
        return None
    return data['history'][-1]


def get_previous_ids(path: str) -> Set[str]:
    """返回上期进入能力提升的工号集合。"""
    last = load_last_period(path)
    if last is None:
        return set()
    return set(last['ids'])
