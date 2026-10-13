# -*- coding: utf-8 -*-
"""app.prepare_table 模式分支测试（FakeStager，不需要 Oracle 实例）。"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app

COLS = [("A", "编号"), ("B", "")]
TYPES = ["NUMBER", "VARCHAR2(10 CHAR)"]


class FakeStager:
    def __init__(self, exists=True, columns=None):
        self.exists = exists
        self.columns = columns if columns is not None else [c for c, _ in COLS]
        self.calls = []

    def table_exists(self, table):
        self.calls.append("exists")
        return self.exists

    def existing_columns(self, table):
        self.calls.append("cols")
        return list(self.columns)

    def drop(self, table, log=print):
        self.calls.append("drop")

    def create(self, table, columns, types, source_desc, log=print):
        self.calls.append("create")


def _no_log(*_a):
    pass


def test_drop_mode_drops_then_creates():
    stg = FakeStager()
    app.prepare_table(stg, "drop", "STG_T", COLS, TYPES, "src", log=_no_log)
    assert stg.calls == ["drop", "create"]


def test_append_mode_never_creates_or_drops():
    """回归（审查 HIGH #1）：append 分支绝不执行 CREATE/DROP。"""
    stg = FakeStager()
    app.prepare_table(stg, "append", "STG_T", COLS, TYPES, "src", log=_no_log)
    assert stg.calls == ["exists", "cols"]


def test_append_missing_table_raises():
    stg = FakeStager(exists=False)
    with pytest.raises(ValueError):
        app.prepare_table(stg, "append", "STG_T", COLS, TYPES, "src", log=_no_log)


def test_append_missing_column_raises():
    stg = FakeStager(columns=["A"])
    with pytest.raises(ValueError):
        app.prepare_table(stg, "append", "STG_T", COLS, TYPES, "src", log=_no_log)


def test_append_extra_columns_ok_and_logged():
    stg = FakeStager(columns=["A", "B", "C"])
    logged = []
    app.prepare_table(stg, "append", "STG_T", COLS, TYPES, "src", log=logged.append)
    assert any("C" in m for m in logged)


def test_unknown_mode_raises():
    """观察项①：mode 拼错不得静默落入 drop 分支（防程序化调用时误删表）。"""
    stg = FakeStager()
    with pytest.raises(ValueError):
        app.prepare_table(stg, "dropp", "STG_T", COLS, TYPES, "src", log=_no_log)
