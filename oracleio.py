# -*- coding: utf-8 -*-
"""Oracle 连接 / 建表 / 批量插入（python-oracledb thin 模式，免装 Oracle Client）。"""
from __future__ import annotations

import time

import oracledb  # thin 模式为默认，无需 init_oracle_client()

import xlio

_THIN_HINT = ("提示：python-oracledb thin 模式只支持 Oracle 12.1+；若是 11g 及更早，"
              "会连不上（常见 ORA-28040 / DPY-4005）—— 需 thick 模式（本工具 v1 未内置）。")


def make_dsn(host: str, port: int, service: str, use_sid: bool = False) -> str:
    host, service = host.strip(), service.strip()
    if use_sid:
        return oracledb.makedsn(host, int(port), sid=service)
    return oracledb.makedsn(host, int(port), service_name=service)


def connect(host, port, service, user, password, use_sid=False):
    try:
        return oracledb.connect(user=user.strip(), password=password,
                                dsn=make_dsn(host, port, service, use_sid))
    except oracledb.Error as e:
        msg = str(e)
        if any(k in msg for k in ("ORA-28040", "DPY-4005", "DPI-1047")):
            raise ConnectionError(f"{msg}\n{_THIN_HINT}") from e
        raise ConnectionError(msg) from e


def server_version(conn) -> str:
    """尽力回显服务器版本（测试连接时展示；无权限则返回『已连通』）。"""
    for sql in ("SELECT banner FROM v$version WHERE banner LIKE 'Oracle%'",
                "SELECT version FROM product_component_version WHERE product LIKE 'Oracle Database%'"):
        try:
            cur = conn.cursor()
            cur.execute(sql)
            row = cur.fetchone()
            cur.close()
            if row and row[0]:
                return str(row[0])
        except oracledb.Error:
            continue
    return "已连通（无权限查版本）"


def split_table(table: str) -> tuple[str | None, str]:
    """'OWNER.TABLE' → ('OWNER','TABLE')；'TABLE' → (None,'TABLE')。"""
    if "." in table:
        owner, t = table.split(".", 1)
        return owner, t
    return None, table


class Stager:
    """在目标连接上建 staging 表并灌数。表名必须先经 xlio.sanitize_table_name 清洗。"""

    def __init__(self, conn):
        self.conn = conn

    def table_exists(self, table: str) -> bool:
        owner, t = split_table(table)
        cur = self.conn.cursor()
        try:
            if owner:
                cur.execute("SELECT COUNT(*) FROM all_tables "
                            "WHERE owner = :o AND table_name = :t", o=owner, t=t)
            else:
                cur.execute("SELECT COUNT(*) FROM user_tables WHERE table_name = :t", t=t)
            return cur.fetchone()[0] > 0
        finally:
            cur.close()

    def existing_columns(self, table: str) -> list[str]:
        owner, t = split_table(table)
        cur = self.conn.cursor()
        try:
            if owner:
                cur.execute("SELECT column_name FROM all_tab_columns "
                            "WHERE owner = :o AND table_name = :t ORDER BY column_id",
                            o=owner, t=t)
            else:
                cur.execute("SELECT column_name FROM user_tab_columns "
                            "WHERE table_name = :t ORDER BY column_id", t=t)
            return [r[0] for r in cur.fetchall()]
        finally:
            cur.close()

    def drop(self, table: str, log=print) -> None:
        cur = self.conn.cursor()
        try:
            cur.execute(f"DROP TABLE {table}")
        except oracledb.Error as e:
            if "ORA-00942" not in str(e):   # 表本来就不存在 → 无事
                raise
        else:
            log(f"已删除旧表 {table}")
        finally:
            cur.close()

    def create(self, table, columns, types, source_desc, log=print) -> None:
        cur = self.conn.cursor()
        try:
            for stmt in xlio.make_ddl(table, columns, types, source_desc):
                cur.execute(stmt)
        finally:
            cur.close()
        log(f"已建表 {table}（{len(columns)} 列，原始表头已写入 COMMENT）")

    def insert(self, table, columns, rows, log=print) -> tuple[int, float]:
        """executemany 分批插入；失败 rollback。返回 (行数, 耗秒)。"""
        sql = xlio.make_insert_sql(table, columns)
        cur = self.conn.cursor()
        t0 = time.time()
        total = 0
        try:
            for i in range(0, len(rows), xlio.INSERT_BATCH):
                batch = rows[i:i + xlio.INSERT_BATCH]
                cur.executemany(sql, batch)
                total += len(batch)
                log(f"  已插入 {total}/{len(rows)} 行")
            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise
        finally:
            cur.close()
        return total, time.time() - t0
