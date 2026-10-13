# -*- coding: utf-8 -*-
"""Excel → Oracle staging 导入工具（tkinter GUI 入口）。"""
from __future__ import annotations

import json
import queue
import threading
import traceback
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import oracleio
import xlio

ENV_DIR = Path.home() / ".excel2oracle"
ENV_FILE = ENV_DIR / "environments.json"   # 只存 host/port/service/user，永不存密码


def load_envs() -> dict:
    try:
        return json.loads(ENV_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_envs(envs: dict) -> None:
    ENV_DIR.mkdir(parents=True, exist_ok=True)
    ENV_FILE.write_text(json.dumps(envs, ensure_ascii=False, indent=2), encoding="utf-8")


def prepare_table(stg, mode, table, columns, types, source_desc, log=print):
    """模式分支（审查 #1 回归点）：
    drop 模式 → drop 旧表 + 建新表（含 COMMENT）；append 模式 → 只做存在性与列名
    一致性校验，**绝不建表/删表**（append 只需 INSERT/SELECT 权限）。"""
    if mode not in ("drop", "append"):
        raise ValueError(f"未知模式：{mode!r}（只允许 drop / append）")
    if mode == "append":
        if not stg.table_exists(table):
            raise ValueError(f"append 模式但表 {table} 不存在 —— 先用 drop 重建模式建表")
        existing = set(stg.existing_columns(table))
        want = {c for c, _ in columns}
        missing = sorted(want - existing)
        if missing:
            raise ValueError(f"表 {table} 缺列：{missing}")
        extra = sorted(existing - want)
        if extra:
            log(f"⚠️ 表中多出 {len(extra)} 列将保持 NULL：{extra[:10]}")
        return
    stg.drop(table, log=log)
    stg.create(table, columns, types, source_desc, log=log)


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Excel → Oracle staging 导入器")
        root.geometry("800x660")
        self.q: queue.Queue = queue.Queue()
        self.busy = False

        self.file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.cols_var = tk.StringVar(value="未选择文件")
        self.env_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.host_var = tk.StringVar()
        self.port_var = tk.StringVar(value="1521")
        self.svc_var = tk.StringVar()
        self.sid_var = tk.BooleanVar(value=False)
        self.user_var = tk.StringVar()
        self.pwd_var = tk.StringVar()
        self.table_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="drop")
        self.infer_var = tk.BooleanVar(value=False)
        self.conn_status = tk.StringVar(value="未测试")

        self._build()
        self._refresh_envs()
        root.after(100, self._pump)

    # ---------- UI ----------
    def _build(self):
        pad = {"padx": 6, "pady": 3}
        frm_file = ttk.LabelFrame(self.root, text="① Excel 文件（.xlsx，首行须为表头）")
        frm_file.pack(fill="x", **pad)
        ttk.Button(frm_file, text="选择文件…", command=self.pick_file).grid(row=0, column=0, **pad)
        ttk.Label(frm_file, textvariable=self.file_var, wraplength=540).grid(row=0, column=1, **pad)
        ttk.Label(frm_file, text="Sheet:").grid(row=1, column=0, **pad)
        self.sheet_box = ttk.Combobox(frm_file, textvariable=self.sheet_var,
                                      state="readonly", width=24)
        self.sheet_box.grid(row=1, column=1, sticky="w", **pad)
        ttk.Label(frm_file, textvariable=self.cols_var, wraplength=540).grid(
            row=2, column=1, sticky="w", **pad)

        frm_conn = ttk.LabelFrame(self.root, text="② Oracle 连接（密码只在内存，不落盘）")
        frm_conn.pack(fill="x", **pad)
        ttk.Label(frm_conn, text="已存环境:").grid(row=0, column=0, sticky="e", **pad)
        self.env_box = ttk.Combobox(frm_conn, textvariable=self.env_var,
                                    state="readonly", width=18)
        self.env_box.grid(row=0, column=1, sticky="w", **pad)
        ttk.Button(frm_conn, text="载入", command=self.load_env).grid(row=0, column=2, **pad)
        ttk.Label(frm_conn, text="名称:").grid(row=1, column=0, sticky="e", **pad)
        ttk.Entry(frm_conn, textvariable=self.name_var, width=18).grid(
            row=1, column=1, sticky="w", **pad)
        ttk.Button(frm_conn, text="💾 保存环境(不含密码)", command=self.save_env).grid(
            row=1, column=2, **pad)
        ttk.Label(frm_conn, text="Host:").grid(row=2, column=0, sticky="e", **pad)
        ttk.Entry(frm_conn, textvariable=self.host_var, width=26).grid(
            row=2, column=1, sticky="w", **pad)
        ttk.Label(frm_conn, text="Port:").grid(row=2, column=2, sticky="e")
        ttk.Entry(frm_conn, textvariable=self.port_var, width=7).grid(
            row=2, column=3, sticky="w", pady=3)
        ttk.Label(frm_conn, text="服务名/SID:").grid(row=3, column=0, sticky="e", **pad)
        ttk.Entry(frm_conn, textvariable=self.svc_var, width=26).grid(
            row=3, column=1, sticky="w", **pad)
        ttk.Checkbutton(frm_conn, text="按 SID 连", variable=self.sid_var).grid(
            row=3, column=2, **pad)
        ttk.Label(frm_conn, text="用户名:").grid(row=4, column=0, sticky="e", **pad)
        ttk.Entry(frm_conn, textvariable=self.user_var, width=18).grid(
            row=4, column=1, sticky="w", **pad)
        ttk.Label(frm_conn, text="密码:").grid(row=4, column=2, sticky="e")
        ttk.Entry(frm_conn, textvariable=self.pwd_var, width=18, show="*").grid(
            row=4, column=3, sticky="w", pady=3)
        ttk.Button(frm_conn, text="🔌 测试连接", command=self.test_conn).grid(
            row=5, column=1, sticky="w", **pad)
        ttk.Label(frm_conn, textvariable=self.conn_status, wraplength=470).grid(
            row=5, column=2, columnspan=2, sticky="w", **pad)

        frm_tgt = ttk.LabelFrame(self.root, text="③ 目标 staging 表")
        frm_tgt.pack(fill="x", **pad)
        ttk.Label(frm_tgt, text="表名:").grid(row=0, column=0, sticky="e", **pad)
        ttk.Entry(frm_tgt, textvariable=self.table_var, width=32).grid(
            row=0, column=1, sticky="w", **pad)
        ttk.Label(frm_tgt, text="（可写 OWNER.TABLE）").grid(row=0, column=2, **pad)
        ttk.Radiobutton(frm_tgt, text="drop 重建（默认）", value="drop",
                        variable=self.mode_var).grid(row=1, column=1, sticky="w", **pad)
        ttk.Radiobutton(frm_tgt, text="append 追加（列名须一致）", value="append",
                        variable=self.mode_var).grid(row=1, column=2, sticky="w", **pad)
        ttk.Checkbutton(frm_tgt, text="自动推断类型 NUMBER/DATE/TIMESTAMP（默认全 VARCHAR2）",
                        variable=self.infer_var).grid(row=2, column=1, columnspan=2,
                                                      sticky="w", **pad)

        frm_run = ttk.Frame(self.root)
        frm_run.pack(fill="x", **pad)
        self.run_btn = ttk.Button(frm_run, text="▶ 开始导入", command=self.run_import)
        self.run_btn.pack(side="left", padx=6)
        self.progress = ttk.Progressbar(frm_run, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True, padx=6)

        self.log_text = ScrolledText(self.root, height=12, state="disabled",
                                     font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True, padx=6, pady=6)

    # ---------- 后台 → UI 队列 ----------
    def log(self, msg: str):
        self.q.put(("log", f"[{datetime.now():%H:%M:%S}] {msg}"))

    def _pump(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self.log_text.configure(state="normal")
                    self.log_text.insert("end", payload + "\n")
                    self.log_text.see("end")
                    self.log_text.configure(state="disabled")
                elif kind == "conn":
                    ok, text = payload
                    self.conn_status.set(("✅ " if ok else "❌ ") + text)
                elif kind == "done":
                    self.busy = False
                    self.run_btn.configure(state="normal")
                    self.progress.stop()
                    if payload is None:
                        messagebox.showinfo("完成", "导入完成，详见日志。")
                    else:
                        messagebox.showerror("失败", payload)
        except queue.Empty:
            pass
        self.root.after(100, self._pump)

    def _start_worker(self, target):
        if self.busy:
            return
        self.busy = True
        self.run_btn.configure(state="disabled")
        self.progress.start(12)
        threading.Thread(target=target, daemon=True).start()

    # ---------- 事件 ----------
    def pick_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            sheets = xlio.list_sheets(path)
            preview = xlio.read_sheet(path, sheets[0] if sheets else None)
        except Exception as e:
            messagebox.showerror("读取失败",
                                 f"{e}\n注意：只支持 .xlsx；公式单元格取的是缓存值。")
            return
        self.file_var.set(path)
        self.sheet_box["values"] = sheets
        self.sheet_var.set(sheets[0] if sheets else "")
        self.table_var.set("STG_" + xlio.sanitize_identifier(Path(path).stem, "T")[:20])
        heads = ", ".join(f"{c}←{h}" if h else c for c, h in preview.columns[:6])
        more = "…" if len(preview.columns) > 6 else ""
        self.cols_var.set(
            f"{len(preview.columns)} 列 × {len(preview.rows)} 行 ｜ {heads}{more}")

    def load_env(self):
        envs = load_envs()
        name = self.env_var.get()
        if name not in envs:
            return
        e = envs[name]
        self.name_var.set(name)
        self.host_var.set(e.get("host", ""))
        self.port_var.set(str(e.get("port", 1521)))
        self.svc_var.set(e.get("service", ""))
        self.sid_var.set(bool(e.get("use_sid", False)))
        self.user_var.set(e.get("user", ""))
        self.pwd_var.set("")
        self.conn_status.set(f"已载入「{name}」（密码需重输）")

    def save_env(self):
        if not all((self.name_var.get().strip(), self.host_var.get().strip(),
                    self.svc_var.get().strip(), self.user_var.get().strip())):
            messagebox.showwarning("缺字段", "名称 / Host / 服务名 / 用户名都要填（密码不存）。")
            return
        try:
            port = int(self.port_var.get() or 1521)
        except ValueError:
            messagebox.showwarning("Port 非法", "Port 必须是数字（默认 1521）。")
            return
        if not 1 <= port <= 65535:
            messagebox.showwarning("Port 超范围", "Port 须在 1-65535 之间（默认 1521）。")
            return
        envs = load_envs()
        envs[self.name_var.get().strip()] = {
            "host": self.host_var.get().strip(),
            "port": port,
            "service": self.svc_var.get().strip(),
            "use_sid": self.sid_var.get(),
            "user": self.user_var.get().strip(),
        }
        save_envs(envs)
        self._refresh_envs()
        self.conn_status.set(f"已保存「{self.name_var.get().strip()}」→ {ENV_FILE}（不含密码）")

    def _refresh_envs(self):
        names = list(load_envs().keys())
        self.env_box["values"] = names
        if names and not self.env_var.get():
            self.env_var.set(names[0])

    # ---------- 动作 ----------
    def _conn_fields(self):
        host = self.host_var.get().strip()
        svc = self.svc_var.get().strip()
        user = self.user_var.get().strip()
        pwd = self.pwd_var.get()
        if not (host and svc and user and pwd):
            raise ValueError("Host / 服务名 / 用户名 / 密码都要填")
        try:
            port = int(self.port_var.get() or 1521)
        except ValueError:
            raise ValueError("Port 必须是数字")
        return host, port, svc, user, pwd, self.sid_var.get()

    def test_conn(self):
        try:
            fields = self._conn_fields()
        except ValueError as e:
            self.conn_status.set(f"❌ {e}")
            return

        def work():
            try:
                conn = oracleio.connect(*fields)
                try:
                    ver = oracleio.server_version(conn)
                finally:
                    conn.close()
                self.q.put(("conn", (True, ver)))
            except Exception as e:
                self.q.put(("conn", (False, str(e))))

        threading.Thread(target=work, daemon=True).start()

    def run_import(self):
        errors = []
        if not self.file_var.get():
            errors.append("先选 Excel 文件")
        if not self.table_var.get().strip():
            errors.append("填目标表名")
        if errors:
            messagebox.showwarning("缺输入", "\n".join(errors))
            return
        try:
            fields = self._conn_fields()
        except ValueError as e:
            messagebox.showwarning("缺输入", str(e))
            return
        path = self.file_var.get()
        sheet = self.sheet_var.get() or None
        table = xlio.sanitize_table_name(self.table_var.get(), "STG_DATA")
        mode = self.mode_var.get()
        infer = self.infer_var.get()

        def work():
            err = None
            conn = None
            try:
                self.log(f"读取 {Path(path).name} · sheet「{sheet or '(第一个)'}」…")
                data = xlio.read_sheet(path, sheet)
                self.log(f"读到 {len(data.columns)} 列 × {len(data.rows)} 行"
                         f"（列名已清洗，原始表头进 COMMENT）")
                types = xlio.infer_types(data.rows, len(data.columns), infer)
                rows = xlio.coerce_rows(data.rows, types)
                self.log(f"连接 Oracle（{fields[0]}:{fields[1]}/{fields[2]}）…")
                conn = oracleio.connect(*fields)
                stg = oracleio.Stager(conn)
                src = (f"staging from {Path(path).name}[{data.name}] "
                       f"by excel2oracle @ {datetime.now():%Y-%m-%d %H:%M}")
                prepare_table(stg, mode, table, data.columns, types, src, log=self.log)
                n, sec = stg.insert(table, data.columns, rows, log=self.log)
                self.log(f"✅ 完成：{n} 行 → {table}（{len(types)} 列），耗时 {sec:.1f}s")
            except Exception as e:
                err = f"{type(e).__name__}: {e}"
                self.log(f"❌ {err}")
                self.log(traceback.format_exc(limit=3))
            finally:
                if conn is not None:
                    try:
                        conn.close()
                    except Exception:
                        pass
            self.q.put(("done", err))

        self._start_worker(work)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
