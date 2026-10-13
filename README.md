# excel2oracle

把 Excel(.xlsx) 一键灌进 Oracle staging 表的小工具：**选文件 → 填连接 → 测试连接 → 建表灌数**。可打包成单 exe，目标机器免装 Python、免装 Oracle Client。

## 特性

- **python-oracledb thin 模式**：目标库 Oracle **12.1+**，无需 Oracle Instant Client（「测试连接」成功会回显数据库版本）
- 表头自动清洗成合法 Oracle 列名（中文 / 重名 / 非法字符自动处理），**原始表头写进列注释**（`COMMENT ON COLUMN`），信息不丢
- 默认全 `VARCHAR2(4000 CHAR)`（staging 稳妥口径），可勾选自动推断 `NUMBER` / `DATE` / `TIMESTAMP`
- **drop 重建**（默认）/ **append 追加**（带列名一致性校验）双模式
- 批量插入：`executemany`，每批 1000 行，失败自动 rollback
- 环境可保存复用（`~/.excel2oracle/environments.json`）；**密码只存在内存，永不落盘**

## 用法

1. 选 .xlsx（**首行必须是表头**；多 sheet 可下拉选）
2. 填 Host / Port / 服务名（或勾「按 SID 连」）/ 用户名 / 密码 → **测试连接**
3. （可选）给环境起名并保存，下次一键载入（密码需重输）
4. 填表名（默认 `STG_<文件名>`，支持 `OWNER.TABLE`）→ 选模式 → **开始导入**

## 源码运行 & 打包 exe

```bash
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python app.py                                    # 源码运行
.venv\Scripts\pyinstaller --onefile --noconsole --name excel2oracle app.py
# → dist\excel2oracle.exe
```

测试（不需要 Oracle 实例）：

```bash
.venv\Scripts\python -m pytest
```

## 边界与已知限制（v1）

- 只支持 `.xlsx`；公式单元格取**缓存值**（即文件最后一次被 Excel 保存时的值）
- append 模式只校验**列名集合**，不校验类型；类型不匹配由数据库如实报错
- 单元格文本 > 4000 字符会插入失败（v1 未用 CLOB）
- Oracle 11g 及更早：thin 模式连不上（常见 `ORA-28040`），v1 未内置 thick 模式
- 需要 CONNECT + CREATE TABLE 权限（append 只需 INSERT/SELECT）

## License

MIT
