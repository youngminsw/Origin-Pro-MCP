# NOTES — origin2018-improvements 單元測試補齊

分支：`origin2018-improvements`（基底 `1e9741a`，被測 commit：`a454113`、`4ef4196`）
環境：Linux 雲端容器、Python 3.13、`mcp` 1.30（`mcp<2`）、pytest 9.1，**沒有 Origin**。
所有新測試都用 `tests/fakes.py` 的假 Origin、假 `winreg`、monkeypatch，不碰真實 COM／程序。

## 1. 測試結果

```
pytest：778 passed, 60 skipped, 0 failed
```

- 60 個 skip 全部是 `requires_origin`（需要 Windows + pywin32 + Origin COM）：`tests/test_live_*.py`
  與 `test_connection/test_graph/test_project/...` 裡標記為 live 的測試。Linux 上 skip 屬正常。
- 新增 118 個測試（見第 3 節）。改動前：657 passed / 3 failed / 60 skipped。
- `ruff`：repo **沒有** ruff 設定（pyproject 無 `[tool.ruff]`、無 `ruff.toml`），用 ruff 0.16.9 預設規則
  跑 `src tests`，和基底 `1e9741a` 逐檔比對：**沒有新增 lint 錯誤**（總數相同；新測試檔 0 錯誤）。
  - 這兩個 commit 原本新增了 7 處 `BLE001`/`S110`/`S112`（`except Exception: pass` 的 best-effort COM
    防護，repo 裡同樣寫法已有上百處）。已在這 7 行加上 `# noqa: ... - best-effort COM guard`，沒有改行為。
  - 舊式預設規則（`--select E4,E7,E9,F`）前後也一樣。

## 2. 修了哪些既有測試／程式，及原因

| 失敗的既有測試 | 原因 | 修法 |
|---|---|---|
| `test_dispatchers.py::test_registry_has_exactly_45_tools` | 新增 `close_origin` 工具，註冊數 45 → 46 | 改名為 `..._46_tools`，斷言 46，註解說明來源 |
| `test_shim.py::test_shim_schema_is_byte_identical_for_all_tools` | 同上，shim 轉發工具數 45 → 46 | 斷言改 46 |
| `test_doc_registry.py::test_readme_tool_count_matches_registry` | README 寫「45 total」 | **改 README**（不改測試）：標題改 46、Project Management 表格加 `close_origin`、`get_worksheet_data` 描述補分頁說明；`skills/publication-figure.md` 的「45 tools」同步改 46 |

注意：任務預期「舊測試預期 `get_worksheet_data` 只回 `{"columns": ...}`」會失敗，實際上**沒有**失敗——
既有測試（`test_worksheet_ops.py::test_get_worksheet_data_maps_missing_to_null`、`test_tool_guards.py`）
只讀 `out["columns"]` 或只檢查錯誤訊息，新增欄位不影響它們。

程式碼修正（皆不改 Windows 上的行為）：

1. **`daemon._registered_origin_images`：`os.path.basename` → `ntpath.basename`**。
   註冊表裡永遠是 Windows 路徑；在 Linux 上 `posixpath.basename(r"C:\PROGRA~1\...\ORIGIN~1.EXE")`
   不會以 `\` 切割，會回傳整串，導致單元測試無法驗證。`ntpath` 在 Windows 上就是 `os.path`，行為不變。
   （驗證：暫時改回 `os.path.basename` 時，`test_origin_image_names.py` 有 10 個測試失敗。）
2. `tools/project.py`：把夾在 `from ..app import mcp` 與其他 import 中間的 `import time` 移回標準庫
   import 區塊（純整理）。

沒有刪任何測試，也沒有加任何 skip。

## 3. 新增的測試（118 個）

### `tests/test_sheet_name_relaxed.py`（29）
- `labtalk_sheet_name` 接受：`"wpd_datasets _1_"`、中文（含空白）、點、括號、`-`。
- 拒絕：空字串、全空白、前後空白、非字串、`"` `[` `]` `!` `;` `\n` `\r` `\t`。
- `labtalk_name`：欄位名含 `sheet` 時錯誤訊息附「rename the sheet」；非 sheet 欄位或空值不附。
- 呼叫點：`get_worksheet_data`、`sort_worksheet`、`stats`（`_column_statistics_impl`）接受含空白的名稱，
  且 LabTalk 裡是 `page.active$ = "..."` 加引號；`interpolate`（範圍語法）維持嚴格並給出 rename 提示。

### `tests/test_worksheet_paging.py`（28）
- 預設 200 列（`GET_WORKSHEET_DEFAULT_MAX_ROWS == 200`）、`start_row`/`max_rows` 切片、最後一頁、
  逐頁走完整張表無缺漏、`max_rows=0` 回全部（含搭配 `start_row`）。
- `truncated`、`next_start_row`、`total_rows`、`total_cols`、`returned_rows` 正確。
- `start_row == total_rows` 回空頁；`start_row > total_rows` 報錯；負數、float、字串、bool 的
  `start_row`/`max_rows` 報錯，且**不會**先呼叫 `GetWorksheet`。空表的輸出。
- 缺值哨兵：`|v| >= 1e100` → `null`（含負值、恰好 1e100、9.9e99 保留、文字原樣），分頁中也正確。
- `long_names`/`units`：正常取值；`Units` 屬性不存在、`LongName`/`Units` 讀取丟例外、`Columns.Item`
  丟例外、`FindWorksheet` 丟例外、值不是字串、COM 欄數少於資料欄數 → 都回 `""` 且整個呼叫成功。

### `tests/test_close_origin.py`（33）
- 沒有 Origin／代理已死 → 「nothing to close」，且**不會**經 factory 重開。
- in-process（無 factory）與 attach（`mark_session_attached(True)`）沒給 `confirm=True` → 拒絕，
  不送 `doc -s;`、不 Exit；`confirm=True` 時 attach 模式即使 PID 已知、Exit 沒生效也**絕不 terminate**。
- `save=False`：順序是 `doc -s;` → `Exit`；`doc -s;` 丟例外也能繼續關。
- `save=True`：`Save("")` → `Exit`；空專案報錯；`Save("")` 回 False 報錯（要先 `save_project`），兩者都不 Exit。
- PID：用 `_pid_alive(pid)` 判斷結束（不 terminate）；Exit 被忽略 → `_terminate_pid(pid)` 並回報 PID；
  terminate 失敗 → 回「has not exited yet」；PID 未知 → 永不 terminate；關閉後清掉 PID。
- 關閉後 `drop_session_origin` 清掉 `project_path`、通知 ledger `(sid, None)`、保留 factory；
  下一次 `get_origin()` 開新的空實例，**不會** `Load` 舊專案（對照組：非預期 crash 仍會 reopen）。
- `_session_origin_pid` 從 `sys.modules["__main__"]` 讀 `_real_pid_tls`，且**不會** import 第二份
  `origin_pro_mcp.daemon`；也能從已 import 的 `origin_pro_mcp.daemon` 讀；兩個名字指向同一模組時只算一次；
  無效 PID（None/0/負數/字串/float）忽略；PID 是 thread-local。`_terminate_pid` 用 `sys.modules` 裡的
  `default_terminate_process`。POSIX 版 `_pid_alive` 的基本行為。
- daemon 整合：`Pool.acquire(attach=True)` 設 `Session.attached=True`，透過 worker 呼叫 `close_origin`
  被拒絕；隔離 session 為 False 且能正常關閉；第二個 attach 請求退回隔離實例，`attached=False`。

### `tests/test_origin_image_names.py`（19）
- 假 `winreg`：`LocalServer32` 為無引號 8.3 短路徑 `C:\PROGRA~1\ORIGIN~1\ORIGIN~1\ORIGIN~1.EXE`
  （含 ` -Embedding`、` /Automation`、尾端空白）→ `ORIGIN~1.EXE`；有引號
  `"C:\Program Files\X\Origin64.exe" -Embedding` → `Origin64.exe`。
- 兩個 ProgID 都讀、大小寫不重複；ProgID 或 key 缺漏時略過；只存在 64-bit view 的 key 也找得到；空值忽略。
- 沒有 `winreg`（`sys.modules["winreg"] = None`，等同 Linux）→ `()`。
- `_origin_image_names`：無覆寫時 = 預設清單 + 註冊表名稱；名稱已在預設清單時不重複（不分大小寫）；
  `ORIGIN_PRO_MCP_ORIGIN_IMAGE` 取代預設清單但註冊表名稱仍附加、大小寫不重複。
- 端到端：`_origin_process_pids` 能從 tasklist CSV 抓到 `ORIGIN~1.EXE` 的 PID；對照組記錄原本的 bug
  （沒有註冊表名稱時抓不到）。

### `tests/test_require_graph_fallback.py`（9）
- 一般路徑用 `FindGraphLayer`；`FindGraphLayer("[G]Layer1")` 回 None 時逐一走 `GraphPages`，回傳**同名頁**
  的第一個圖層；同名但沒有圖層的頁跳過；`GraphPages` 丟 COM 例外時視為找不到。
- 找不到時錯誤訊息列出現有的圖，無圖時顯示 `(none)`。
- 端到端：`export_graph` 經 fallback 成功；不存在的圖仍回「not found」。
- 突變檢查：暫時停用 fallback 後有 4 個測試失敗，確認測試確實覆蓋到它。

## 4. 只能在 Windows + Origin 2018 實測的項目（檢查清單）

假物件無法證明以下行為，下次在實機上請逐項確認：

- [ ] **FindGraphLayer**：Origin 2018 對 `[Graph1]Layer1` 確實回 None、fallback 拿到的 `page.Layers.Item(0)`
      可用於 `export_graph`、`set_axis` 等所有 `require_graph` 呼叫者（不只匯出）。
- [ ] 圖名比對是**大小寫敏感**的（`page.Name == graph_name`）；Origin 本身不分大小寫，
      若使用者給 `graph1` 會找不到——確認是否需要放寬。
- [ ] `get_worksheet_data`：Origin 2018 的 Column COM 物件確實有 `LongName`、`Units` 屬性（名稱拼法）；
      大表（數萬列）分頁時 `GetWorksheet` 仍是整張讀進來再切片，確認耗時／記憶體可接受。
- [ ] 含空白的 sheet 名稱在 15 個放寬的工具（worksheet.py 8、analysis.py 7）實際可用，
      尤其 `set_worksheet_data`、`export_worksheet`、`set_column_formula`、各 analysis 工具。
- [ ] sheet 名稱含 `%` 或 `$(`（目前**允許**）時，`page.active$ = "..."` 會不會被 LabTalk 字串替換展開。
      這是推測的風險，未驗證；由於這些工具都先用 `require_worksheet` 確認 sheet 存在，影響僅限於
      使用者自己就有這種 sheet 名稱的情況。若實測有問題，保守做法是把 `%`、`$` 加入 `_SHEET_FORBIDDEN`。
- [ ] `close_origin(save=False)`：`doc -s;` 後 Exit 不跳「Save changes?」對話框。
- [ ] `close_origin(save=True)`：已存過檔的專案 `Save("")` 存回原檔；未命名專案 `Save("")` 確實回 False。
- [ ] Origin 2018 `-Embedding` 實例忽略 Exit → 10 秒後被 terminate，工作管理員裡**沒有**殘留
      `ORIGIN~1.EXE`；Origin 2020 Exit 正常生效時不會走 terminate。
- [ ] 關閉後下一個工具呼叫開出**空的**新實例（不重開舊專案），且新實例的 PID 被正確擷取。
- [ ] attach 模式（`ORIGIN_PRO_MCP_ATTACH=1`）：沒有 `confirm` 時拒絕；`confirm=True` 時只送 Exit、
      絕不 terminate 使用者的 Origin。
- [ ] 註冊表讀取：實機的 `HKCR\Origin.Application\CLSID` → `CLSID\{...}\LocalServer32` 確實是
      8.3 短路徑，`_origin_image_names()` 包含 `ORIGIN~1.EXE`；對話框 watchdog、強制結束、
      孤兒回收（`record_spawned_pid`）在 Origin 2018 上恢復運作。
- [ ] `ORIGIN_PRO_MCP_ORIGIN_IMAGE` 設成錯的名稱時，PID 擷取仍靠註冊表名稱成功。
- [ ] 同時裝了多個 Origin 版本時，`ORIGIN~1.EXE` 這種短名稱可能對應到**別的** Origin 版本的程序
      （短名稱不唯一）。PID 擷取用 launch 前後的快照差集，理論上不受影響，但孤兒回收請實測。

## 5. 已知風險（從程式碼推論，未修改）

- **daemon 的 `Session.pid` 在重開後不會更新**（`daemon.py` 只在 `Session._run` 啟動時設定一次）。
  `close_origin` 之後下一個工具呼叫會經 factory 重開新的 Origin，但 `Session.pid`、對話框 watchdog、
  dispatch 逾時強制結束、reap、自動存檔判斷仍指向**舊的、已結束的 PID**：新實例卡住時 force-kill
  backstop 失效；理論上若 Windows 把舊 PID 重新分配給別的程序，逾時時可能誤殺。
  這個問題在原本的 crash-relaunch 路徑就存在，但 `close_origin` 讓「重開」變成常見路徑。
  因為修正會動到 daemon 的 watchdog 行為、且無法在 Linux 實測，這次**沒有**修改，只記錄於此。
  建議修法：每次 dispatch 結束後，若 thread 上的 `_real_pid_tls.pid` 與 `Session.pid` 不同，
  就更新 `Session.pid` 並重啟該 session 的 dialog watchdog。
