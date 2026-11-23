# origin2018-fixes-2 工作筆記

分支：`origin2018-fixes-2`，從 `origin/origin2018-improvements`（4ef4196）切出。
`origin2018-improvements` 本身沒有動過，另一個雲端 session 正在那裡補測試，之後在本機合併兩個分支。

雲端環境是 Linux、沒有 Origin，所以以下都只用假 COM 物件跑過單元測試，**還沒在 Origin 2018 上實測**。最後一節是實測清單。

---

## 1. 打包

| 檔案 | 改了什麼 |
|---|---|
| `pyproject.toml` | 相依從 `mcp>=1.0.0` 改成 `mcp>=1.8.0,<2` |
| `requirements.txt` | 一起改成 `mcp>=1.8.0,<2`（原本也寫 `mcp>=1.0.0`，會有一樣的問題） |
| `src/origin_pro_mcp/__main__.py`（新檔） | 讓 `python -m origin_pro_mcp` 會呼叫 `app.main()` |
| `src/origin_pro_mcp/app.py` | 結尾加上 `if __name__ == "__main__":` |

**`app.py` 的 `__main__` 寫法跟原本的規格不同，原因如下**：
用 `python -m origin_pro_mcp.app` 執行時，這個檔案是以 `__main__` 的名稱載入，等於多了第二份模組。工具會註冊到 `origin_pro_mcp.app.mcp` 上，而不是 `__main__` 那一份的 `mcp`。如果照字面寫成 `main()`，在 in-process 模式（`ORIGIN_PRO_MCP_USE_DAEMON=0`）下，啟動的 server **一個工具都沒有**。我在雲端用真的 stdio 測過：照字面寫時，`tools/list` 沒有回傳任何工具；改成轉呼叫真正的 `origin_pro_mcp.app.main` 後，兩種啟動方式都回傳 46 個工具。預設的 daemon/shim 模式不受這個問題影響。

## 2. `list_worksheets(details=False)`（`tools/worksheet.py`）

- 新增參數 `details: bool = False`。
- `details=False`：輸出和舊版**逐字相同**，有測試直接比對舊版的 JSON 字串。
- `details=True` 會**額外**回傳兩個欄位：
  - `sheet_details: [{"book", "sheet", "rows", "cols"}]`：讀 `o.FindWorksheet("[book]sheet")` 回傳物件的 `Rows`、`Cols`。每個屬性各自包 try/except，讀不到、不是數字或是 bool 時回 `null`。不會呼叫 `GetWorksheet`，也就是不讀儲存格資料。
  - `graph_details: [{"name", "layers", "plots", "empty"}]`：`GraphPages.Item(i).Layers.Item(j).DataPlots.Count` 逐層加總。只要有任何一層讀不到，`plots` 就是 `null`，不會把只加到一半的數字當成總數。`plots` 是 `null` 時，`empty` 也是 `null`。讀不到名稱的圖頁會跳過，和原本的 `graphs` 清單一致。
- 沒有新增任何 LabTalk，也不切換作用中視窗。有測試確認 `details=True` 和 `details=False` 送出的 LabTalk 完全相同，而且圖沒有被 activate。原本的保存／還原作用中視窗邏輯沒有改。
- **已知限制（推論，尚未在 2018 上驗證）**：程式碼裡有關於「從 .opju 載入的圖在 activate 之前，COM 回報 0 個 plot」的註解（請見 `tools/graph.py` 和 `FakeGraph(loaded=True)`）。所以剛載入的專案中，有資料的圖可能會被回報成 `plots: 0, empty: true`。我已把這點寫進 docstring。為了不切換視窗，這裡沒有用 activate 去繞過這個問題。
- **風險**：`origin_connection.sheet_names` 的註解提到，在大型專案上做 `page.Layers.Item(j)` 這種深層 COM 走訪，曾在 Origin 2020 上讓 Origin 當掉。`graph_details` 對圖頁做的正是這種走訪，這是規格要求的作法，而且只有在 `details=True` 時才會執行。預設不會跑到這段，所以舊的使用方式不受影響。

## 3. 收尾時關掉「空白」的 Origin（`daemon.py`）

- `Session._graceful_reap` 的 detach 分支（也就是沒有設定 `ORIGIN_PRO_MCP_REAP_CLOSE` 時）會先呼叫 `Session._close_if_empty()`，再照舊結束。整段包在 try/except 裡，出錯也不會影響收尾流程。
- 關閉 Origin 的條件，**全部成立**才會關：
  1. `self.attached` 是 False，也就是不是使用者自己開的 Origin。
  2. `WorksheetPages`、`GraphPages`、`MatrixPages` 三個 `.Count` 都讀得到，而且是整數，總和為 0。任何一個讀不到或不是數字，就當作「不確定」，保留不關。
  3. 額外的保守條件（規格沒有要求）：如果 `o.Pages.Count` 讀得到而且大於 0，也保留。這是為了避免一個只有 Notes 視窗的專案被誤關。讀不到 `Pages` 時就忽略這一條，只看上面三個數字。**`Application.Pages` 在 2018 上是否存在，我不確定（推論）。**
- 關閉流程：`Execute("doc -s;")`，接著 `Exit()`（沒有 Exit 時改用 `Close()`），然後用 `self.pid` 每 0.1 秒確認一次程序是否結束，最多等 **3.5 秒**。程序還在、而且 `self.pid` 不是 None，就呼叫 `default_terminate_process(pid)`（可注入），再確認 1 秒。`self.pid` 是 None 時只送 Exit，不強制結束。
- **為什麼是 3.5 秒而不是 5 秒**：reap 的 watchdog 第二階段預設寬限期是 5 秒（`reap_grace`）。如果等待加上確認超過 5 秒，watchdog 會先把同一個 PID 殺掉，結束 session，接著 graceful 的完成回呼又會再跑一次 `_finish_reap`。這是既有的競態：如果中間剛好有重新連線，可能會誤刪新的 session。所以把 3.5 + 1 秒壓在 5 秒以內。
- 程序是否還在：新增 `default_process_running(pid)`。Windows 上用 `WaitForSingleObject(handle, 0) == WAIT_TIMEOUT` 判斷，和 `tools/project.py` 的 `_pid_alive` 相同；原本的 `default_is_alive` 只檢查能不能 OpenProcess，程序已結束但 handle 還被持有時也會回 True。開不了 handle 時回 False，所以會被當成已結束，不會去殺。這也是偏保守的作法。
- 注入點：`Session(terminate_process=, process_running=, empty_close_wait=)`、`Pool(terminate_process=, process_running=)`。`Daemon.start` 會把它既有的 `terminate_process` 和 `is_alive` 傳下去，測試因此不會對真實的 PID 送訊號。
- 結果會記在 `session.empty_close_result`，值是 `kept`、`closed`、`forced`、`exit-sent`、`still-running` 其中之一，方便除錯。
- 有內容的專案一律照舊 detach 保留，不存檔、也不會送 `doc -s;`。`ORIGIN_PRO_MCP_REAP_CLOSE=1` 的舊行為完全沒動，有測試確認走這條路時不會呼叫 `_close_if_empty`。

## 4. 順手修的既有失敗測試（合併時要注意）

`origin2018-improvements` 新增了 `close_origin`，工具數從 45 變成 46。這導致 base 分支本來就有 3 個測試失敗，和這次的修改無關：

- `tests/test_dispatchers.py::test_registry_has_exactly_45_tools`：已改名為 `..._46_tools`，斷言改成 46。
- `tests/test_shim.py::test_shim_schema_is_byte_identical_for_all_tools`：斷言改成 46。
- `tests/test_doc_registry.py::test_readme_tool_count_matches_registry`：README 改成「46 total」，並在 Project Management 表格加上 `close_origin` 一列。

**另一個補測試的 session 可能也會改這幾行**。如果合併時衝突，保留任一邊都可以，只要總數是 46、README 有 `close_origin` 那一列。

## 測試結果（Linux，Python 3.13，mcp 1.x）

```
python -m pytest -q
699 passed, 60 skipped
```

- 60 個 skip 都是 `requires_origin` 或 live 測試，Linux 上本來就會 skip。
- 新增的測試檔：
  - `tests/test_list_worksheets_details.py`（12 個）：舊輸出逐字相同；屬性存在、屬性不存在、單一屬性出錯；`FindWorksheet` 出錯；不讀儲存格資料；多層圖加總；單層出錯時回 null；單一壞頁不影響其他頁；`GraphPages` 讀不到；不切換視窗。
  - `tests/test_empty_origin_reap.py`（21 個）：空白專案會關；Exit 沒生效時強制結束；kill 也失敗時回報 `still-running`；有內容時保留（三種視窗各測一次）；任一 Count 讀取失敗時保留；Count 不是數字時保留；`Pages` > 0 時保留；attach 模式保留；pid 是 None 時不強制結束；Exit 拋例外時不影響收尾；`REAP_CLOSE=1` 的路徑不變；Pool 和 Daemon 有把注入函式接上；等待時間有上限。
  - `tests/test_packaging.py`（6 個）：mcp 版本限制；`python -m origin_pro_mcp` 和 `python -m origin_pro_mcp.app` 都呼叫真正的 `app.main`。
- 此外在雲端用 subprocess 走真的 stdio 測過：`python -m origin_pro_mcp` 和 `python -m origin_pro_mcp.app`，在 daemon 模式與 `ORIGIN_PRO_MCP_USE_DAEMON=0` 下，`tools/list` 都回傳 46 個工具，`list_worksheets` 的 schema 也有 `details`。
- ruff：repo 沒有 ruff 設定。用 `ruff check --isolated --select E,F,W --ignore E501` 檢查所有改過和新增的檔案，全部通過。不加 `--select` 時錯誤很多，是使用者層級的設定造成的，修改前的檔案也一樣有那些錯誤。

---

## Windows + Origin 2018 實測清單

建議先用 `ORIGIN_PRO_MCP_USE_DAEMON=0` 或 CLI 做 1～3，再用正常的 daemon 模式做 4～8。

### 打包
- [ ] 1. `pip install -e .` 裝到的 mcp 是 1.x：`python -c "import importlib.metadata as m; print(m.version('mcp'))"`。
- [ ] 2. `python -m origin_pro_mcp` 能啟動，MCP client 看得到 46 個工具。

### `list_worksheets(details=True)`
- [ ] 3. **`Worksheet.Rows`、`Worksheet.Cols` 在 2018 上是否存在**：開一個有資料的 workbook，呼叫 `list_worksheets(details=true)`，`rows`、`cols` 不是 null，而且數字正確。若是 null，代表 2018 沒有這兩個屬性，需要改用別的方式讀。
  - 補充：`Rows` 回傳的是「有資料的列數」還是「工作表配置的列數」，請和 Origin 畫面對照。
- [ ] 4. **`GraphLayer.DataPlots.Count` 在 2018 上是否存在**：有兩條曲線的圖 → `plots: 2, empty: false`；空白的圖 → `plots: 0, empty: true`；多層圖 → 各層的總和。
- [ ] 5. 開啟一個已存好的 .opju/.opj，**不要先點任何圖**，直接呼叫 `details=true`，看有資料的圖是否被誤報成 `plots: 0`（請見第 2 節的已知限制）。
- [ ] 6. 呼叫前後，Origin 的作用中視窗沒有改變。
- [ ] 7. 視窗很多的大專案上，`details=true` 不會卡住或讓 Origin 當掉（第 2 節提到的深層 COM 走訪風險）。

### session 結束時關掉空白 Origin
- [ ] 8. **空白 Origin 是否真的被關掉**：讓 MCP 開一個 Origin，什麼都不建立（或建立後刪光），結束 client 或 session，等收尾跑完（預設寬限期 + 數秒）。`tasklist | findstr /i origin` 看不到這個實例；在 2018 上要確認 Exit 無效時，有用 PID 強制結束。
- [ ] 9. **有內容的 Origin 是否保留**：session 裡建立一個 workbook（不存檔），結束 session。Origin 還在，資料還在，標題列沒有變成 recovery 檔名。
- [ ] 10. attach 模式下（使用者自己開的 Origin，即使是空的），結束 session 後不會被關掉。
- [ ] 11. 只有 Notes 視窗的專案（三種 Count 都是 0），結束 session 後是否被保留。這要看 2018 是否有 `Application.Pages`：如果被關掉，代表 `Pages` 不存在，需要再補別的檢查。
- [ ] 12. 收尾期間沒有跳出「Save changes?」對話框（`doc -s;` 有生效）。
- [ ] 13. `ORIGIN_PRO_MCP_REAP_CLOSE=1` 時的行為和之前相同（先存 recovery 檔再關）。
