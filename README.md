# CE Tool

CE Tool 是一個桌面應用程式（Flask + pywebview），用來協助 CE/BOM 的檢查與註記（review / import / update）。

> 主要功能：對 BOM 檔案做料號比對、根據資料庫(mapping)補 comment、支援替料判斷與批次更新，並可將結果以 Excel 輸出。

---

## 📥 下載軟體

建議先下載我們打包好的版本，直接點擊即可使用：

💾 下載 [CE_Tool_Installer\_\*.exe](https://github.com/Andywang-95/new_cetool/releases/)

點擊執行 **CE_Tool_Installer\_\*.exe** 即可開始安裝。

---

## 🖥 使用說明（使用者）

- 啟動：下載並解壓後，執行 `run.exe`（Windows）或對應的可執行檔。
- 預設應用會啟動一個內部的 Flask server 並以 webview 顯示前端介面。

主畫面常見操作：

- **Select BOM**：選擇要檢查的 BOM Excel 檔。
- **Save Settings**：設定並儲存資料庫路徑等設定（會寫入設定檔、也更新 `app.config`）。
- **Start Review**（run_review）：會根據 BOM 預設的 `Action` 欄（`Add`、`Add Substitute`）逐列比對 mapping 資料庫，為主料填入 comment，替料若與主料相同則填「同上」，否則填原 comment。
  - import 模式:
  - _TBD_
- **Start Import**（run_import）：將外部資料匯入到專案需要的格式或資料庫（請依介面提示操作）。
- **Start Update**（run_update）：將已處理資料回寫或更新至資料來源。
- **Logs 顯示**：畫面下方會顯示處理過程的 log 與錯誤訊息。

重要檔案與格式：

- 資料庫（Database）為一個資料夾，內含 `mapping.xlsx`（料號 → comment）以及 `maintain.xlsx` 等檔案。
- 在執行前，請先於 Settings 指定正確的 `database_path`（程式會檢查 mapping.xlsx / maintain.xlsx 是否存在，並確認是否有 lock 檔 `~$...` 表示檔案被開啟）。

操作順序建議：

1. 先在 Settings 指定並儲存 `database_path`。
2. 選擇要處理的 BOM（Select BOM）。
3. 點選 `Start Review`（或 `Start Import` 依需求）。
4. 檢視下方 Logs，確認處理結果。

---

## 🛠 開發環境安裝

1. 安裝 Python
   - 推薦版本：Python 3.11
   - [Python 官方下載](https://www.python.org/downloads/)
2. 安裝套件管理工具
   - 使用 uv 管理 Python 套件 ([uv 安裝說明](https://docs.astral.sh/uv/getting-started/installation/#__tabbed_1_2))
   - 專案 clone 之後先執行：
     ```
     uv sync
     ```
3. 前端套件安裝
   - 使用 npm 管理前端套件 ([Node.js 安裝說明](https://nodejs.org/zh-tw/download))
     ```
     npm i
     ```
4. 前端資源生成
   - 生成靜態資源：
     ```
     npm run build
     ```
   - 或開啟開發模式監聽靜態檔案：
     ```
     npm run dev
     ```
5. 啟動開發環境
   ```
   uv run run.py
   ```
   - Flask 會在本地 `http://127.0.0.1:5001` 啟動
   - pywebview 會自動打開桌面視窗

---

## 開發與維護提示

- 若要修改 review 邏輯，建議把純資料處理改成用 `pandas`（速度與可讀性較好），再用 `openpyxl` 處理 Excel 樣式（底色、字體）。
- 共用函式請放在 `app/services` 或以語意命名（例如 `excel_utils.py`、`utils.py`），避免使用不直觀的檔名。

---

## Contributing

- 歡迎提交 Issue 或 Pull Request。請在 PR 描述中說明修改目的與測試方式。

---
