# ✈️ 台北 (TPE) ⇋ 洛杉磯 (LAX) 機票價格智慧監控儀表板

專為王妍鈞（Brittany，就讀柯爾本音樂學院）往返洛杉磯與台北的直飛航線打造。
自動監控 **長榮航空 (EVA Air)**、**中華航空 (China Airlines)** 與 **星宇航空 (STARLUX Airlines)** 的真實直飛票價，內建**互動式網頁儀表板**、**LINE 自動降價推播通知**、**前後彈性日期比價 (±1~3天)** 與 **Excel/CSV 雙備份歷史資料庫**。

---

## 📂 專案檔案結構清單 (路徑：`E:\Antigravity_Data\flight-price-tracker\`)

| 檔案名稱 | 功能說明 |
| :--- | :--- |
| **`dashboard.html`** | **互動式網頁儀表板**。取代複雜的 Excel 試算表，包含即時最低價卡片、歷史走勢折線圖、航司比價長條圖與支援搜尋/篩選的完整動態報表。 |
| **`open_dashboard.bat`** | **一鍵開啟網頁儀表板**。雙擊即可直接在 Chrome / Edge 中打開 `dashboard.html` 檢視數據。 |
| **`run_monitor.bat`** | **一鍵執行爬蟲**。點擊後自動抓取最新票價、寫入 Excel/CSV 並**自動刷新同步 `dashboard.html`**。 |
| **`flight_tracker.py`** | 核心監控與分析引擎。支援來回/單程解析、LINE 推播、彈性日期比價與儀表板自動生成。 |
| **`dashboard_generator.py`** | 專責將 CSV/Excel 轉換為純前端無外部相依之美觀繁體中文網頁。 |
| **`config.json`** | 航線、日期、目標航司、直飛篩選、降價門檻、LINE Token 與彈性比價設定檔。 |
| **`price_history.xlsx`** | 自動生成的 Excel 歷史價格記錄表。 |
| **`price_history.csv`** | 自動生成的 CSV 歷史價格資料檔（UTF-8-SIG 編碼）。 |
| **`.venv/`** | 專屬獨立 Python 虛擬環境（已預載 `fast-flights`, `pandas`, `openpyxl` 等套件）。 |

---

## 🌐 如何使用「網頁儀表板」檢視數據？（推薦首選！）

您現在完全不需要手動開啟龐雜的 Excel 試算表：

1. 前往專案資料夾 `E:\Antigravity_Data\flight-price-tracker\`。
2. 對著 **`open_dashboard.bat`** 點擊滑鼠左鍵兩下。
3. 瀏覽器（Chrome 或 Edge）將立即開啟 **`dashboard.html`**：
   - 📊 **即時數據卡片**：一眼看清長榮、華航、星宇最新直飛價格與最低價。
   - 📈 **視覺化走勢圖**：SVG 向量繪製歷史票價波動折線與航司對比條形圖。
   - 🔍 **互動式篩選表**：支援依航空公司、航程類型篩選，輸入關鍵字即時搜尋機型與班次時間，並可一鍵匯出篩選後 CSV。

---

## ⚙️ 參數設定教學 (`config.json`)

使用「記事本」或任何文字編輯器打開 `config.json` 即可自訂設定：

```json
{
  "search_settings": {
    "from_airport": "LAX",              // 出發機場代碼 (LAX: 洛杉磯, TPE: 桃園)
    "to_airport": "TPE",                // 抵達機場代碼
    "depart_date": "2026-12-15",        // 去程出發日期 (YYYY-MM-DD)
    "return_date": "2027-01-10",        // 回程出發日期
    "trip_type": "round-trip",          // "round-trip" (來回) 或 "one-way" (單程)
    "seat_class": "economy",            // "economy" (經濟艙), "premium-economy" (豪經), "business" (商務)
    "direct_only": true,                // true = 僅直飛 (推薦)
    "currency": "TWD",                  // 計價幣別
    "target_airlines": [
      "長榮航空", "中華航空", "星宇航空"
    ]
  },
  "alert_settings": {
    "price_alert_threshold": 65000,     // 降價警報門檻 (低於此金額觸發警報與 LINE 通知)
    "notify_on_price_drop": true
  },
  "line_settings": {
    "enabled": false,                   // 改為 true 即可開啟 LINE 自動推播
    "channel_access_token": "YOUR_LINE_TOKEN",
    "user_id": "YOUR_USER_ID"
  },
  "flexible_date_settings": {
    "enabled": false,                   // 改為 true 每次查詢額外比對前後 1~2 天
    "range_days": 1
  },
  "schedule_settings": {
    "auto_loop_interval_hours": 0       // 0 為單次執行；若設為 6 或 12 則常駐定時重跑
  }
}
```

---

## 🔔 如何啟用 LINE 自動降價推播通知？

1. 至 [LINE Developers Console](https://developers.line.biz/) 免費註冊個人 Provider 並建立 **Messaging API Channel**。
2. 在 Channel 頁面獲取 **Channel Access Token (Long-lived)**。
3. 加入自己官方帳號為好友，並在 Basic Settings 找到您的 **Your user ID**（以 `U...` 開頭之字串）。
4. 將兩組字串填入 `config.json` 的 `line_settings`，並將 `"enabled": true`。
5. 執行終端測試：
   ```cmd
   .venv\Scripts\python.exe flight_tracker.py --test-line
   ```
   若手機收到「連線測試成功」，即代表日後只要偵測到降價或低於門檻，手機將自動響起推播！

---

## 🚀 常用指令與操作模式

- **正常更新票價並刷新網頁**：雙擊 `run_monitor.bat`
- **僅重新生成網頁儀表板**：
  ```cmd
  .venv\Scripts\python.exe flight_tracker.py --dashboard
  ```
- **執行前後彈性日期比較 (±1~2天比價)**：
  ```cmd
  .venv\Scripts\python.exe flight_tracker.py --flex
  ```
- **設定 Windows 每日自動背景排程**：
  開啟「工作排程器 (taskschd.msc)」，設定每日清晨 01:00 自動執行 `run_monitor.bat` 即可全自動累積票價趨勢。
