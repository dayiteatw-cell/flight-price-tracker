# -*- coding: utf-8 -*-
"""
機票數據網頁儀表板生成器 (Flight Dashboard Generator)
將 price_history.csv / price_history.xlsx 的數據轉化為高質感、互動式、具備視覺化圖表的繁體中文網頁。
支援獨立瀏覽器開啟 (Chrome / Edge) 與 Antigravity 視窗預覽。
"""

import os
import sys
import json

# Windows Console UTF-8 相容性處理
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import pandas as pd
from datetime import datetime

# 標準化歷史資料庫
def clean_and_load_data(csv_path):
    if not os.path.exists(csv_path):
        return pd.DataFrame()
    try:
        df = pd.read_csv(csv_path, encoding="utf-8-sig")
    except Exception:
        return pd.DataFrame()
    
    records = []
    for _, r in df.iterrows():
        # 兼容不同版本的欄位命名
        airline = str(r.get("航空公司", "")) if pd.notna(r.get("航空公司")) else ""
        if not airline or airline == "nan":
            continue
            
        raw_price = r.get("票價(TWD)")
        try:
            price = int(float(raw_price)) if pd.notna(raw_price) else 0
        except Exception:
            price = 0
            
        query_time = str(r.get("查詢時間", "")) if pd.notna(r.get("查詢時間")) else ""
        
        # 航程類型
        trip_type = str(r.get("航程類型", "")) if pd.notna(r.get("航程類型")) else ""
        if not trip_type or trip_type == "nan":
            if pd.notna(r.get("回程日期")):
                trip_type = "來回套票 (直飛)"
            else:
                trip_type = "單程機票"
                
        # 航班日期
        flight_date = str(r.get("航班日期", "")) if pd.notna(r.get("航班日期")) else ""
        if not flight_date or flight_date == "nan":
            dep = str(r.get("去程日期", "")) if pd.notna(r.get("去程日期")) else ""
            ret = str(r.get("回程日期", "")) if pd.notna(r.get("回程日期")) else ""
            if dep and ret:
                flight_date = f"{dep} 往返 {ret}"
            elif dep:
                flight_date = dep
            else:
                flight_date = "未載日期"
                
        # 航班時間與機型
        itinerary = str(r.get("航班時間與機型", "")) if pd.notna(r.get("航班時間與機型")) else ""
        if not itinerary or itinerary == "nan":
            itinerary = str(r.get("航程時間與機型", "")) if pd.notna(r.get("航程時間與機型")) else "機型班次未載"
            
        date_type = str(r.get("日期類別", "")) if pd.notna(r.get("日期類別")) else ""
        if not date_type or date_type == "nan":
            date_type = "🎯基準目標"

        from_ap = str(r.get("出發地", "LAX")) if pd.notna(r.get("出發地")) else "LAX"
        to_ap = str(r.get("目的地", "TPE")) if pd.notna(r.get("目的地")) else "TPE"

        records.append({
            "query_time": query_time,
            "date_type": date_type,
            "trip_type": trip_type,
            "from_ap": from_ap,
            "to_ap": to_ap,
            "flight_date": flight_date,
            "airline": airline,
            "itinerary": itinerary,
            "price": price
        })
        
    return pd.DataFrame(records)

def generate_dashboard_html(csv_path, config_path, output_path):
    df = clean_and_load_data(csv_path)
    
    config = {}
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
        except Exception:
            pass

    records_json = df.to_json(orient="records", force_ascii=False) if not df.empty else "[]"
    config_json = json.dumps(config, ensure_ascii=False)

    html_content = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>洛杉磯 ⇋ 台北 機票價格智慧監控儀表板</title>
  <script src="https://www.gstatic.com/antigravity/web/dev/tailwindcss.min.js"></script>
  <style>
    :root {{
      --primary-eva: #00873e;
      --primary-ci: #8b1e4f;
      --primary-jx: #9d7c38;
    }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, "Noto Sans TC", sans-serif;
    }}
    .custom-scroll::-webkit-scrollbar {{
      width: 6px;
      height: 6px;
    }}
    .custom-scroll::-webkit-scrollbar-track {{
      background: rgba(0,0,0,0.05);
      border-radius: 4px;
    }}
    .custom-scroll::-webkit-scrollbar-thumb {{
      background: rgba(100,116,139,0.3);
      border-radius: 4px;
    }}
    .custom-scroll::-webkit-scrollbar-thumb:hover {{
      background: rgba(100,116,139,0.5);
    }}
  </style>
</head>
<body class="bg-slate-50 text-slate-800 antialiased p-4 md:p-8 min-h-screen">
  <div class="max-w-7xl mx-auto space-y-6">

    <!-- 頂部頁首與導航 -->
    <header class="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
      <div class="space-y-1">
        <div class="flex items-center gap-2">
          <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold bg-emerald-100 text-emerald-800">
            即時監控中
          </span>
          <span class="text-xs text-slate-500" id="header-last-updated">更新時間：載入中...</span>
        </div>
        <h1 class="text-2xl md:text-3xl font-extrabold text-slate-900 tracking-tight flex items-center gap-2">
          ✈️ 洛杉磯 (LAX) ⇋ 台北 (TPE) 機票價格監控儀表板
        </h1>
        <p class="text-sm text-slate-600">
          專為王妍鈞（Brittany，柯爾本音樂學院）往返台美直飛航班打造・涵蓋長榮 (BR)、華航 (CI)、星宇 (JX)
        </p>
      </div>
      <div class="flex flex-wrap items-center gap-2">
        <button onclick="location.reload()" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-medium rounded-xl transition flex items-center gap-1.5 shadow-sm">
          🔄 重新整理網頁
        </button>
        <button onclick="exportFilteredCSV()" class="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium rounded-xl transition flex items-center gap-1.5 shadow-sm">
          📥 匯出當前視圖 CSV
        </button>
      </div>
    </header>

    <!-- 精選超值航班情報推薦橫幅 -->
    <div id="insight-banner" class="bg-gradient-to-r from-amber-500/10 via-emerald-500/10 to-teal-500/10 border border-amber-300/40 rounded-2xl p-4 md:p-5 flex flex-col md:flex-row md:items-center justify-between gap-3 shadow-sm">
      <div class="flex items-start gap-3">
        <div class="w-10 h-10 rounded-xl bg-amber-500/20 text-amber-600 flex items-center justify-center font-bold text-xl flex-shrink-0">
          💡
        </div>
        <div>
          <div class="font-bold text-slate-900 flex items-center gap-2">
            <span id="banner-title">最佳彈性日期省錢洞察</span>
            <span id="banner-badge" class="bg-amber-100 text-amber-800 text-xs px-2 py-0.5 rounded-full font-semibold">分析中</span>
          </div>
          <p id="banner-desc" class="text-xs md:text-sm text-slate-600 mt-0.5">
            正在分析前後 ±1~2 天航班比價數據...
          </p>
        </div>
      </div>
      <a href="https://www.evaair.com/" target="_blank" class="px-4 py-2 bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold rounded-xl text-center whitespace-nowrap transition shadow-sm">
        前往長榮官網 ↗
      </a>
    </div>

    <!-- 核心指標統計卡片 -->
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
      <!-- 最低總價卡 -->
      <div class="bg-white border border-slate-200 rounded-2xl p-5 shadow-sm flex flex-col justify-between hover:border-emerald-300 transition">
        <div class="flex items-center justify-between text-slate-500 text-xs font-medium mb-2">
          <span>🎯 目標來回最低直飛</span>
          <span class="p-1 rounded-lg bg-emerald-50 text-emerald-600 font-bold">全場最低</span>
        </div>
        <div>
          <div class="text-2xl md:text-3xl font-black text-emerald-600" id="stat-min-rt">NT$ --</div>
          <div class="text-xs text-slate-500 mt-1" id="stat-min-rt-airline">--</div>
        </div>
        <div class="mt-3 pt-3 border-t border-slate-100 text-xs text-slate-600 flex justify-between">
          <span>警報門檻</span>
          <span class="font-semibold text-slate-700" id="stat-threshold">NT$ 65,000</span>
        </div>
      </div>

      <!-- 長榮航空卡 -->
      <div class="bg-white border border-slate-200 rounded-2xl p-5 shadow-sm flex flex-col justify-between hover:border-emerald-400 transition">
        <div class="flex items-center justify-between text-slate-500 text-xs font-medium mb-2">
          <span class="flex items-center gap-1.5 font-bold text-emerald-700">
            <span class="w-2.5 h-2.5 rounded-full bg-emerald-600"></span> 長榮航空 (EVA Air)
          </span>
          <span class="text-xs text-slate-400">BR</span>
        </div>
        <div>
          <div class="text-2xl md:text-3xl font-black text-slate-900" id="stat-eva-price">NT$ --</div>
          <div class="text-xs text-slate-500 mt-1" id="stat-eva-flight">班次檢索中...</div>
        </div>
        <div class="mt-3 pt-3 border-t border-slate-100 text-xs text-slate-600 flex justify-between">
          <span>去程直飛單程</span>
          <span class="font-semibold text-slate-700" id="stat-eva-one-way">NT$ --</span>
        </div>
      </div>

      <!-- 中華航空卡 -->
      <div class="bg-white border border-slate-200 rounded-2xl p-5 shadow-sm flex flex-col justify-between hover:border-rose-400 transition">
        <div class="flex items-center justify-between text-slate-500 text-xs font-medium mb-2">
          <span class="flex items-center gap-1.5 font-bold text-rose-700">
            <span class="w-2.5 h-2.5 rounded-full bg-rose-600"></span> 中華航空 (China Airlines)
          </span>
          <span class="text-xs text-slate-400">CI</span>
        </div>
        <div>
          <div class="text-2xl md:text-3xl font-black text-slate-900" id="stat-ci-price">NT$ --</div>
          <div class="text-xs text-slate-500 mt-1" id="stat-ci-flight">班次檢索中...</div>
        </div>
        <div class="mt-3 pt-3 border-t border-slate-100 text-xs text-slate-600 flex justify-between">
          <span>去程直飛單程</span>
          <span class="font-semibold text-slate-700" id="stat-ci-one-way">NT$ --</span>
        </div>
      </div>

      <!-- 星宇航空卡 -->
      <div class="bg-white border border-slate-200 rounded-2xl p-5 shadow-sm flex flex-col justify-between hover:border-amber-400 transition">
        <div class="flex items-center justify-between text-slate-500 text-xs font-medium mb-2">
          <span class="flex items-center gap-1.5 font-bold text-amber-700">
            <span class="w-2.5 h-2.5 rounded-full bg-amber-600"></span> 星宇航空 (STARLUX)
          </span>
          <span class="text-xs text-slate-400">JX</span>
        </div>
        <div>
          <div class="text-2xl md:text-3xl font-black text-slate-900" id="stat-jx-price">NT$ --</div>
          <div class="text-xs text-slate-500 mt-1" id="stat-jx-flight">班次檢索中...</div>
        </div>
        <div class="mt-3 pt-3 border-t border-slate-100 text-xs text-slate-600 flex justify-between">
          <span>去程直飛單程</span>
          <span class="font-semibold text-slate-700" id="stat-jx-one-way">NT$ --</span>
        </div>
      </div>
    </div>

    <!-- 視覺化趨勢圖表區塊 (純 SVG 前端繪製，無外部 CDN 相依) -->
    <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
      <!-- 歷史價格趨勢折線圖 -->
      <div class="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm lg:col-span-2">
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-4">
          <div>
            <h2 class="text-lg font-bold text-slate-900 flex items-center gap-2">
              📈 歷史票價趨勢走勢圖
            </h2>
            <p class="text-xs text-slate-500">直飛來回套票最低價變動追蹤（依查詢時間紀錄）</p>
          </div>
          <div class="flex items-center gap-3 text-xs font-medium">
            <span class="flex items-center gap-1.5 text-emerald-700"><span class="w-3 h-1 bg-emerald-500 rounded"></span>長榮</span>
            <span class="flex items-center gap-1.5 text-rose-700"><span class="w-3 h-1 bg-rose-500 rounded"></span>華航</span>
            <span class="flex items-center gap-1.5 text-amber-700"><span class="w-3 h-1 bg-amber-500 rounded"></span>星宇</span>
          </div>
        </div>
        <div id="trend-chart-container" class="w-full h-64 relative bg-slate-50/50 rounded-xl border border-slate-100 p-2 overflow-hidden flex items-center justify-center">
          <!-- 由 JS 繪製 SVG 折線圖 -->
        </div>
      </div>

      <!-- 航司目前價格對比長條圖 -->
      <div class="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm flex flex-col justify-between">
        <div>
          <h2 class="text-lg font-bold text-slate-900 flex items-center gap-2 mb-1">
            📊 當前直飛最低票價對比
          </h2>
          <p id="bar-chart-subtitle" class="text-xs text-slate-500 mb-4">以 2026-12-15 往返 2027-01-10 直飛來回為基準</p>
          <div id="bar-chart-container" class="space-y-4">
            <!-- 由 JS 渲染橫向長條圖 -->
          </div>
        </div>
        <div class="p-3 bg-slate-50 rounded-xl border border-slate-100 text-xs text-slate-600 mt-4 space-y-1">
          <div class="font-semibold text-slate-700">💡 航情重點筆記：</div>
          <div>• <strong>長榮航空</strong>：目前具備最佳來回套票優惠，比分開單程買省下逾萬元。</div>
          <div>• <strong>中華航空</strong>：回程 1/10 經濟艙常態客滿，需提早鎖定釋票。</div>
        </div>
      </div>
    </div>

    <!-- 歷史詳細資料表 (取代 Excel) -->
    <div class="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
      <!-- 篩選控制列 -->
      <div class="p-5 border-b border-slate-100 bg-slate-50/50 flex flex-col lg:flex-row lg:items-center justify-between gap-4">
        <div>
          <h2 class="text-lg font-bold text-slate-900 flex items-center gap-2">
            📋 完整歷史報表資料庫
            <span class="text-xs font-semibold px-2 py-0.5 rounded-full bg-slate-200 text-slate-700" id="table-row-count">0 筆</span>
          </h2>
          <p class="text-xs text-slate-500">所有由 Python 爬蟲抓取的完整紀錄（即時篩選、搜尋與排序）</p>
        </div>

        <!-- 篩選按鈕與搜尋 -->
        <div class="flex flex-wrap items-center gap-2">
          <!-- 日期類別下拉 -->
          <select id="filter-date-type" onchange="applyFilters()" class="text-xs border border-slate-200 rounded-xl px-3 py-2 bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-emerald-500">
            <option value="ALL">全部日期（基準＋彈性）</option>
            <option value="🎯基準目標">🎯 僅看基準目標日期</option>
            <option value="💡彈性比價">💡 僅看前後彈性比價</option>
          </select>

          <!-- 航司下拉選單 -->
          <select id="filter-airline" onchange="applyFilters()" class="text-xs border border-slate-200 rounded-xl px-3 py-2 bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-emerald-500">
            <option value="ALL">全部航空公司</option>
            <option value="長榮航空">長榮航空</option>
            <option value="中華航空">中華航空</option>
            <option value="星宇航空">星宇航空</option>
          </select>

          <!-- 航程類型下拉 -->
          <select id="filter-type" onchange="applyFilters()" class="text-xs border border-slate-200 rounded-xl px-3 py-2 bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-emerald-500">
            <option value="ALL">全部航程類型</option>
            <option value="來回套票 (直飛)">來回套票 (直飛)</option>
            <option value="去程單程 (直飛)">去程單程 (直飛)</option>
            <option value="回程單程 (直飛)">回程單程 (直飛)</option>
          </select>

          <!-- 排序 -->
          <select id="filter-sort" onchange="applyFilters()" class="text-xs border border-slate-200 rounded-xl px-3 py-2 bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-emerald-500">
            <option value="price-asc">票價：由低到高</option>
            <option value="price-desc">票價：由高到低</option>
            <option value="time-desc" selected>查詢時間：最新優先</option>
          </select>

          <!-- 搜尋框 -->
          <div class="relative">
            <input type="text" id="filter-search" oninput="applyFilters()" placeholder="搜尋班次、機型、日期..." class="text-xs border border-slate-200 rounded-xl pl-8 pr-3 py-2 bg-white text-slate-700 w-44 md:w-56 focus:outline-none focus:ring-2 focus:ring-emerald-500">
            <span class="absolute left-2.5 top-2.5 text-slate-400 text-xs">🔍</span>
          </div>
        </div>
      </div>

      <!-- 表格本體 -->
      <div class="overflow-x-auto custom-scroll">
        <table class="w-full text-left border-collapse text-xs md:text-sm">
          <thead>
            <tr class="bg-slate-100/70 border-b border-slate-200 text-slate-600 font-semibold uppercase tracking-wider text-[11px] md:text-xs">
              <th class="py-3.5 px-4">查詢時間</th>
              <th class="py-3.5 px-4">航程類型</th>
              <th class="py-3.5 px-4">航線與航班日期</th>
              <th class="py-3.5 px-4">航空公司</th>
              <th class="py-3.5 px-4">起降時間與機型</th>
              <th class="py-3.5 px-4 text-right">票價 (TWD)</th>
            </tr>
          </thead>
          <tbody id="table-body" class="divide-y divide-slate-100 text-slate-700">
            <!-- 由 JS 渲染動態表格內容 -->
          </tbody>
        </table>
      </div>

      <!-- 頁籤與提示 -->
      <div class="p-4 border-t border-slate-100 bg-slate-50 flex items-center justify-between text-xs text-slate-500">
        <span id="pagination-info">正在顯示 1 到 0 筆</span>
        <div class="flex items-center gap-1">
          <button id="btn-prev" onclick="changePage(-1)" class="px-3 py-1 bg-white border border-slate-200 rounded-lg text-slate-600 hover:bg-slate-100 disabled:opacity-40">上一頁</button>
          <span id="page-num" class="px-2 font-medium">1 / 1</span>
          <button id="btn-next" onclick="changePage(1)" class="px-3 py-1 bg-white border border-slate-200 rounded-lg text-slate-600 hover:bg-slate-100 disabled:opacity-40">下一頁</button>
        </div>
      </div>
    </div>

    <!-- 底部功能模組：LINE 通知與自動設定導引 -->
    <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
      <!-- LINE 推播設定狀態卡 -->
      <div class="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm space-y-4">
        <div class="flex items-center justify-between">
          <h3 class="text-base font-bold text-slate-900 flex items-center gap-2">
            🔔 LINE 自動降價推播通知
          </h3>
          <span class="text-xs px-2.5 py-0.5 rounded-full bg-slate-100 text-slate-700 font-medium" id="line-status-badge">
            已就緒 (可於 config 設定 Token)
          </span>
        </div>
        <p class="text-xs text-slate-600">
          當爬蟲在每日凌晨或定期輪詢時，發現直飛票價低於設定門檻（例如 NT$ 65,000）或有大幅降價，將第一時間直接發送 LINE 訊息到您的手機！
        </p>
        <div class="p-3 bg-slate-50 rounded-xl border border-slate-100 text-xs space-y-1.5 text-slate-600">
          <div class="font-semibold text-slate-800">📱 啟用步驟：</div>
          <div>1. 在 LINE Developers 申請免費 Messaging API 頻道。</div>
          <div>2. 將取得的 <code class="bg-slate-200 px-1 py-0.5 rounded">Channel Access Token</code> 與 <code class="bg-slate-200 px-1 py-0.5 rounded">User ID</code> 填入 <code class="bg-slate-200 px-1 py-0.5 rounded">config.json</code>。</div>
          <div>3. 執行 <code class="bg-slate-200 px-1 py-0.5 rounded">run_monitor.bat</code> 即可在手機立即收到降價通知！</div>
        </div>
      </div>

      <!-- 快速執行與檔案連結指南 -->
      <div class="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm space-y-4">
        <h3 class="text-base font-bold text-slate-900 flex items-center gap-2">
          ⚡ 快速啟動與檔案清單
        </h3>
        <p class="text-xs text-slate-600">
          本專案所有腳本與資料庫均持久化儲存於本機資料夾，支援無伺服器直接開啟：
        </p>
        <ul class="text-xs text-slate-600 space-y-2">
          <li class="flex items-center gap-2">
            <span class="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
            <strong>專案目錄</strong>：<span class="font-mono bg-slate-100 px-1.5 py-0.5 rounded text-slate-700">E:\\\\Antigravity_Data\\\\flight-price-tracker\\\\</span>
          </li>
          <li class="flex items-center gap-2">
            <span class="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
            <strong>一鍵爬取</strong>：點擊 <span class="font-mono bg-slate-100 px-1.5 py-0.5 rounded text-slate-700">run_monitor.bat</span> 即可更新最新票價
          </li>
          <li class="flex items-center gap-2">
            <span class="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
            <strong>歷史檔案</strong>：支援 Excel (<span class="font-mono">.xlsx</span>) 與 CSV 雙備份同步更新
          </li>
        </ul>
      </div>
    </div>

  </div>

  <!-- 資料處理與前端互動 JavaScript -->
  <script>
    // 注入後台數據
    const RAW_DATA = {records_json};
    const CONFIG = {config_json};

    let filteredData = [...RAW_DATA];
    let currentPage = 1;
    const pageSize = 15;

    // 初始化頁面
    window.addEventListener("DOMContentLoaded", () => {{
      renderSummaryCards();
      renderBarChart();
      renderTrendChart();
      applyFilters();
    }});

    // 格式化千分位
    function formatMoney(num) {{
      if (!num || isNaN(num)) return "未載";
      return "NT$ " + Number(num).toLocaleString();
    }}

    // 渲染指標卡片
    function renderSummaryCards() {{
      if (!RAW_DATA || RAW_DATA.length === 0) return;

      const latestTime = RAW_DATA[RAW_DATA.length - 1].query_time;
      document.getElementById("header-last-updated").textContent = `最新查詢時間：${{latestTime}}`;

      // 設定動態副標
      const depDate = (CONFIG.flight_dates && CONFIG.flight_dates.base_target) ? CONFIG.flight_dates.base_target.departure_date : "2026-12-15";
      const retDate = (CONFIG.flight_dates && CONFIG.flight_dates.base_target) ? CONFIG.flight_dates.base_target.return_date : "2027-01-10";
      const subTitle = document.getElementById("bar-chart-subtitle");
      if (subTitle) {{
        subTitle.textContent = `以 ${{depDate}} 往返 ${{retDate}} 直飛來回為基準`;
      }}

      // 最新一批查詢記錄
      const latestRecords = RAW_DATA.filter(d => d.query_time === latestTime);

      // 基準目標來回航班 (排除彈性比價)
      const targetRTs = latestRecords.filter(d => (d.date_type === "🎯基準目標" || !d.date_type.includes("彈性")) && d.trip_type && d.trip_type.includes("來回") && d.price > 0).sort((a, b) => a.price - b.price);
      
      // 彈性比價來回航班
      const flexRTs = latestRecords.filter(d => (d.date_type === "💡彈性比價" || d.date_type.includes("彈性")) && d.trip_type && d.trip_type.includes("來回") && d.price > 0).sort((a, b) => a.price - b.price);

      // 全場基準目標最低來回
      if (targetRTs.length > 0) {{
        const lowest = targetRTs[0];
        document.getElementById("stat-min-rt").textContent = formatMoney(lowest.price);
        document.getElementById("stat-min-rt-airline").textContent = `${{lowest.airline}} (基準 ${{lowest.flight_date}})`;
      }} else if (latestRecords.length > 0) {{
        const anyRT = latestRecords.filter(d => d.trip_type && d.trip_type.includes("來回") && d.price > 0).sort((a, b) => a.price - b.price);
        if (anyRT.length > 0) {{
          document.getElementById("stat-min-rt").textContent = formatMoney(anyRT[0].price);
          document.getElementById("stat-min-rt-airline").textContent = `${{anyRT[0].airline}} (${{anyRT[0].flight_date}})`;
        }}
      }}

      // 設定門檻
      if (CONFIG.alert_settings && CONFIG.alert_settings.price_alert_threshold) {{
        document.getElementById("stat-threshold").textContent = formatMoney(CONFIG.alert_settings.price_alert_threshold);
      }}

      // LINE 推播狀態
      if (CONFIG.line_settings && CONFIG.line_settings.enabled) {{
        const badge = document.getElementById("line-status-badge");
        if (badge) {{
          badge.textContent = "✅ 已成功連線啟用";
          badge.className = "text-xs px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-800 font-semibold";
        }}
      }}

      // 頂部彈性省錢洞察動態更新 (Banner)
      updateInsightBanner(targetRTs, flexRTs);

      // 各航司最新即時數據 (以最新批次中的基準目標為準，取最低票價)
      updateAirlineStat("長榮", "stat-eva-price", "stat-eva-flight", "stat-eva-one-way", latestRecords);
      updateAirlineStat("中華", "stat-ci-price", "stat-ci-flight", "stat-ci-one-way", latestRecords);
      updateAirlineStat("星宇", "stat-jx-price", "stat-jx-flight", "stat-jx-one-way", latestRecords);
    }}

    // 動態更新頂部推薦情報
    function updateInsightBanner(targetRTs, flexRTs) {{
      const bTitle = document.getElementById("banner-title");
      const bBadge = document.getElementById("banner-badge");
      const bDesc = document.getElementById("banner-desc");
      if (!bTitle || !bBadge || !bDesc) return;

      const baseLowest = targetRTs && targetRTs.length > 0 ? targetRTs[0] : null;
      const flexLowest = flexRTs && flexRTs.length > 0 ? flexRTs[0] : null;

      if (baseLowest && flexLowest && flexLowest.price < baseLowest.price) {{
        const saved = baseLowest.price - flexLowest.price;
        bTitle.textContent = `最佳彈性日期省錢洞察：${{flexLowest.flight_date}} 大特惠`;
        bBadge.textContent = `現省 NT$ ${{saved.toLocaleString()}}`;
        bDesc.innerHTML = `原基準目標（${{baseLowest.flight_date}}）最低為 <strong>${{formatMoney(baseLowest.price)}}</strong>（${{baseLowest.airline}}），若彈性調整為 <strong>${{flexLowest.flight_date}}</strong>，${{flexLowest.airline}}直飛來回總價只要 <strong>${{formatMoney(flexLowest.price)}}</strong>！班次：${{flexLowest.itinerary}}`;
      }} else if (baseLowest) {{
        bTitle.textContent = `航情重點洞察：基準目標即為目前最低票價！`;
        bBadge.textContent = `最佳買點`;
        bDesc.innerHTML = `目前基準目標（<strong>${{baseLowest.flight_date}}</strong>）的 <strong>${{formatMoney(baseLowest.price)}}</strong>（${{baseLowest.airline}}）為當前台美直飛最優票價。`;
      }}
    }}

    function updateAirlineStat(keyword, priceId, flightId, oneWayId, latestRecords) {{
      const records = latestRecords || RAW_DATA;
      // 基準目標來回：依票價升序排序，取最低者
      const targetRTs = records.filter(d => d.airline && d.airline.includes(keyword) && (d.date_type === "🎯基準目標" || !d.date_type.includes("彈性")) && d.trip_type && d.trip_type.includes("來回") && d.price > 0).sort((a, b) => a.price - b.price);

      if (targetRTs.length > 0) {{
        const best = targetRTs[0];
        document.getElementById(priceId).textContent = formatMoney(best.price);
        document.getElementById(flightId).textContent = best.itinerary ? best.itinerary.split("｜")[0] : best.flight_date;
      }} else {{
        // 若基準目標售罄，查詢彈性日
        const anyRTs = records.filter(d => d.airline && d.airline.includes(keyword) && d.trip_type && d.trip_type.includes("來回") && d.price > 0).sort((a, b) => a.price - b.price);
        if (anyRTs.length > 0) {{
          document.getElementById(priceId).textContent = formatMoney(anyRTs[0].price);
          document.getElementById(flightId).textContent = `(彈性) ${{anyRTs[0].flight_date}}`;
        }} else {{
          document.getElementById(priceId).textContent = "來回客滿";
          document.getElementById(flightId).textContent = "直飛無空位";
        }}
      }}

      // 去程單程：取最低者
      const outList = records.filter(d => d.airline && d.airline.includes(keyword) && d.trip_type && d.trip_type.includes("去程") && d.price > 0).sort((a, b) => a.price - b.price);
      if (outList.length > 0) {{
        document.getElementById(oneWayId).textContent = formatMoney(outList[0].price);
      }} else {{
        document.getElementById(oneWayId).textContent = "單程售罄";
      }}
    }}

    // 渲染各航司橫向長條圖 (以最新一批查詢的基準目標最低價為準)
    function renderBarChart() {{
      const container = document.getElementById("bar-chart-container");
      if (!container) return;

      const latestTime = RAW_DATA && RAW_DATA.length > 0 ? RAW_DATA[RAW_DATA.length - 1].query_time : null;
      const latestRecords = latestTime ? RAW_DATA.filter(d => d.query_time === latestTime) : RAW_DATA;

      const airlines = [
        {{ name: "長榮航空 (EVA Air)", key: "長榮", color: "bg-emerald-600", border: "border-emerald-700" }},
        {{ name: "中華航空 (China Airlines)", key: "中華", color: "bg-rose-600", border: "border-rose-700" }},
        {{ name: "星宇航空 (STARLUX)", key: "星宇", color: "bg-amber-600", border: "border-amber-700" }}
      ];

      let html = "";
      const maxPrice = 80000;

      airlines.forEach(a => {{
        // 取得該航司在最新查詢中的基準目標最低來回票價
        const targetMatches = latestRecords.filter(d => d.airline && d.airline.includes(a.key) && (d.date_type === "🎯基準目標" || !d.date_type.includes("彈性")) && d.trip_type && d.trip_type.includes("來回") && d.price > 0).sort((x, y) => x.price - y.price);

        let price = targetMatches.length > 0 ? targetMatches[0].price : 0;
        let percent = price > 0 ? Math.min(100, Math.round((price / maxPrice) * 100)) : 0;
        let priceText = price > 0 ? formatMoney(price) : "來回客滿";

        // 檢查該航司是否有更平價的「彈性日期」推薦
        const flexMatches = latestRecords.filter(d => d.airline && d.airline.includes(a.key) && (d.date_type === "💡彈性比價" || d.date_type.includes("彈性")) && d.trip_type && d.trip_type.includes("來回") && d.price > 0).sort((x, y) => x.price - y.price);

        let flexTip = "";
        if (flexMatches.length > 0 && price > 0 && flexMatches[0].price < price) {{
          const saved = price - flexMatches[0].price;
          flexTip = `
            <div class="mt-1 flex items-center justify-between text-[11px] text-emerald-700 bg-emerald-50/80 px-2 py-0.5 rounded-md border border-emerald-100">
              <span>💡 彈性最平價：<strong>${{formatMoney(flexMatches[0].price)}}</strong> (${{flexMatches[0].flight_date.replace(" 往返 ", " ⇋ ")}})</span>
              <span class="font-bold text-emerald-800">現省 NT$ ${{saved.toLocaleString()}}</span>
            </div>
          `;
        }}

        html += `
          <div class="space-y-1.5 p-2 rounded-xl hover:bg-slate-50 transition border border-transparent hover:border-slate-100">
            <div class="flex justify-between text-xs font-semibold text-slate-700">
              <span class="flex items-center gap-1.5">
                <span class="w-2.5 h-2.5 rounded-full ${{a.color}}"></span>
                ${{a.name}}
              </span>
              <span class="text-slate-900 font-bold">${{priceText}}</span>
            </div>
            <div class="w-full bg-slate-100 rounded-full h-3 overflow-hidden flex">
              <div class="${{a.color}} h-3 rounded-full transition-all duration-500" style="width: ${{percent}}%"></div>
            </div>
            ${{flexTip}}
          </div>
        `;
      }});

      container.innerHTML = html;
    }}

    // 繪製純 SVG 歷史趨勢折線圖
    function renderTrendChart() {{
      const container = document.getElementById("trend-chart-container");
      if (!container) return;

      // 提取基準目標來回紀錄 (排除彈性日期干擾走勢)
      const rts = RAW_DATA.filter(d => (d.date_type === "🎯基準目標" || !d.date_type.includes("彈性")) && d.trip_type && d.trip_type.includes("來回") && d.price > 0);
      if (rts.length === 0) {{
        container.innerHTML = '<span class="text-xs text-slate-400">目前尚無足夠的基準目標來回數據繪製圖表</span>';
        return;
      }}

      // 按查詢時間分組
      const timePoints = [...new Set(rts.map(d => d.query_time))].sort();
      if (timePoints.length < 2) {{
        // 若只有一個時間點，提供友好呈現
        container.innerHTML = `
          <div class="text-center space-y-1">
            <div class="text-emerald-600 font-bold text-sm">已記錄首波數據 (${{timePoints[0]}})</div>
            <div class="text-xs text-slate-500">隨著每日排程執行，此處將自動繪出長榮、華航與星宇的波動折線！</div>
          </div>
        `;
        return;
      }}

      // 繪製 SVG 折線
      const w = 600;
      const h = 200;
      const padL = 60;
      const padR = 20;
      const padT = 20;
      const padB = 30;

      const minP = 60000;
      const maxP = 75000;

      function getX(idx) {{
        return padL + (idx / (timePoints.length - 1)) * (w - padL - padR);
      }}

      function getY(price) {{
        return padT + (1 - (price - minP) / (maxP - minP)) * (h - padT - padB);
      }}

      const svgLines = [];
      const airlines = [
        {{ key: "長榮", stroke: "#059669" }},
        {{ key: "中華", stroke: "#e11d48" }},
        {{ key: "星宇", stroke: "#d97706" }}
      ];

      airlines.forEach(a => {{
        const pts = [];
        timePoints.forEach((t, idx) => {{
          // 篩選該時間點的基準目標來回航班最低價
          const matches = rts.filter(d => d.query_time === t && d.airline && d.airline.includes(a.key)).sort((x, y) => x.price - y.price);
          if (matches.length > 0) {{
            const best = matches[0];
            pts.push({{ x: getX(idx), y: getY(best.price), p: best.price }});
          }}
        }});

        if (pts.length > 1) {{
          let pathD = `M ${{pts[0].x}} ${{pts[0].y}}`;
          for (let i = 1; i < pts.length; i++) {{
            pathD += ` L ${{pts[i].x}} ${{pts[i].y}}`;
          }}
          svgLines.push(`<path d="${{pathD}}" fill="none" stroke="${{a.stroke}}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>`);
          pts.forEach(p => {{
            svgLines.push(`<circle cx="${{p.x}}" cy="${{p.y}}" r="4" fill="${{a.stroke}}" stroke="#ffffff" stroke-width="1.5"/>`);
          }});
        }}
      }});

      container.innerHTML = `
        <svg viewBox="0 0 ${{w}} ${{h}}" class="w-full h-full">
          <!-- 刻度底線 -->
          <line x1="${{padL}}" y1="${{getY(minP)}}" x2="${{w - padR}}" y2="${{getY(minP)}}" stroke="#e2e8f0" stroke-dasharray="3,3"/>
          <text x="${{padL - 8}}" y="${{getY(minP) + 4}}" font-size="10" fill="#94a3b8" text-anchor="end">${{minP / 1000}}k</text>

          <line x1="${{padL}}" y1="${{getY(65000)}}" x2="${{w - padR}}" y2="${{getY(65000)}}" stroke="#e2e8f0" stroke-dasharray="3,3"/>
          <text x="${{padL - 8}}" y="${{getY(65000) + 4}}" font-size="10" fill="#94a3b8" text-anchor="end">65k</text>

          <line x1="${{padL}}" y1="${{getY(maxP)}}" x2="${{w - padR}}" y2="${{getY(maxP)}}" stroke="#e2e8f0" stroke-dasharray="3,3"/>
          <text x="${{padL - 8}}" y="${{getY(maxP) + 4}}" font-size="10" fill="#94a3b8" text-anchor="end">${{maxP / 1000}}k</text>

          <!-- 時間軸標記 -->
          <text x="${{padL}}" y="${{h - 10}}" font-size="10" fill="#94a3b8">${{timePoints[0].split(" ")[0]}}</text>
          <text x="${{w - padR}}" y="${{h - 10}}" font-size="10" fill="#94a3b8" text-anchor="end">${{timePoints[timePoints.length - 1].split(" ")[0]}}</text>

          ${{svgLines.join("")}}
        </svg>
      `;
    }}

    // 篩選與排序表格
    function applyFilters() {{
      const dateTypeFilter = document.getElementById("filter-date-type").value;
      const airlineFilter = document.getElementById("filter-airline").value;
      const typeFilter = document.getElementById("filter-type").value;
      const sortFilter = document.getElementById("filter-sort").value;
      const searchKey = document.getElementById("filter-search").value.trim().toLowerCase();

      filteredData = RAW_DATA.filter(row => {{
        if (dateTypeFilter !== "ALL" && row.date_type !== dateTypeFilter) {{
          return false;
        }}
        if (airlineFilter !== "ALL" && (!row.airline || !row.airline.includes(airlineFilter))) {{
          return false;
        }}
        if (typeFilter !== "ALL" && row.trip_type !== typeFilter) {{
          return false;
        }}
        if (searchKey) {{
          const combined = `${{row.date_type}} ${{row.airline}} ${{row.flight_date}} ${{row.itinerary}} ${{row.query_time}}`.toLowerCase();
          if (!combined.includes(searchKey)) return false;
        }}
        return true;
      }});

      // 排序
      if (sortFilter === "price-asc") {{
        filteredData.sort((a, b) => a.price - b.price);
      }} else if (sortFilter === "price-desc") {{
        filteredData.sort((a, b) => b.price - a.price);
      }} else if (sortFilter === "time-desc") {{
        filteredData.sort((a, b) => b.query_time.localeCompare(a.query_time));
      }}

      currentPage = 1;
      renderTable();
    }}

    // 渲染表格內容
    function renderTable() {{
      const tbody = document.getElementById("table-body");
      const rowCountSpan = document.getElementById("table-row-count");
      const paginationInfo = document.getElementById("pagination-info");
      const pageNum = document.getElementById("page-num");

      rowCountSpan.textContent = `${{filteredData.length}} 筆`;

      if (filteredData.length === 0) {{
        tbody.innerHTML = `<tr><td colspan="6" class="text-center py-8 text-slate-400">查無符合條件的航班數據</td></tr>`;
        paginationInfo.textContent = `顯示 0 到 0 筆`;
        pageNum.textContent = "1 / 1";
        return;
      }}

      const totalPages = Math.ceil(filteredData.length / pageSize);
      const startIdx = (currentPage - 1) * pageSize;
      const endIdx = Math.min(startIdx + pageSize, filteredData.length);
      const pageRows = filteredData.slice(startIdx, endIdx);

      // 最低價識別
      const minPrice = Math.min(...filteredData.filter(d => d.price > 0).map(d => d.price));

      let rowsHtml = "";
      pageRows.forEach(r => {{
        const isLowest = r.price === minPrice && r.price > 0;

        // 航司色彩徽章
        let badgeClass = "bg-slate-100 text-slate-700";
        if (r.airline.includes("長榮")) badgeClass = "bg-emerald-100 text-emerald-800 border border-emerald-200";
        else if (r.airline.includes("中華")) badgeClass = "bg-rose-100 text-rose-800 border border-rose-200";
        else if (r.airline.includes("星宇")) badgeClass = "bg-amber-100 text-amber-800 border border-amber-200";

        // 類型徽章
        let typeBadge = "bg-slate-50 text-slate-600";
        if (r.trip_type.includes("來回")) typeBadge = "bg-teal-50 text-teal-700 font-medium";

        // 日期屬性徽章 (明確區隔基準 vs 彈性)
        let dateAttrBadge = (r.date_type && r.date_type.includes("彈性"))
          ? '<span class="inline-flex items-center px-1.5 py-0.5 rounded text-[11px] bg-amber-50 text-amber-700 font-bold border border-amber-200">💡彈性比價</span>'
          : '<span class="inline-flex items-center px-1.5 py-0.5 rounded text-[11px] bg-emerald-50 text-emerald-700 font-bold border border-emerald-200">🎯基準目標</span>';

        rowsHtml += `
          <tr class="hover:bg-slate-50 transition border-b border-slate-100">
            <td class="py-3 px-4 font-mono text-slate-500 text-xs whitespace-nowrap">${{r.query_time}}</td>
            <td class="py-3 px-4 whitespace-nowrap">
              <span class="inline-flex px-2 py-0.5 rounded text-xs ${{typeBadge}}">${{r.trip_type}}</span>
            </td>
            <td class="py-3 px-4 whitespace-nowrap">
              <div class="flex items-center gap-1.5">
                <span class="font-semibold text-slate-800">${{r.from_ap}} ➔ ${{r.to_ap}}</span>
                ${{dateAttrBadge}}
              </div>
              <div class="text-xs text-slate-500 mt-0.5">${{r.flight_date}}</div>
            </td>
            <td class="py-3 px-4 whitespace-nowrap">
              <span class="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-semibold ${{badgeClass}}">
                ${{r.airline}}
              </span>
            </td>
            <td class="py-3 px-4 text-xs text-slate-600">${{r.itinerary}}</td>
            <td class="py-3 px-4 text-right whitespace-nowrap">
              <span class="font-extrabold text-sm md:text-base ${{isLowest ? 'text-emerald-600' : 'text-slate-900'}}">
                ${{formatMoney(r.price)}}
              </span>
              ${{isLowest ? '<span class="ml-1 text-[11px] bg-emerald-500 text-white px-1.5 py-0.2 rounded font-bold">最低</span>' : ''}}
            </td>
          </tr>
        `;
      }});

      tbody.innerHTML = rowsHtml;
      paginationInfo.textContent = `顯示 ${{startIdx + 1}} 到 ${{endIdx}} 筆（共 ${{filteredData.length}} 筆）`;
      pageNum.textContent = `${{currentPage}} / ${{totalPages}}`;

      document.getElementById("btn-prev").disabled = (currentPage === 1);
      document.getElementById("btn-next").disabled = (currentPage >= totalPages);
    }}

    function changePage(delta) {{
      const totalPages = Math.ceil(filteredData.length / pageSize);
      currentPage += delta;
      if (currentPage < 1) currentPage = 1;
      if (currentPage > totalPages) currentPage = totalPages;
      renderTable();
    }}

    // 匯出 CSV
    function exportFilteredCSV() {{
      if (!filteredData || filteredData.length === 0) return;
      const headers = ["查詢時間", "航程類型", "出發地", "目的地", "航班日期", "航空公司", "起降時間與機型", "票價(TWD)"];
      const rows = filteredData.map(d => [
        `"${{d.query_time}}"`,
        `"${{d.trip_type}}"`,
        `"${{d.from_ap}}"`,
        `"${{d.to_ap}}"`,
        `"${{d.flight_date}}"`,
        `"${{d.airline}}"`,
        `"${{d.itinerary.replace(/"/g, '""')}}"`,
        d.price
      ]);

      const csvContent = "\uFEFF" + [headers.join(","), ...rows.map(r => r.join(","))].join("\\n");
      const blob = new Blob([csvContent], {{ type: "text/csv;charset=utf-8;" }});
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.setAttribute("href", url);
      link.setAttribute("download", `機票歷史報表_${{new Date().toISOString().slice(0,10)}}.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }}
  </script>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"✅ 網頁儀表板已成功生成至：{output_path}")

if __name__ == "__main__":
    base_dir = r"E:\Antigravity_Data\flight-price-tracker"
    csv = os.path.join(base_dir, "price_history.csv")
    cfg = os.path.join(base_dir, "config.json")
    out = os.path.join(base_dir, "dashboard.html")
    generate_dashboard_html(csv, cfg, out)
