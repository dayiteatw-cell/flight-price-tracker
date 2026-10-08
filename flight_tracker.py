# -*- coding: utf-8 -*-
"""
機票即時價格監控與歷史追蹤機器人 (Flight Price Tracker) - 高智商智慧升級版
特色功能：
1. 【明確區隔】「我的基準目標日期」與「前後彈性比價（±1~2天）」
2. 【智能買點天梯】自動比對前後日期排列組合，計算現省金額並抓出神仙買點
3. 【結構化 LINE 戰報】手機推播一眼看清基準日 vs 彈性推薦日的性價比
4. 【動態網頁儀表板】同步更新 dashboard.html，具備九宮格交叉矩陣與走勢圖
5. 【雙格式儲存】CSV 與 Excel 同步追加「日期屬性」標籤
"""

import os
import sys
import json
import time
import urllib.request
import urllib.error
from datetime import datetime, timedelta

# Windows Console UTF-8 相容性處理
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import pandas as pd
from fast_flights import FlightQuery, create_query, get_flights

try:
    from dashboard_generator import generate_dashboard_html
except ImportError:
    generate_dashboard_html = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
HISTORY_CSV = os.path.join(BASE_DIR, "price_history.csv")
HISTORY_EXCEL = os.path.join(BASE_DIR, "price_history.xlsx")
DASHBOARD_HTML = os.path.join(BASE_DIR, "dashboard.html")

AIRLINE_ALIASES = {
    "EVA Air": ["eva air", "長榮", "長榮航空", "br"],
    "China Airlines": ["china airlines", "華航", "中華航空", "ci"],
    "STARLUX Airlines": ["starlux", "星宇", "星宇航空", "jx"],
}

CONFIG_LOCAL_FILE = os.path.join(BASE_DIR, "config.local.json")
INDEX_HTML = os.path.join(BASE_DIR, "index.html")

def load_config():
    if not os.path.exists(CONFIG_FILE):
        raise FileNotFoundError(f"找不到設定檔：{CONFIG_FILE}")
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    # 若存在本地私密設定檔，優先覆蓋合併
    if os.path.exists(CONFIG_LOCAL_FILE):
        try:
            with open(CONFIG_LOCAL_FILE, "r", encoding="utf-8") as f:
                local_cfg = json.load(f)
                if "line_settings" in local_cfg:
                    cfg["line_settings"] = {**cfg.get("line_settings", {}), **local_cfg["line_settings"]}
        except Exception:
            pass

    # 環境變數支援（供 GitHub Actions Secrets 注入）
    env_token = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")
    env_uid = os.environ.get("LINE_USER_ID")
    if env_token:
        cfg.setdefault("line_settings", {})["channel_access_token"] = env_token.strip()
        cfg["line_settings"]["enabled"] = True
    if env_uid:
        cfg.setdefault("line_settings", {})["user_id"] = env_uid.strip()
        cfg["line_settings"]["enabled"] = True

    return cfg

def parse_flight_dates(config):
    """解析日期設定，支援新版明確區隔之結構，並維持舊版相容性"""
    fd = config.get("flight_dates", {})
    base = fd.get("base_target", {})
    flex = fd.get("flexible_scan", {})
    s = config.get("search_settings", {})

    dep_date = base.get("depart_date") or s.get("depart_date", "2026-12-15")
    ret_date = base.get("return_date") or s.get("return_date", "2027-01-10")

    flex_enabled = flex.get("enabled", False) if "enabled" in flex else config.get("flexible_date_settings", {}).get("enabled", False)
    dep_flex_days = flex.get("depart_flex_days", 1)
    ret_flex_days = flex.get("return_flex_days", 1)
    min_savings = flex.get("min_savings_to_alert", 1500)

    return {
        "depart_date": dep_date,
        "return_date": ret_date,
        "flex_enabled": flex_enabled,
        "depart_flex_days": dep_flex_days,
        "return_flex_days": ret_flex_days,
        "min_savings": min_savings
    }

def format_time(t):
    if not t:
        return "--:--"
    return f"{t[0]:02d}:{t[1]:02d}"

def is_target_airline(airline_name, target_airlines):
    if not target_airlines:
        return True
    a_lower = str(airline_name).lower()
    for target in target_airlines:
        t_lower = str(target).lower()
        if t_lower in a_lower or a_lower in t_lower:
            return True
        for canon, aliases in AIRLINE_ALIASES.items():
            if any(al in t_lower for al in aliases) and any(al in a_lower for al in aliases):
                return True
    return False

def fetch_flights(from_code, to_code, depart_date, return_date=None, trip_type="one-way", direct_only=True, currency="TWD", seat_class="economy"):
    flights_query = [
        FlightQuery(date=depart_date, from_airport=from_code, to_airport=to_code, max_stops=0 if direct_only else None)
    ]
    if trip_type == "round-trip" and return_date:
        flights_query.append(
            FlightQuery(date=return_date, from_airport=to_code, to_airport=from_code, max_stops=0 if direct_only else None)
        )

    query = create_query(
        flights=flights_query,
        trip=trip_type,
        seat=seat_class,
        max_stops=0 if direct_only else None,
        currency=currency,
        language="zh-TW"
    )

    try:
        results = get_flights(query)
    except Exception as e:
        return []

    parsed = []
    for item in results:
        airline_names = item.airlines if isinstance(item.airlines, list) else [str(item.airlines)]
        segments_info = []
        for seg in item.flights:
            dep_time = format_time(seg.departure.time)
            arr_time = format_time(seg.arrival.time)
            plane = seg.plane_type or "機型未定"
            segments_info.append(f"{seg.from_airport.code}->{seg.to_airport.code} {dep_time}~{arr_time} ({plane})")

        parsed.append({
            "airline": " / ".join(airline_names),
            "price": item.price,
            "itinerary": " ｜ ".join(segments_info)
        })
    return parsed

def send_line_notification(title, message_text, config):
    """透過 LINE Messaging API 發送推播訊息"""
    line_cfg = config.get("line_settings", {})
    if not line_cfg.get("enabled", False):
        return False

    token = line_cfg.get("channel_access_token", "").strip()
    user_id = line_cfg.get("user_id", "").strip()

    if not token or not user_id:
        return False

    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }

    full_message = f"{title}\n\n{message_text}"
    payload = {
        "to": user_id,
        "messages": [
            {
                "type": "text",
                "text": full_message
            }
        ]
    }

    try:
        req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=10) as resp:
            if resp.status == 200:
                print("📱 【LINE 通知成功】已成功推播訊息至您的 LINE！")
                return True
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode("utf-8", errors="replace")
        print(f"❌ 【LINE 推播失敗】HTTP {e.code}: {err_msg}")
    except Exception as e:
        print(f"❌ 【LINE 推播異常】：{e}")
    return False

def scan_flexible_matrix(s, date_cfg, base_lowest_price):
    """
    前後彈性日期交叉矩陣比價
    自動計算各組合與基準日之價差（現省多少元），並依性價比排序
    """
    base_dep = date_cfg["depart_date"]
    base_ret = date_cfg["return_date"]
    dep_days = date_cfg["depart_flex_days"]
    ret_days = date_cfg["return_flex_days"]

    print(f"\n🔎 正在啟動「前後彈性日期矩陣比價」（去程 ±{dep_days} 天 ｜ 回程 ±{ret_days} 天）...")
    dep_center = datetime.strptime(base_dep, "%Y-%m-%d")
    ret_center = datetime.strptime(base_ret, "%Y-%m-%d") if base_ret else None

    results = []
    # 遍歷去程與回程的所有排列組合
    for d_offset in range(-dep_days, dep_days + 1):
        dep_d = (dep_center + timedelta(days=d_offset)).strftime("%Y-%m-%d")
        if ret_center:
            for r_offset in range(-ret_days, ret_days + 1):
                ret_d = (ret_center + timedelta(days=r_offset)).strftime("%Y-%m-%d")
                if ret_d <= dep_d:
                    continue
                is_base = (dep_d == base_dep and ret_d == base_ret)
                raw = fetch_flights(s["from_airport"], s["to_airport"], dep_d, ret_d, "round-trip", s.get("direct_only", True), s.get("currency", "TWD"), s.get("seat_class", "economy"))
                matched = [f for f in raw if is_target_airline(f["airline"], s.get("target_airlines", [])) and f["price"]]
                if matched:
                    matched.sort(key=lambda x: x["price"])
                    best_match = matched[0]
                    price = best_match["price"]
                    savings = (base_lowest_price - price) if base_lowest_price else 0
                    results.append({
                        "depart_date": dep_d,
                        "return_date": ret_d,
                        "is_base": is_base,
                        "airline": best_match["airline"],
                        "price": price,
                        "savings": savings,
                        "itinerary": best_match["itinerary"]
                    })
        else:
            is_base = (dep_d == base_dep)
            raw = fetch_flights(s["from_airport"], s["to_airport"], dep_d, None, "one-way", s.get("direct_only", True), s.get("currency", "TWD"), s.get("seat_class", "economy"))
            matched = [f for f in raw if is_target_airline(f["airline"], s.get("target_airlines", [])) and f["price"]]
            if matched:
                matched.sort(key=lambda x: x["price"])
                best_match = matched[0]
                price = best_match["price"]
                savings = (base_lowest_price - price) if base_lowest_price else 0
                results.append({
                    "depart_date": dep_d,
                    "return_date": "單程",
                    "is_base": is_base,
                    "airline": best_match["airline"],
                    "price": price,
                    "savings": savings,
                    "itinerary": best_match["itinerary"]
                })

    if results:
        results.sort(key=lambda x: x["price"])
        print("\n" + "=" * 72)
        print("💡 【前後彈性比價智能分析天梯】")
        print("=" * 72)
        print(f"{'去程日期':<11} | {'回程日期':<11} | {'最低航司':<8} | {'最低總價':<12} | {'與基準日價差對比'}")
        print("-" * 72)
        for idx, r in enumerate(results):
            tag = ""
            if r["is_base"]:
                diff_str = "🎯 我的基準目標日期"
            elif r["savings"] > 0:
                diff_str = f"🟢 現省 NT$ {r['savings']:,}！"
            elif r["savings"] < 0:
                diff_str = f"🔴 貴 NT$ {-r['savings']:,}"
            else:
                diff_str = "持平"

            print(f"{r['depart_date']:<11} | {r['return_date']:<11} | {r['airline'][:8]:<8} | NT$ {r['price']:<8,} | {diff_str}")

    return results

def execute_tracking_cycle(config, run_flex=False):
    s = config["search_settings"]
    date_cfg = parse_flight_dates(config)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    base_dep = date_cfg["depart_date"]
    base_ret = date_cfg["return_date"]

    print("=" * 72)
    print("✈️  機票價格自動監控追蹤機器人 (TPE <-> LAX)  ✈️")
    print("=" * 72)
    print(f"查詢時間：{now_str}")
    print(f"航線設定：{s['from_airport']} ⇋ {s['to_airport']} ({'來回' if s['trip_type'] == 'round-trip' else '單程'})")
    print(f"🎯 基準目標日期：去程 {base_dep} ｜ 回程 {base_ret}")
    if date_cfg["flex_enabled"] or run_flex:
        print(f"💡 前後彈性範圍：去程 ±{date_cfg['depart_flex_days']} 天 ｜ 回程 ±{date_cfg['return_flex_days']} 天")
    print(f"指定航司：{', '.join(s.get('target_airlines', []))}")
    print(f"直飛篩選：{'僅直飛 (Non-stop)' if s['direct_only'] else '包含轉機'}")
    print("-" * 72)
    print("⏳ 正在即時抓取最新票價數據，請稍候...")

    target_airlines = s.get("target_airlines", [])
    records = []

    # 1. 抓取【基準目標日期】來回直飛套票
    round_trip_flights = []
    if s.get("trip_type") == "round-trip" and base_ret:
        raw_rt = fetch_flights(
            from_code=s["from_airport"],
            to_code=s["to_airport"],
            depart_date=base_dep,
            return_date=base_ret,
            trip_type="round-trip",
            direct_only=s.get("direct_only", True),
            currency=s.get("currency", "TWD"),
            seat_class=s.get("seat_class", "economy")
        )
        round_trip_flights = [f for f in raw_rt if is_target_airline(f["airline"], target_airlines)]
        round_trip_flights.sort(key=lambda x: (x["price"] is None, x["price"]))

        for f in round_trip_flights:
            records.append({
                "查詢時間": now_str,
                "日期類別": "🎯基準目標",
                "航程類型": "來回套票 (直飛)",
                "出發地": s["from_airport"],
                "目的地": s["to_airport"],
                "航班日期": f"{base_dep} 往返 {base_ret}",
                "航空公司": f["airline"],
                "航班時間與機型": f["itinerary"],
                "票價(TWD)": f["price"]
            })

    # 2. 基準去程獨立單程
    raw_outbound = fetch_flights(
        from_code=s["from_airport"],
        to_code=s["to_airport"],
        depart_date=base_dep,
        trip_type="one-way",
        direct_only=s.get("direct_only", True),
        currency=s.get("currency", "TWD"),
        seat_class=s.get("seat_class", "economy")
    )
    outbound_flights = [f for f in raw_outbound if is_target_airline(f["airline"], target_airlines)]
    outbound_flights.sort(key=lambda x: (x["price"] is None, x["price"]))

    for f in outbound_flights:
        records.append({
            "查詢時間": now_str,
            "日期類別": "🎯基準目標",
            "航程類型": "去程單程 (直飛)",
            "出發地": s["from_airport"],
            "目的地": s["to_airport"],
            "航班日期": base_dep,
            "航空公司": f["airline"],
            "航班時間與機型": f["itinerary"],
            "票價(TWD)": f["price"]
        })

    # 3. 基準回程獨立單程
    inbound_flights = []
    if base_ret:
        raw_inbound = fetch_flights(
            from_code=s["to_airport"],
            to_code=s["from_airport"],
            depart_date=base_ret,
            trip_type="one-way",
            direct_only=s.get("direct_only", True),
            currency=s.get("currency", "TWD"),
            seat_class=s.get("seat_class", "economy")
        )
        inbound_flights = [f for f in raw_inbound if is_target_airline(f["airline"], target_airlines)]
        inbound_flights.sort(key=lambda x: (x["price"] is None, x["price"]))

        for f in inbound_flights:
            records.append({
                "查詢時間": now_str,
                "日期類別": "🎯基準目標",
                "航程類型": "回程單程 (直飛)",
                "出發地": s["to_airport"],
                "目的地": s["from_airport"],
                "航班日期": base_ret,
                "航空公司": f["airline"],
                "航班時間與機型": f["itinerary"],
                "票價(TWD)": f["price"]
            })

    # 基準最低價
    lowest_rt = round_trip_flights[0]["price"] if round_trip_flights and round_trip_flights[0]["price"] else None
    lowest_out = outbound_flights[0]["price"] if outbound_flights and outbound_flights[0]["price"] else None
    base_current_lowest = lowest_rt if lowest_rt else lowest_out
    base_lowest_airline = round_trip_flights[0]["airline"] if round_trip_flights else (outbound_flights[0]["airline"] if outbound_flights else "")

    # 4. 執行【前後彈性日期比價】(若開啟)
    flex_results = []
    should_run_flex = run_flex or date_cfg["flex_enabled"]
    if should_run_flex:
        flex_results = scan_flexible_matrix(s, date_cfg, base_current_lowest)
        # 把非基準目標的彈性日期也記錄進歷史庫
        for fr in flex_results:
            if not fr["is_base"]:
                records.append({
                    "查詢時間": now_str,
                    "日期類別": "💡彈性比價",
                    "航程類型": "來回套票 (直飛)",
                    "出發地": s["from_airport"],
                    "目的地": s["to_airport"],
                    "航班日期": f"{fr['depart_date']} 往返 {fr['return_date']}",
                    "航空公司": fr["airline"],
                    "航班時間與機型": fr["itinerary"],
                    "票價(TWD)": fr["price"]
                })

    # 寫入歷史資料庫 (標準 9 欄位格式)
    if records:
        df_new = pd.DataFrame(records)
        standard_cols = ["查詢時間", "日期類別", "航程類型", "出發地", "目的地", "航班日期", "航空公司", "航班時間與機型", "票價(TWD)"]
        for c in standard_cols:
            if c not in df_new.columns:
                df_new[c] = ""
        df_new = df_new[standard_cols]

        if os.path.exists(HISTORY_CSV):
            try:
                df_old = pd.read_csv(HISTORY_CSV, encoding="utf-8-sig")
                for col in standard_cols:
                    if col not in df_old.columns:
                        df_old[col] = "🎯基準目標" if col == "日期類別" else ""
                df_old = df_old[standard_cols]
                df_all = pd.concat([df_old, df_new], ignore_index=True)
            except Exception:
                df_all = df_new
        else:
            df_all = df_new

        df_all.to_csv(HISTORY_CSV, index=False, encoding="utf-8-sig")
        try:
            df_all.to_excel(HISTORY_EXCEL, index=False)
        except Exception:
            pass

    # 螢幕終端報告
    print("\n✅ 票價數據解析完成！\n")
    print(f"【🎯 我的基準目標日期實時行情】 (去程 {base_dep} ⇋ 回程 {base_ret})")
    if round_trip_flights:
        print(f"{'航空公司':<10} | {'來回總價 (TWD)':<14} | {'去程班次時間與機型'}")
        print("-" * 72)
        for f in round_trip_flights:
            price_str = f"NT$ {f['price']:,}" if f['price'] else "價格未載"
            print(f"{f['airline']:<10} | {price_str:<14} | {f['itinerary']}")
    else:
        print("  ⚠️ 本指定日期無整組直飛來回套票。")
    print()

    # 警報判斷與 LINE 智慧通知
    threshold = config.get("alert_settings", {}).get("price_alert_threshold")
    alert_cfg = config.get("alert_settings", {})
    notify_flex = alert_cfg.get("notify_if_flexible_cheaper", True)
    min_savings_target = date_cfg.get("min_savings", 1500)

    # 找出最佳彈性組合
    best_flex = flex_results[0] if flex_results else None
    has_cheaper_flex = (best_flex and not best_flex["is_base"] and best_flex["savings"] >= min_savings_target)

    # 判斷是否觸發門檻
    threshold_met = (threshold and base_current_lowest and base_current_lowest <= threshold)

    print("-" * 72)
    if base_current_lowest:
        print(f"🎯 基準目標最低價：NT$ {base_current_lowest:,} ({base_lowest_airline})")
    if has_cheaper_flex:
        print(f"💡 彈性推薦神仙買點：{best_flex['depart_date']} ⇋ {best_flex['return_date']} 只要 NT$ {best_flex['price']:,}（現省 NT$ {best_flex['savings']:,}！）")

    # 組裝結構化智慧 LINE 戰報
    should_send_line = threshold_met or has_cheaper_flex

    if should_send_line:
        line_title = "✈️【台美航線・智慧最佳買點快報】"
        line_lines = [
            f"🎯 您的基準目標日期：",
            f"• 航線：{s['from_airport']} ⇋ {s['to_airport']} (直飛來回)",
            f"• 日期：去程 {base_dep} ｜ 回程 {base_ret}",
            f"• 目前最低：NT$ {base_current_lowest:,} 元 ({base_lowest_airline})",
            f"• 警報門檻：NT$ {threshold:,} 元"
        ]

        if has_cheaper_flex:
            line_lines.append("")
            line_lines.append("💡 智能推薦【彈性出發省更多】：")
            line_lines.append(f"• 推薦日期：{best_flex['depart_date']} ⇋ {best_flex['return_date']}")
            line_lines.append(f"• 最低總價：NT$ {best_flex['price']:,} 元 ({best_flex['airline']})")
            line_lines.append(f"✨ 效益：比基準日現省 NT$ {best_flex['savings']:,} 元！")
            line_lines.append(f"🌟 班次：{best_flex['itinerary'].split('｜')[0]}")

            # 列出天梯排行
            line_lines.append("")
            line_lines.append("📊 前後日期行情天梯：")
            for fr in flex_results[:3]:
                d_desc = f"{fr['depart_date']}➔{fr['return_date']}"
                if fr["is_base"]:
                    line_lines.append(f"  • {d_desc}：NT$ {fr['price']:,} 🟡 (基準日)")
                elif fr["savings"] > 0:
                    line_lines.append(f"  • {d_desc}：NT$ {fr['price']:,} 🟢 (省 {fr['savings']:,})")
                else:
                    line_lines.append(f"  • {d_desc}：NT$ {fr['price']:,} 🔴")

        line_lines.append("")
        line_lines.append("👉 席位有限，若日期合適建議及早開票！")

        line_body = "\n".join(line_lines)
        send_line_notification(line_title, line_body, config)

    # 自動同步刷新網頁儀表板
    if generate_dashboard_html:
        try:
            generate_dashboard_html(HISTORY_CSV, CONFIG_FILE, DASHBOARD_HTML)
            import shutil
            shutil.copy2(DASHBOARD_HTML, INDEX_HTML)
            print(f"\n🌐 【網頁儀表板已同步更新 (含 index.html)】點開 open_dashboard.bat 即可立即檢視：\n  • {DASHBOARD_HTML}")
        except Exception as e:
            print(f"⚠️ 網頁儀表板同步更新異常：{e}")

    print(f"\n📁 歷史價格已儲存至：\n  • {HISTORY_EXCEL}\n  • {HISTORY_CSV}")
    print("=" * 72)

def main():
    config = load_config()

    if "--test-line" in sys.argv:
        print("🧪 正在發送 LINE 推播測試訊息...")
        test_title = "✈️【機票監控機器人】LINE 推播連線測試"
        test_body = f"連線測試成功！當前時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n系統已就緒，隨時待命為您監控票價！"
        send_line_notification(test_title, test_body, config)
        return

    if "--dashboard" in sys.argv:
        if generate_dashboard_html:
            generate_dashboard_html(HISTORY_CSV, CONFIG_FILE, DASHBOARD_HTML)
            import shutil
            shutil.copy2(DASHBOARD_HTML, INDEX_HTML)
            print("✅ 網頁儀表板已單獨重新生成完畢！")
        return

    run_flex = ("--flex" in sys.argv)
    interval_hours = config.get("schedule_settings", {}).get("auto_loop_interval_hours", 0)

    if interval_hours and interval_hours > 0:
        print(f"🚀 已啟動常駐背景輪詢模式（每隔 {interval_hours} 小時自動執行一次）。")
        print("💡 按鍵盤 [Ctrl + C] 可隨時停止程式。\n")
        try:
            while True:
                execute_tracking_cycle(config, run_flex=run_flex)
                next_time = (datetime.now() + timedelta(hours=interval_hours)).strftime("%Y-%m-%d %H:%M:%S")
                print(f"\n💤 進入休眠等待中... 下次自動執行時間：{next_time}")
                print("-" * 72)
                time.sleep(interval_hours * 3600)
                config = load_config()
        except KeyboardInterrupt:
            print("\n🛑 已手動停止自動監控。")
    else:
        execute_tracking_cycle(config, run_flex=run_flex)

if __name__ == "__main__":
    main()
