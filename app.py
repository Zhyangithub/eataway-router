# ============================================================
# app.py — 完全自包含版本，不依赖 route_optimizer.py
# ============================================================
from flask import Flask, jsonify, render_template, request, send_file
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
import json, os, io, urllib.parse, threading
from datetime import datetime
import pandas as pd
import requests as http_requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)

# ── 配置 ────────────────────────────────────────────────────
API_KEY        = "AIzaSyDzAxmeKWeWQGK3G5VVnEDLM0IY-RTjzrw"
WAREHOUSE_COORD = "59.8542194,17.6650221"
DRIVERS        = ["Abbe", "Saman", "Sarkis", "Cornelia", "Pawlos"]

STATE_FILE  = "last_results.json"
PHONES_FILE = "driver_phones.json"

state = {
    "results": {},
    "generated_at": None,
    "schedule_hour": 7,
    "schedule_minute": 0,
    "running": False,
}
driver_phones = {d: "" for d in DRIVERS}
scheduler = BackgroundScheduler()


# ── 持久化 ──────────────────────────────────────────────────
def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            saved = json.load(f)
            state["results"]         = saved.get("results", {})
            state["generated_at"]    = saved.get("generated_at")
            state["schedule_hour"]   = saved.get("schedule_hour", 7)
            state["schedule_minute"] = saved.get("schedule_minute", 0)
    if os.path.exists(PHONES_FILE):
        with open(PHONES_FILE, "r", encoding="utf-8") as f:
            driver_phones.update(json.load(f))

def save_state():
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "results":         state["results"],
            "generated_at":    state["generated_at"],
            "schedule_hour":   state["schedule_hour"],
            "schedule_minute": state["schedule_minute"],
        }, f, ensure_ascii=False, indent=2)

def save_phones():
    with open(PHONES_FILE, "w", encoding="utf-8") as f:
        json.dump(driver_phones, f, ensure_ascii=False, indent=2)


# ── 核心逻辑（来自 main.py）────────────────────────────────
def load_and_merge_data(driver_name):
    coords_file = 'coords.xlsx' if os.path.exists('coords.xlsx') else 'coords.csv'
    try:
        df_coords = pd.read_excel(coords_file, engine='openpyxl')
    except Exception as e:
        return [], f"读取坐标文件失败: {e}"

    header_row_idx = 0
    for i in range(min(15, len(df_coords))):
        if any('namn' in str(v).lower() for v in df_coords.iloc[i].tolist()):
            header_row_idx = i
            break
    df_coords.columns = df_coords.iloc[header_row_idx]
    df_coords = df_coords.iloc[header_row_idx+1:].reset_index(drop=True).iloc[:, :3]
    df_coords.columns = ['Namn', 'Latitude', 'Longitude']
    df_coords['match_name'] = df_coords['Namn'].astype(str).str.strip().str.lower()
    df_coords = df_coords.drop_duplicates(subset=['match_name'], keep='first')
    coord_dict = df_coords.set_index('match_name')[['Latitude', 'Longitude']].to_dict('index')

    routes_file = 'routes.xlsx' if os.path.exists('routes.xlsx') else 'routes.csv'
    try:
        df_routes = pd.read_excel(routes_file, engine='openpyxl')
    except Exception as e:
        return [], f"读取路线文件失败: {e}"

    driver_col = None
    name_row   = -1
    for col in df_routes.columns:
        if driver_name.lower() in str(col).lower():
            driver_col = col; break
        for r_idx in range(min(15, len(df_routes))):
            if driver_name.lower() in str(df_routes[col].iloc[r_idx]).lower():
                driver_col = col; name_row = r_idx; break
        if driver_col: break

    if driver_col is None:
        return [], f"找不到司机 {driver_name}"

    raw = (df_routes[driver_col].iloc[1:].tolist()
           if name_row == -1 else df_routes[driver_col].iloc[name_row+2:].tolist())

    matched, unmatched = [], []
    for store in raw:
        if pd.isna(store) or str(store).strip() in ("", "nan"):
            continue
        key = str(store).strip().lower()
        if key in coord_dict:
            matched.append({
                "name": str(store).strip(),
                "lat":  str(coord_dict[key]['Latitude']),
                "lng":  str(coord_dict[key]['Longitude']),
            })
        else:
            unmatched.append(str(store).strip())
    return matched, unmatched


def optimize_route(stores):
    waypoints_str = "optimize:true|" + "|".join(f"{s['lat']},{s['lng']}" for s in stores)
    params = {
        "origin":      WAREHOUSE_COORD,
        "destination": WAREHOUSE_COORD,
        "waypoints":   waypoints_str,
        "key":         API_KEY,
        "departure_time": "now",
    }
    resp = http_requests.get(
        "https://maps.googleapis.com/maps/api/directions/json", params=params)
    data = resp.json()
    if data['status'] == 'OK':
        route  = data['routes'][0]
        order  = route['waypoint_order']
        legs   = route['legs']
        dur_s  = sum(l['duration']['value']  for l in legs)
        dist_m = sum(l['distance']['value']  for l in legs)
        return [stores[i] for i in order], {
            "duration_min": round(dur_s / 60),
            "distance_km":  round(dist_m / 1000, 1),
        }
    return None, data.get('error_message', data['status'])


def generate_urls(optimized_stores):
    urls   = []
    wh_lat, wh_lng = WAREHOUSE_COORD.split(',')
    wh     = {"name": "Lager (Uppsala)", "lat": wh_lat, "lng": wh_lng}
    path   = [wh] + optimized_stores + [wh]
    for i in range(0, len(path) - 1, 10):
        chunk  = path[i: i+11]
        origin = chunk[0]; dest = chunk[-1]; wps = chunk[1:-1]
        wp_str = ""
        if wps:
            wp_str = "&waypoints=" + urllib.parse.quote(
                "|".join(f"{s['lat']},{s['lng']}" for s in wps))
        urls.append(
            f"https://www.google.com/maps/dir/?api=1"
            f"&origin={origin['lat']},{origin['lng']}"
            f"&destination={dest['lat']},{dest['lng']}{wp_str}"
        )
    return urls


def run_all_drivers():
    results = {}
    for driver in DRIVERS:
        try:
            stores, unmatched = load_and_merge_data(driver)
            if not stores:
                err = unmatched if isinstance(unmatched, str) else f"未匹配到任何门店"
                results[driver] = {"status": "error", "error": err}
                continue

            optimized, stats_or_err = optimize_route(stores)
            if not optimized:
                results[driver] = {"status": "error", "error": str(stats_or_err)}
                continue

            urls = generate_urls(optimized)
            results[driver] = {
                "status":       "ok",
                "stores":       [s["name"] for s in optimized],
                "store_count":  len(optimized),
                "urls":         urls,
                "duration":     f"{stats_or_err['duration_min']} min",
                "distance":     f"{stats_or_err['distance_km']} km",
                "unmatched":    unmatched if isinstance(unmatched, list) else [],
                "unmatched_count": len(unmatched) if isinstance(unmatched, list) else 0,
            }
        except Exception as e:
            results[driver] = {"status": "error", "error": str(e)}
    return results


# ── 后台任务 ─────────────────────────────────────────────────
def do_generate():
    if state["running"]:
        return
    state["running"] = True
    try:
        state["results"]      = run_all_drivers()
        state["generated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        save_state()
    finally:
        state["running"] = False


def reschedule(hour, minute):
    if scheduler.get_job("daily_gen"):
        scheduler.remove_job("daily_gen")
    scheduler.add_job(do_generate, CronTrigger(hour=hour, minute=minute),
                      id="daily_gen", replace_existing=True)


# ── Flask 路由 ───────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/status")
def api_status():
    return jsonify({
        "results":         state["results"],
        "generated_at":    state["generated_at"],
        "schedule_hour":   state["schedule_hour"],
        "schedule_minute": state["schedule_minute"],
        "running":         state["running"],
        "drivers":         DRIVERS,
        "phones":          driver_phones,
    })


@app.route("/api/generate", methods=["POST"])
def api_generate():
    if state["running"]:
        return jsonify({"ok": False, "message": "Already running"}), 409
    threading.Thread(target=do_generate, daemon=True).start()
    return jsonify({"ok": True})


@app.route("/api/schedule", methods=["POST"])
def api_schedule():
    data   = request.json
    hour   = int(data.get("hour",   7))
    minute = int(data.get("minute", 0))
    state["schedule_hour"]   = hour
    state["schedule_minute"] = minute
    reschedule(hour, minute)
    save_state()
    return jsonify({"ok": True, "hour": hour, "minute": minute})


@app.route("/api/phones", methods=["POST"])
def api_set_phones():
    for driver, number in request.json.items():
        if driver in driver_phones:
            driver_phones[driver] = str(number).strip()
    save_phones()
    return jsonify({"ok": True, "phones": driver_phones})


@app.route("/api/export")
def api_export():
    if not state["results"]:
        return jsonify({"error": "暂无结果"}), 400

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Route Links"

    hdr_fill   = PatternFill("solid", fgColor="1A1A2E")
    hdr_font   = Font(color="F5A623", bold=True, size=11)
    ok_fill    = PatternFill("solid", fgColor="0D2137")
    err_fill   = PatternFill("solid", fgColor="2D1A1A")
    wht_font   = Font(color="E0E0E0", size=10)
    center     = Alignment(horizontal="center",  vertical="center")
    left_wrap  = Alignment(horizontal="left",    vertical="center", wrap_text=True)
    thin       = Border(**{s: Side(style='thin', color='333355')
                           for s in ('left','right','top','bottom')})

    headers    = ["Chaufför","Butiker","Tid","Distans","Status",
                  "Segment 1","Segment 2","Segment 3"]
    widths     = [12, 8, 12, 12, 10, 60, 60, 60]

    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.fill = hdr_fill; c.font = hdr_font
        c.alignment = center; c.border = thin
        ws.column_dimensions[c.column_letter].width = w
    ws.row_dimensions[1].height = 22

    for ri, driver in enumerate(DRIVERS, 2):
        r    = state["results"].get(driver, {})
        fill = ok_fill if r.get("status") == "ok" else err_fill
        ws.row_dimensions[ri].height = 30

        def mc(col, val, row=ri, f=fill):
            c = ws.cell(row=row, column=col, value=val)
            c.fill = f; c.font = wht_font; c.border = thin
            return c

        mc(1, driver).alignment = center
        mc(2, r.get("store_count", "—")).alignment = center
        mc(3, r.get("duration",    "—")).alignment = center
        mc(4, r.get("distance",    "—")).alignment = center
        mc(5, "Klar" if r.get("status")=="ok" else f"Fel: {r.get('error','')}").alignment = center
        for si in range(3):
            url = (r.get("urls") or [])[si] if si < len(r.get("urls") or []) else ""
            c = mc(6+si, url); c.alignment = left_wrap
            if url: c.font = Font(color="4A9FD4", size=9, underline="single")

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return send_file(buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"rutter_{datetime.now().strftime('%Y%m%d')}.xlsx")



@app.route("/links/<driver_name>")
def driver_links(driver_name):
    r = state["results"].get(driver_name)
    date = state.get("generated_at", "—")
    if not r or r.get("status") != "ok":
        return f"""<!DOCTYPE html><html><head><meta charset="UTF-8">
        <meta name="viewport" content="width=device-width,initial-scale=1">
        <title>{driver_name}</title></head>
        <body style="font-family:sans-serif;padding:2rem;background:#111;color:#fff">
        <h2>Inga rutter för {driver_name} ännu.</h2></body></html>""", 404

    urls   = r.get("urls", [])
    stores = r.get("stores", [])
    store_list = "".join(
        f'<div style="padding:6px 0;border-bottom:1px solid #222;font-size:14px">'
        f'<span style="color:#f5a623;margin-right:8px">{i+1}.</span>{s}</div>'
        for i, s in enumerate(stores)
    )
    link_btns = "".join(
        f'<a href="{u}" style="display:block;margin:10px 0;padding:14px 18px;'
        f'background:#1a73e8;color:#fff;text-decoration:none;border-radius:8px;'
        f'font-size:15px;font-weight:600;text-align:center">🗺 Segment {i+1} — Öppna i Google Maps</a>'
        for i, u in enumerate(urls)
    )
    return f"""<!DOCTYPE html>
<html><head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Körorder — {driver_name}</title>
  <style>
    * {{ box-sizing:border-box; margin:0; padding:0; }}
    body {{ font-family:-apple-system,BlinkMacSystemFont,sans-serif; background:#0f111a; color:#e0e6f0; min-height:100vh; }}
    .header {{ background:#161b27; border-bottom:1px solid #1e2d45; padding:18px 20px; }}
    .name {{ font-size:26px; font-weight:800; color:#fff; }}
    .meta {{ font-size:13px; color:#6b7a99; margin-top:4px; }}
    .stats {{ display:grid; grid-template-columns:repeat(3,1fr); border-bottom:1px solid #1e2d45; }}
    .stat {{ padding:16px; text-align:center; border-right:1px solid #1e2d45; }}
    .stat:last-child {{ border-right:none; }}
    .stat-val {{ font-size:22px; font-weight:700; color:#f5a623; display:block; }}
    .stat-lbl {{ font-size:11px; color:#6b7a99; text-transform:uppercase; letter-spacing:.5px; }}
    .section {{ padding:16px 20px; }}
    .section-title {{ font-size:11px; color:#6b7a99; text-transform:uppercase; letter-spacing:1px; margin-bottom:12px; }}
    .store-list {{ background:#161b27; border-radius:8px; padding:4px 12px; max-height:260px; overflow-y:auto; margin-bottom:4px; }}
    .toggle {{ background:none; border:1px solid #1e2d45; color:#6b7a99; padding:8px 14px; border-radius:6px; font-size:13px; cursor:pointer; margin-bottom:16px; width:100%; }}
    #store-list {{ display:none; }}
  </style>
</head>
<body>
  <div class="header">
    <div class="name">🚛 {driver_name}</div>
    <div class="meta">Genererad: {date}</div>
  </div>
  <div class="stats">
    <div class="stat"><span class="stat-val">{r.get("store_count","—")}</span><span class="stat-lbl">Butiker</span></div>
    <div class="stat"><span class="stat-val">{r.get("duration","—")}</span><span class="stat-lbl">Est. tid</span></div>
    <div class="stat"><span class="stat-val">{r.get("distance","—")}</span><span class="stat-lbl">Distans</span></div>
  </div>
  <div class="section">
    <div class="section-title">Navigationslänkar</div>
    {link_btns}
    <button class="toggle" onclick="var l=document.getElementById('store-list');l.style.display=l.style.display=='none'?'block':'none';this.textContent=l.style.display=='block'?'▲ Dölj körordning':'▼ Visa körordning ({len(stores)} stopp)'">▼ Visa körordning ({len(stores)} stopp)</button>
    <div id="store-list"><div class="store-list">{store_list}</div></div>
  </div>
</body></html>"""


# ── 启动 ─────────────────────────────────────────────────────
# 启动时始终加载状态和调度器（gunicorn 也需要）
load_state()
scheduler.start()
reschedule(state["schedule_hour"], state["schedule_minute"])

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(debug=False, host="0.0.0.0", port=port, use_reloader=False)
