"""Local commute estimate collector and manual trip timer."""

from __future__ import annotations

import csv
import concurrent.futures
import io
import json
import os
import sqlite3
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


ROOT = Path(__file__).resolve().parent
DB_PATH = ROOT / "commute_data.sqlite3"
SETTINGS_PATH = ROOT / "settings.json"
EXCEL_PATH = ROOT / "commute_data.xlsx"
EXCEL_TEMP_PATH = ROOT / "commute_data.tmp.xlsx"
HEALTH_PATH = ROOT / "collection-health.json"
HOST = "127.0.0.1"
PORT = int(os.environ.get("COMMUTE_PORT", "8765"))
ADDRESSES = {
    "home": "경기도 수원시 영통구 봉영로1517번길 73",
    "work": "세종특별자치시 소정면 소정산단4로 31",
    "via": "충남 천안시 동남구 풍세면 풍서리 산 18-3",
}
ROUTES = {
    "via": {"label": "경유지 경로 (평택파주고속도로 확인)", "priority": "TIME", "must_contain": "평택파주"},
    "shortest": {"label": "최단거리 (경부고속도로 확인)", "priority": "DISTANCE", "must_contain": "경부"},
}
DEFAULT_SETTINGS = {
    "api_key": "",
    "interval_minutes": 10,
    "weekdays_only": True,
    "coordinates": {},
    "collection_start": "2026-11-24T16:00:00+09:00",
    "collection_end": "2026-12-31T20:00:00+09:00",
}
_db_lock = threading.RLock()
_excel_lock = threading.RLock()
_test_lock = threading.Lock()
_state_lock = threading.RLock()
_state = {"last_poll": None, "last_error": None, "excel_error": None, "last_message": "시작 대기 중", "polling": False, "test_running": False, "test_samples": 0}

WEEKDAYS_KO = ("월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일")
WEATHER_FIELDS = {
    "weather_time": "TEXT",
    "weather_code": "INTEGER",
    "weather_description": "TEXT",
    "temperature_c": "REAL",
    "relative_humidity_pct": "REAL",
    "precipitation_mm": "REAL",
    "rain_mm": "REAL",
    "snowfall_cm": "REAL",
    "wind_speed_kmh": "REAL",
    "weather_error": "TEXT",
}


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def db_connect():
    class ClosingConnection(sqlite3.Connection):
        def __exit__(self, exc_type, exc_value, traceback):
            try:
                return super().__exit__(exc_type, exc_value, traceback)
            finally:
                self.close()

    conn = sqlite3.connect(DB_PATH, timeout=20, factory=ClosingConnection)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with _db_lock, db_connect() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS estimates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                observed_at TEXT NOT NULL,
                direction TEXT NOT NULL,
                route_id TEXT NOT NULL,
                duration_seconds INTEGER,
                distance_meters INTEGER,
                road_names TEXT NOT NULL DEFAULT '[]',
                route_match INTEGER NOT NULL DEFAULT 0,
                collection_mode TEXT NOT NULL DEFAULT '정기',
                departure_location TEXT,
                weekday TEXT,
                local_time TEXT,
                weather_time TEXT,
                weather_code INTEGER,
                weather_description TEXT,
                temperature_c REAL,
                relative_humidity_pct REAL,
                precipitation_mm REAL,
                rain_mm REAL,
                snowfall_cm REAL,
                wind_speed_kmh REAL,
                weather_error TEXT,
                error TEXT
            );
            CREATE INDEX IF NOT EXISTS estimates_observed_idx ON estimates(observed_at);
            CREATE TABLE IF NOT EXISTS trips (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                started_at TEXT NOT NULL,
                ended_at TEXT,
                direction TEXT NOT NULL,
                route_id TEXT NOT NULL,
                note TEXT NOT NULL DEFAULT '',
                departure_location TEXT,
                weekday TEXT,
                local_time TEXT,
                weather_time TEXT,
                weather_code INTEGER,
                weather_description TEXT,
                temperature_c REAL,
                relative_humidity_pct REAL,
                precipitation_mm REAL,
                rain_mm REAL,
                snowfall_cm REAL,
                wind_speed_kmh REAL,
                weather_error TEXT
            );
            CREATE INDEX IF NOT EXISTS trips_started_idx ON trips(started_at);
            """
        )
        extra_columns = {
            "departure_location": "TEXT",
            "weekday": "TEXT",
            "local_time": "TEXT",
            "collection_mode": "TEXT NOT NULL DEFAULT '정기'",
            **WEATHER_FIELDS,
        }
        for table in ("estimates", "trips"):
            existing = {row["name"] for row in conn.execute(f"PRAGMA table_info({table})")}
            for name, kind in extra_columns.items():
                if name not in existing:
                    conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {kind}")


def as_excel_datetime(value: str | None):
    if not value:
        return None
    return datetime.fromisoformat(value).replace(tzinfo=None)


def build_workbook() -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    with _db_lock, db_connect() as conn:
        estimates = [dict(row) for row in conn.execute("SELECT * FROM estimates ORDER BY observed_at")]
        trips = [dict(row) for row in conn.execute("SELECT * FROM trips ORDER BY started_at")]

    workbook = Workbook()
    estimate_sheet = workbook.active
    estimate_sheet.title = "예상시간"
    estimate_headers = [
        "조회 시각", "요일", "시간", "출발지", "수집 유형", "방향", "경로", "예상시간(분)", "거리(km)",
        "지정 도로 확인", "도로명", "날씨 기준 시각", "날씨", "기온(°C)", "습도(%)",
        "강수량(mm)", "비(mm)", "강설(cm)", "풍속(km/h)", "날씨 오류", "경로 조회 오류",
    ]
    estimate_sheet.append(estimate_headers)
    for row in estimates:
        roads = row.get("road_names") or "[]"
        try:
            roads_text = ", ".join(json.loads(roads))
        except (TypeError, json.JSONDecodeError):
            roads_text = str(roads)
        estimate_sheet.append([
            as_excel_datetime(row["observed_at"]), row.get("weekday"), row.get("local_time"), row.get("departure_location"),
            row.get("collection_mode", "정기"), row["direction"], ROUTES[row["route_id"]]["label"],
            round(row["duration_seconds"] / 60, 1) if row.get("duration_seconds") is not None else None,
            round(row["distance_meters"] / 1000, 1) if row.get("distance_meters") is not None else None,
            "확인" if row.get("route_match") else "미확인", roads_text,
            row.get("weather_time"), row.get("weather_description"), row.get("temperature_c"),
            row.get("relative_humidity_pct"), row.get("precipitation_mm"), row.get("rain_mm"),
            row.get("snowfall_cm"), row.get("wind_speed_kmh"), row.get("weather_error"), row.get("error"),
        ])

    trip_sheet = workbook.create_sheet("실제운행")
    trip_sheet.append([
        "출발 시각", "도착 시각", "요일", "시간", "출발지", "방향", "경로", "실제 소요(분)",
        "날씨 기준 시각", "날씨", "기온(°C)", "습도(%)", "강수량(mm)", "비(mm)", "강설(cm)",
        "풍속(km/h)", "날씨 오류", "메모",
    ])
    for row in trips:
        actual = None
        if row.get("ended_at"):
            actual = round((datetime.fromisoformat(row["ended_at"]) - datetime.fromisoformat(row["started_at"])).total_seconds() / 60, 1)
        trip_sheet.append([
            as_excel_datetime(row["started_at"]), as_excel_datetime(row.get("ended_at")), row.get("weekday"),
            row.get("local_time"), row.get("departure_location"), row["direction"],
            ROUTES[row["route_id"]]["label"], actual, row.get("weather_time"), row.get("weather_description"),
            row.get("temperature_c"), row.get("relative_humidity_pct"), row.get("precipitation_mm"),
            row.get("rain_mm"), row.get("snowfall_cm"), row.get("wind_speed_kmh"),
            row.get("weather_error"), row.get("note"),
        ])

    guide = workbook.create_sheet("안내")
    guide.append(["항목", "내용"])
    guide.append(["저장 위치", str(EXCEL_PATH)])
    guide.append(["자동 수집", "평일 05:00–10:00 및 16:00–20:00, 기본 10분 간격 (설정에서 변경 가능)"])
    guide.append(["요일·시간", "각 조회 시각 또는 운행 시작 시각의 PC 현지 시간 기준"])
    guide.append(["예상시간", "지도 API 조회 시점의 교통 반영 경로 시간"])
    guide.append(["실제 운행시간", "출발 버튼부터 도착 버튼까지 직접 측정한 시간; GPS 자동 추적 아님"])
    guide.append(["날씨", "Open-Meteo 현재 날씨. 동의에 따라 카카오 좌표를 전송하며, 좌표 격자 기준 자료라 실제 출발점의 관측값과 차이가 있을 수 있음"])
    guide.append(["도로 확인", "경유지·우선순위로 만든 경로의 도로명을 응답에서 확인; 미확인은 고속도로 통과가 검증되지 않은 상태"])

    header_fill = PatternFill("solid", fgColor="17365D")
    for sheet in (estimate_sheet, trip_sheet, guide):
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        sheet.row_dimensions[1].height = 32
        for row_cells in sheet.iter_rows(min_row=2):
            for cell in row_cells:
                cell.font = Font(name="Arial", size=10)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        for column_cells in sheet.columns:
            letter = get_column_letter(column_cells[0].column)
            width = min(42, max(12, max(len(str(cell.value or "")) for cell in column_cells[:100]) + 2))
            sheet.column_dimensions[letter].width = width
    for sheet, timestamp_columns in ((estimate_sheet, (1,)), (trip_sheet, (1, 2))):
        for col in timestamp_columns:
            for row in range(2, sheet.max_row + 1):
                sheet.cell(row, col).number_format = "yyyy-mm-dd hh:mm:ss"

    workbook.save(EXCEL_TEMP_PATH)
    os.replace(EXCEL_TEMP_PATH, EXCEL_PATH)


def sync_workbook() -> None:
    with _excel_lock:
        try:
            build_workbook()
        except Exception as exc:
            try:
                EXCEL_TEMP_PATH.unlink(missing_ok=True)
            except OSError:
                pass
            with _state_lock:
                _state["excel_error"] = f"Excel 저장 실패: {exc}"
            return
        with _state_lock:
            _state["excel_error"] = None




def load_settings() -> dict:
    try:
        data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        return {**DEFAULT_SETTINGS, **data}
    except (OSError, json.JSONDecodeError):
        return DEFAULT_SETTINGS.copy()


def save_settings(settings: dict) -> None:
    safe = {**DEFAULT_SETTINGS, **settings}
    SETTINGS_PATH.write_text(json.dumps(safe, ensure_ascii=False, indent=2), encoding="utf-8")


def kakao_get(url: str, api_key: str) -> dict:
    request = urllib.request.Request(url, headers={"Authorization": f"KakaoAK {api_key}"})
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")[:500]
        raise RuntimeError(f"지도 API 응답 오류 ({exc.code}): {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"지도 API 연결 실패: {exc.reason}") from exc


def geocode(address: str, api_key: str) -> tuple[float, float]:
    query = urllib.parse.urlencode({"query": address})
    data = kakao_get(f"https://dapi.kakao.com/v2/local/search/address.json?{query}", api_key)
    docs = data.get("documents", [])
    if not docs:
        # Parcel/landmark addresses may be found by Kakao keyword search instead.
        data = kakao_get(
            f"https://dapi.kakao.com/v2/local/search/keyword.json?{query}&size=5", api_key
        )
        docs = data.get("documents", [])
    if not docs:
        raise RuntimeError(f"주소를 찾지 못했습니다: {address}")
    return float(docs[0]["x"]), float(docs[0]["y"])


def local_context(direction: str, timestamp: str | None = None) -> dict:
    current = datetime.fromisoformat(timestamp) if timestamp else datetime.now().astimezone()
    return {
        "departure_location": "집" if direction == "출근" else "회사",
        "weekday": WEEKDAYS_KO[current.weekday()],
        "local_time": current.strftime("%H:%M"),
    }


def weather_description(code: int | None) -> str:
    descriptions = {
        0: "맑음", 1: "대체로 맑음", 2: "부분적으로 흐림", 3: "흐림",
        45: "안개", 48: "착빙성 안개", 51: "약한 이슬비", 53: "보통 이슬비", 55: "강한 이슬비",
        56: "약한 어는 이슬비", 57: "강한 어는 이슬비", 61: "약한 비", 63: "보통 비", 65: "강한 비",
        66: "약한 어는 비", 67: "강한 어는 비", 71: "약한 눈", 73: "보통 눈", 75: "강한 눈",
        77: "싸락눈", 80: "약한 소나기", 81: "보통 소나기", 82: "강한 소나기",
        85: "약한 눈 소나기", 86: "강한 눈 소나기", 95: "뇌우", 96: "약한 우박 동반 뇌우", 99: "강한 우박 동반 뇌우",
    }
    return descriptions.get(code, "알 수 없음")


def fetch_weather(latitude: float, longitude: float) -> dict:
    params = urllib.parse.urlencode({
        "latitude": latitude,
        "longitude": longitude,
        "current": "temperature_2m,relative_humidity_2m,precipitation,rain,snowfall,weather_code,wind_speed_10m",
        "timezone": "Asia/Seoul",
    })
    request = urllib.request.Request(
        f"https://api.open-meteo.com/v1/forecast?{params}",
        headers={"User-Agent": "CommuteTimeLogger/1.0"},
    )
    try:
        with urllib.request.urlopen(request, timeout=15) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")[:300]
        raise RuntimeError(f"날씨 API 응답 오류 ({exc.code}): {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"날씨 API 연결 실패: {exc.reason}") from exc
    current = payload.get("current")
    if not current:
        raise RuntimeError("날씨 API 응답에 현재 날씨가 없습니다.")
    code = int(current["weather_code"]) if current.get("weather_code") is not None else None
    return {
        "weather_time": current.get("time"),
        "weather_code": code,
        "weather_description": weather_description(code),
        "temperature_c": current.get("temperature_2m"),
        "relative_humidity_pct": current.get("relative_humidity_2m"),
        "precipitation_mm": current.get("precipitation"),
        "rain_mm": current.get("rain"),
        "snowfall_cm": current.get("snowfall"),
        "wind_speed_kmh": current.get("wind_speed_10m"),
        "weather_error": None,
    }


def origin_weather(direction: str, timestamp: str, settings: dict) -> dict:
    context = local_context(direction, timestamp)
    if not settings.get("api_key"):
        return {**context, "weather_error": "카카오 API 키를 등록하면 출발지 좌표를 조회할 수 있습니다."}
    try:
        coordinates = ensure_coordinates(settings)
        origin_key = "home" if direction == "출근" else "work"
        longitude, latitude = coordinates[origin_key]
        return {**context, **fetch_weather(latitude, longitude)}
    except Exception as exc:
        return {**context, "weather_error": str(exc)}


def collect_trip_weather(trip_id: int, direction: str, started_at: str) -> None:
    weather = origin_weather(direction, started_at, load_settings())
    with _db_lock, db_connect() as conn:
        conn.execute(
            """UPDATE trips SET weather_time=?,weather_code=?,weather_description=?,temperature_c=?,
               relative_humidity_pct=?,precipitation_mm=?,rain_mm=?,snowfall_cm=?,wind_speed_kmh=?,weather_error=?
               WHERE id=?""",
            (
                weather.get("weather_time"), weather.get("weather_code"), weather.get("weather_description"),
                weather.get("temperature_c"), weather.get("relative_humidity_pct"), weather.get("precipitation_mm"),
                weather.get("rain_mm"), weather.get("snowfall_cm"), weather.get("wind_speed_kmh"),
                weather.get("weather_error"), trip_id,
            ),
        )
    sync_workbook()


def ensure_coordinates(settings: dict) -> dict:
    coordinates = settings.get("coordinates") or {}
    changed = False
    for key, address in ADDRESSES.items():
        if key not in coordinates:
            coordinates[key] = geocode(address, settings["api_key"])
            changed = True
    if changed:
        settings["coordinates"] = coordinates
        save_settings(settings)
    return coordinates


def get_roads(route: dict) -> list[str]:
    names = []
    for section in route.get("sections", []):
        for road in section.get("roads", []):
            name = road.get("name")
            if name and name not in names:
                names.append(name)
    return names


def get_route_duration(route_id: str, direction: str, settings: dict) -> dict:
    coords = ensure_coordinates(settings)
    origin_key, destination_key = ("home", "work") if direction == "출근" else ("work", "home")
    origin = ",".join(map(str, coords[origin_key]))
    destination = ",".join(map(str, coords[destination_key]))
    params = {
        "origin": origin,
        "destination": destination,
        "priority": ROUTES[route_id]["priority"],
        "alternatives": "true",
        "road_details": "true",
    }
    if route_id == "via":
        params["waypoints"] = ",".join(map(str, coords["via"]))
    url = "https://apis-navi.kakaomobility.com/v1/directions?" + urllib.parse.urlencode(params)
    payload = kakao_get(url, settings["api_key"])
    choices = [r for r in payload.get("routes", []) if r.get("result_code", 0) == 0]
    if not choices:
        raise RuntimeError(payload.get("routes", [{}])[0].get("result_msg", "경로를 찾지 못했습니다."))
    if route_id == "shortest":
        chosen = min(choices, key=lambda r: r.get("summary", {}).get("distance", 10**12))
    else:
        wanted = ROUTES[route_id]["must_contain"]
        chosen = next((r for r in choices if wanted in " ".join(get_roads(r))), choices[0])
    summary = chosen.get("summary", {})
    roads = get_roads(chosen)
    matched = ROUTES[route_id]["must_contain"] in " ".join(roads)
    return {
        "duration_seconds": int(summary.get("duration", 0)),
        "distance_meters": int(summary.get("distance", 0)),
        "road_names": roads,
        "route_match": matched,
    }


def insert_estimate(direction: str, route_id: str, result: dict | None = None, error: str | None = None, collection_mode: str = "정기", weather: dict | None = None) -> None:
    observed_at = now_iso()
    observed = datetime.fromisoformat(observed_at)
    context = local_context(direction, observed_at)
    weather = weather or {}
    with _db_lock, db_connect() as conn:
        conn.execute(
            """INSERT INTO estimates(
                observed_at,direction,route_id,duration_seconds,distance_meters,road_names,route_match,
                departure_location,weekday,local_time,collection_mode,weather_time,weather_code,weather_description,
                temperature_c,relative_humidity_pct,precipitation_mm,rain_mm,snowfall_cm,wind_speed_kmh,weather_error,error
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                observed_at, direction, route_id,
                result.get("duration_seconds") if result else None,
                result.get("distance_meters") if result else None,
                json.dumps(result.get("road_names", []) if result else [], ensure_ascii=False),
                int(result.get("route_match", False)) if result else 0,
                context["departure_location"], context["weekday"], context["local_time"], collection_mode,
                weather.get("weather_time"), weather.get("weather_code"), weather.get("weather_description"),
                weather.get("temperature_c"), weather.get("relative_humidity_pct"), weather.get("precipitation_mm"),
                weather.get("rain_mm"), weather.get("snowfall_cm"), weather.get("wind_speed_kmh"),
                weather.get("weather_error"), error,
            ),
        )


def poll_routes(direction_override: str | None = None) -> None:
    settings = load_settings()
    if not settings.get("api_key"):
        with _state_lock:
            _state.update(last_error=None, last_message="API 키 설정을 기다리는 중")
        return
    now = datetime.now().astimezone()
    in_morning_window = 5 * 60 <= now.hour * 60 + now.minute <= 10 * 60
    direction = direction_override or ("출근" if in_morning_window or now.hour < 12 else "퇴근")
    with _state_lock:
        _state.update(polling=True, last_message=f"{direction} 경로 조회 중", last_error=None)
    errors = []
    weather = origin_weather(direction, now_iso(), settings)
    if weather.get("weather_error"):
        errors.append(f"날씨: {weather['weather_error']}")
    for route_id in ROUTES:
        try:
            result = get_route_duration(route_id, direction, settings)
            insert_estimate(direction, route_id, result=result, weather=weather)
        except Exception as exc:  # Keep the other route flowing when one request fails.
            message = str(exc)
            insert_estimate(direction, route_id, error=message, weather=weather)
            errors.append(f"{ROUTES[route_id]['label']}: {message}")
    with _state_lock:
        _state.update(
            polling=False,
            last_poll=now_iso(),
            last_error=" / ".join(errors) if errors else None,
            last_message="일부 경로 조회 실패" if errors else f"{direction} 예상시간 저장 완료",
        )
    sync_workbook()
    write_collection_health()


def is_poll_time(now: datetime, interval: int, weekdays_only: bool) -> bool:
    if weekdays_only and now.weekday() >= 5:
        return False
    minute_of_day = now.hour * 60 + now.minute
    in_window = (5 * 60 <= minute_of_day <= 10 * 60) or (16 * 60 <= minute_of_day <= 20 * 60)
    return in_window and now.minute % interval == 0


def within_collection_period(now: datetime, settings: dict) -> bool:
    now = now.replace(second=0, microsecond=0)
    start = settings.get("collection_start")
    end = settings.get("collection_end")
    return (not start or now >= datetime.fromisoformat(start)) and (not end or now <= datetime.fromisoformat(end))


def write_collection_health() -> None:
    """Persist source and local export health without API keys."""
    from openpyxl import load_workbook

    with _db_lock, db_connect() as conn:
        source_count = conn.execute("SELECT COUNT(*) FROM estimates").fetchone()[0]
        last = conn.execute("SELECT observed_at,direction FROM estimates WHERE collection_mode='정기' ORDER BY id DESC LIMIT 1").fetchone()
        latest = []
        if last:
            latest = [dict(row) for row in conn.execute(
                "SELECT observed_at,direction,route_id,duration_seconds,weather_description,error,weather_error,route_match FROM estimates WHERE collection_mode='정기' AND direction=? ORDER BY id DESC LIMIT 2", (last["direction"],)
            )]
    excel_count = None
    excel_error = None
    try:
        workbook = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
        try:
            excel_count = workbook["예상시간"].max_row - 1
        finally:
            workbook.close()
    except Exception as exc:
        excel_error = str(exc)
    with _state_lock:
        state = dict(_state)
    checks = {
        "local_excel_matches_database": excel_count == source_count,
        "latest_routes_successful": len(latest) == 2 and all(row["duration_seconds"] is not None and not row["error"] for row in latest),
        "latest_weather_successful": len(latest) == 2 and all(row["weather_description"] and not row["weather_error"] for row in latest),
    }
    health = {"checked_at": now_iso(), "database_rows": source_count, "excel_rows": excel_count,
              "excel_error": excel_error, "checks": checks, "latest_regular_samples": latest,
              "collection_start": load_settings().get("collection_start"), "collection_end": load_settings().get("collection_end")}
    temporary = HEALTH_PATH.with_suffix(".tmp")
    temporary.write_text(json.dumps(health, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, HEALTH_PATH)


def scheduler() -> None:
    last_slot = None
    while True:
        try:
            settings = load_settings()
            interval = max(1, min(60, int(settings.get("interval_minutes", 10))))
            now = datetime.now().astimezone()
            slot = (now.date().isoformat(), now.hour, now.minute)
            with _state_lock:
                test_running = _state["test_running"]
            if not within_collection_period(now, settings):
                with _state_lock:
                    _state["last_message"] = "예약 시작 대기" if settings.get("collection_start") and now < datetime.fromisoformat(settings["collection_start"]) else "수집 기간 종료"
            elif not test_running and is_poll_time(now, interval, bool(settings.get("weekdays_only", True))) and slot != last_slot:
                last_slot = slot
                direction = "출근" if now.hour <= 10 else "퇴근"
                poll_routes(direction)
        except Exception as exc:
            with _state_lock:
                _state.update(polling=False, last_error=str(exc), last_message="수집 오류 · 다음 주기에 계속 점검")
        time.sleep(5)


def run_one_minute_test() -> None:
    if not _test_lock.acquire(blocking=False):
        return
    try:
        settings = load_settings()
        if not settings.get("api_key"):
            with _state_lock:
                _state.update(test_running=False, last_message="테스트를 위해 카카오 API 키를 먼저 등록해 주세요.")
            return
        with _state_lock:
            _state.update(test_running=True, test_samples=0, test_error=None, last_message="출근·퇴근 경로를 5초 간격으로 1분간 조회 중")
        ensure_coordinates(settings)
        deadline = time.monotonic() + 60
        specifications = [(direction, route_id) for direction in ("출근", "퇴근") for route_id in ROUTES]
        while time.monotonic() < deadline:
            round_started = time.monotonic()
            errors = []
            route_results = {}
            route_errors = {}
            weather_by_direction = {}
            with concurrent.futures.ThreadPoolExecutor(max_workers=6) as pool:
                futures = {
                    pool.submit(get_route_duration, route_id, direction, settings): ("route", direction, route_id)
                    for direction, route_id in specifications
                }
                for direction in ("출근", "퇴근"):
                    futures[pool.submit(origin_weather, direction, now_iso(), settings)] = ("weather", direction, None)
                for future in concurrent.futures.as_completed(futures):
                    kind, direction, route_id = futures[future]
                    try:
                        value = future.result()
                        if kind == "weather":
                            weather_by_direction[direction] = value
                            if value.get("weather_error"):
                                errors.append(f"{direction} 날씨: {value['weather_error']}")
                        else:
                            route_results[(direction, route_id)] = value
                    except Exception as exc:
                        errors.append(str(exc))
                        route_errors[(direction, route_id)] = str(exc)
                for direction, route_id in specifications:
                    result = route_results.get((direction, route_id))
                    error = None if result else route_errors.get((direction, route_id), f"{ROUTES[route_id]['label']} 경로 조회 실패")
                    insert_estimate(
                        direction, route_id, result=result, error=error,
                        collection_mode="1분 테스트", weather=weather_by_direction.get(direction),
                    )
            sync_workbook()
            with _state_lock:
                _state["test_samples"] += len(specifications)
                _state["last_poll"] = now_iso()
                _state["test_error"] = " / ".join(errors) if errors else None
            remaining = 5 - (time.monotonic() - round_started)
            if remaining > 0:
                time.sleep(min(remaining, max(0, deadline - time.monotonic())))
        with _state_lock:
            _state.update(test_running=False, last_message="1분 테스트 수집 완료")
    except Exception as exc:
        with _state_lock:
            _state.update(test_running=False, test_error=str(exc), last_message="1분 테스트가 완료되지 않았습니다.")
    finally:
        _test_lock.release()


def latest_estimate_for_trip(conn, started_at: str, direction: str, route_id: str):
    target = datetime.fromisoformat(started_at)
    rows = conn.execute(
        """SELECT * FROM estimates WHERE direction=? AND route_id=? AND duration_seconds IS NOT NULL
           ORDER BY observed_at DESC LIMIT 1000""",
        (direction, route_id),
    ).fetchall()
    closest = None
    closest_delta = None
    for row in rows:
        observed = datetime.fromisoformat(row["observed_at"])
        delta = abs((observed - target).total_seconds())
        if delta <= 15 * 60 and (closest_delta is None or delta < closest_delta):
            closest, closest_delta = row, delta
    return dict(closest) if closest else None


def dashboard_data() -> dict:
    with _db_lock, db_connect() as conn:
        estimates = [dict(r) for r in conn.execute("SELECT * FROM estimates ORDER BY observed_at DESC LIMIT 300")]
        trips = [dict(r) for r in conn.execute("SELECT * FROM trips ORDER BY started_at DESC LIMIT 100")]
        estimates_total = conn.execute("SELECT COUNT(*) FROM estimates").fetchone()[0]
        trips_total = conn.execute("SELECT COUNT(*) FROM trips WHERE ended_at IS NOT NULL").fetchone()[0]
        active = conn.execute("SELECT * FROM trips WHERE ended_at IS NULL ORDER BY started_at DESC LIMIT 1").fetchone()
        matched_trips = []
        for trip in trips:
            actual_seconds = None
            if trip["ended_at"]:
                start = datetime.fromisoformat(trip["started_at"])
                end = datetime.fromisoformat(trip["ended_at"])
                actual_seconds = max(0, int((end - start).total_seconds()))
            estimate = latest_estimate_for_trip(conn, trip["started_at"], trip["direction"], trip["route_id"])
            matched_trips.append({**trip, "actual_seconds": actual_seconds, "estimate_seconds": estimate["duration_seconds"] if estimate else None})
        estimate_hourly = {}
        for estimate in conn.execute("SELECT direction,route_id,observed_at,duration_seconds,route_match FROM estimates WHERE duration_seconds IS NOT NULL AND collection_mode!='1분 테스트'"):
            observed = datetime.fromisoformat(estimate["observed_at"])
            hour = observed.hour
            key = (estimate["direction"], estimate["route_id"], hour)
            estimate_hourly.setdefault(key, []).append((estimate["duration_seconds"], observed.date().isoformat(), bool(estimate["route_match"])))
        actual_hourly = {}
        for trip in matched_trips:
            if trip["actual_seconds"] is None:
                continue
            hour = datetime.fromisoformat(trip["started_at"]).hour
            key = (trip["direction"], trip["route_id"], hour)
            actual_hourly.setdefault(key, []).append((trip["actual_seconds"], datetime.fromisoformat(trip["started_at"]).date().isoformat()))
        stats = []
        for key in sorted(set(estimate_hourly) | set(actual_hourly)):
            estimates_for_hour = estimate_hourly.get(key, [])
            direction, route_id, hour = key
            actual_values = actual_hourly.get(key, [])
            stats.append({
                "direction": direction,
                "route_id": route_id,
                "hour": hour,
                "samples": len(estimates_for_hour),
                "estimate_days": len({day for _, day, _ in estimates_for_hour}),
                "avg_seconds": round(sum(seconds for seconds, _, _ in estimates_for_hour) / len(estimates_for_hour)) if estimates_for_hour else None,
                "verified_samples": sum(1 for _, _, matched in estimates_for_hour if matched),
                "verified_days": len({day for _, day, matched in estimates_for_hour if matched}),
                "verified_avg_seconds": round(sum(seconds for seconds, _, matched in estimates_for_hour if matched) / sum(1 for _, _, matched in estimates_for_hour if matched)) if any(matched for _, _, matched in estimates_for_hour) else None,
                "actual_samples": len(actual_values),
                "actual_days": len({day for _, day in actual_values}),
                "actual_avg_seconds": round(sum(seconds for seconds, _ in actual_values) / len(actual_values)) if actual_values else None,
            })
    settings = load_settings()
    return {
        "settings": {"api_key_configured": bool(settings.get("api_key")), "interval_minutes": settings.get("interval_minutes", 10), "weekdays_only": settings.get("weekdays_only", True), "coordinates_ready": len(settings.get("coordinates", {})) == 3, "collection_start": settings.get("collection_start"), "collection_end": settings.get("collection_end")},
        "state": dict(_state),
        "estimates": estimates,
        "estimates_total": estimates_total,
        "trips": matched_trips,
        "trips_total": trips_total,
        "active_trip": dict(active) if active else None,
        "stats": stats,
    }


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        # Avoid writing request bodies or secrets to console logs.
        print(f"[{self.log_date_time_string()}] {fmt % args}")

    def send_json(self, data: dict, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def read_json(self):
        size = int(self.headers.get("Content-Length", "0"))
        return json.loads(self.rfile.read(size).decode("utf-8")) if size else {}

    def do_GET(self):
        if self.path == "/api/dashboard":
            self.send_json(dashboard_data())
        elif self.path == "/api/export.csv":
            self.export_csv()
        elif self.path == "/api/export.xlsx":
            sync_workbook()
            if not EXCEL_PATH.exists():
                return self.send_json({"error": "Excel 파일을 만들지 못했습니다."}, 500)
            body = EXCEL_PATH.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", "attachment; filename=commute_data.xlsx")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif self.path in ("/", "/index.html"):
            body = (ROOT / "index.html").read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        else:
            self.send_error(404)

    def do_POST(self):
        try:
            data = self.read_json()
            if self.path == "/api/settings":
                settings = load_settings()
                if data.get("api_key"):
                    settings["api_key"] = data["api_key"].strip()
                    settings["coordinates"] = {}
                settings["interval_minutes"] = max(1, min(60, int(data.get("interval_minutes", 10))))
                settings["weekdays_only"] = bool(data.get("weekdays_only", True))
                save_settings(settings)
                sync_workbook()
                self.send_json({"ok": True, "message": "설정을 저장했습니다. 좌표 조회는 다음 수집 때 진행합니다."})
            elif self.path == "/api/trips/start":
                direction, route_id = data.get("direction"), data.get("route_id")
                if direction not in ("출근", "퇴근") or route_id not in ROUTES:
                    return self.send_json({"error": "출퇴근 방향과 경로를 확인해 주세요."}, 400)
                with _db_lock, db_connect() as conn:
                    active = conn.execute("SELECT id FROM trips WHERE ended_at IS NULL LIMIT 1").fetchone()
                    if active:
                        return self.send_json({"error": "진행 중인 운행 기록을 먼저 종료해 주세요."}, 409)
                    started_at = now_iso()
                    started = datetime.fromisoformat(started_at)
                    departure_location = "집" if direction == "출근" else "회사"
                    cur = conn.execute(
                        """INSERT INTO trips(started_at,direction,route_id,note,departure_location,weekday,local_time)
                           VALUES(?,?,?,?,?,?,?)""",
                        (started_at, direction, route_id, str(data.get("note", ""))[:300], departure_location,
                         WEEKDAYS_KO[started.weekday()], started.strftime("%H:%M")),
                    )
                    trip_id = cur.lastrowid
                sync_workbook()
                threading.Thread(target=collect_trip_weather, args=(trip_id, direction, started_at), daemon=True).start()
                self.send_json({"ok": True, "trip_id": trip_id})
            elif self.path == "/api/trips/finish":
                with _db_lock, db_connect() as conn:
                    active = conn.execute("SELECT id FROM trips WHERE ended_at IS NULL ORDER BY started_at DESC LIMIT 1").fetchone()
                    if not active:
                        return self.send_json({"error": "종료할 운행 기록이 없습니다."}, 404)
                    conn.execute("UPDATE trips SET ended_at=?, note=CASE WHEN ?='' THEN note ELSE ? END WHERE id=?", (now_iso(), str(data.get("note", ""))[:300], str(data.get("note", ""))[:300], active["id"]))
                sync_workbook()
                self.send_json({"ok": True})
            elif self.path == "/api/poll-now":
                direction = data.get("direction")
                if direction not in ("출근", "퇴근"):
                    return self.send_json({"error": "조회 방향을 선택해 주세요."}, 400)
                threading.Thread(target=poll_routes, args=(direction,), daemon=True).start()
                self.send_json({"ok": True})
            elif self.path == "/api/test-start":
                if not load_settings().get("api_key"):
                    return self.send_json({"error": "카카오 API 키를 먼저 등록해 주세요."}, 400)
                if _state["test_running"]:
                    return self.send_json({"error": "1분 테스트가 이미 진행 중입니다."}, 409)
                threading.Thread(target=run_one_minute_test, daemon=True).start()
                self.send_json({"ok": True}, 202)
            else:
                self.send_error(404)
        except (ValueError, json.JSONDecodeError) as exc:
            self.send_json({"error": str(exc)}, 400)
        except Exception as exc:
            self.send_json({"error": str(exc)}, 500)

    def export_csv(self):
        with _db_lock, db_connect() as conn:
            estimates = conn.execute("SELECT * FROM estimates ORDER BY observed_at").fetchall()
            trips = conn.execute("SELECT * FROM trips ORDER BY started_at").fetchall()
        out = io.StringIO()
        writer = csv.writer(out)
        writer.writerow(["종류", "시작/조회시각", "종료시각", "요일", "시간", "출발지", "방향", "경로", "예상시간(분)", "실제시간(분)", "거리(km)", "경로확인", "도로명", "날씨", "기온(°C)", "습도(%)", "강수량(mm)", "풍속(km/h)", "날씨 오류", "경로 오류", "메모"])
        for row in estimates:
            writer.writerow(["예상", row["observed_at"], "", row["weekday"], row["local_time"], row["departure_location"], row["direction"], ROUTES[row["route_id"]]["label"], round(row["duration_seconds"] / 60, 1) if row["duration_seconds"] is not None else "", "", round(row["distance_meters"] / 1000, 1) if row["distance_meters"] is not None else "", "확인" if row["route_match"] else "미확인", row["road_names"], row["weather_description"], row["temperature_c"], row["relative_humidity_pct"], row["precipitation_mm"], row["wind_speed_kmh"], row["weather_error"], row["error"] or "", ""])
        for row in trips:
            actual = (datetime.fromisoformat(row["ended_at"]) - datetime.fromisoformat(row["started_at"])).total_seconds() / 60 if row["ended_at"] else ""
            writer.writerow(["실제", row["started_at"], row["ended_at"] or "", row["weekday"], row["local_time"], row["departure_location"], row["direction"], ROUTES[row["route_id"]]["label"], "", round(actual, 1) if actual != "" else "", "", "", "", row["weather_description"], row["temperature_c"], row["relative_humidity_pct"], row["precipitation_mm"], row["wind_speed_kmh"], row["weather_error"], "", row["note"]])
        body = out.getvalue().encode("utf-8-sig")
        self.send_response(200)
        self.send_header("Content-Type", "text/csv; charset=utf-8")
        self.send_header("Content-Disposition", "attachment; filename=commute-records.csv")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main():
    init_db()
    sync_workbook()
    write_collection_health()
    threading.Thread(target=scheduler, daemon=True).start()
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"통근 기록기 실행: http://{HOST}:{PORT}")
    print("종료하려면 이 창에서 Ctrl+C를 누르세요.")
    import webbrowser
    webbrowser.open(f"http://{HOST}:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n통근 기록기를 종료합니다.")
    finally:
        server.server_close()


if __name__ == "__main__":
    import sys
    if "--check" in sys.argv:
        import openpyxl
        print(f"Local collector ready; Python {sys.version.split()[0]}; openpyxl {openpyxl.__version__}")
    else:
        main()
