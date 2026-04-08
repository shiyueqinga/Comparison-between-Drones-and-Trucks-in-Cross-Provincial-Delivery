from pathlib import Path
import math
import random
import subprocess
import pandas as pd
import osmnx as ox
import networkx as nx

# =========================
# 0. 路径与参数
# =========================
BASE_DIR = Path(r"C:\Users\wlmlt\Desktop\软件\pythonProject\低空经济")

EXCEL_PATH = BASE_DIR / "最新的快递量分析_补充蜂巢城市经纬度.xlsx"
SHEET_NAME = "Sheet1"

OSM_PBF_PATH = BASE_DIR / "china-260401.osm.pbf"
OSMIUM_CMD = r"C:\Users\wlmlt\miniconda3\envs\osmium_env\Library\bin\osmium.exe"

CLIP_CACHE_DIR = BASE_DIR / "offline_osm_clips_osmium"
GRAPH_CACHE_DIR = BASE_DIR / "offline_graph_cache_osmium"

POINT_JITTER_KM = 8
BBOX_BUFFER_DEG_LIST = [0.60]
RANDOM_SEED = 42
random.seed(RANDOM_SEED)

# 新增：距离分段步长（用于输出“多少公里时无人机更快”）
DISTANCE_BIN_STEP_KM = 20

# 无人机参数
DRONE_CONFIG = {
    20: {"max_range_km": 60, "cruise_speed_kmh": 90},
    30: {"max_range_km": 30, "cruise_speed_kmh": 90},
}
DRONE_HANDLING_MIN = 8
DRONE_CHARGE_MIN = 60
TAKEOFF_LANDING_MIN = 2

# 货车参数
TRUCK_TANK_L = 120.0
TRUCK_FUEL_CONSUMPTION_L_PER_100KM = (15.0 + 9.0 + 13.0) / 3.0
TRUCK_SINGLE_TANK_RANGE_KM = TRUCK_TANK_L / TRUCK_FUEL_CONSUMPTION_L_PER_100KM * 100.0
TRUCK_PICKUP_MIN = 25
TRUCK_DELIVERY_MIN = 25
TRUCK_REFUEL_MIN = 30
TRUCK_REST_PER_DAY_MIN = 60

EXCLUDE_HIGHWAY = {
    "footway", "cycleway", "path", "steps", "pedestrian",
    "bridleway", "corridor", "track", "proposed", "construction"
}

# 内存缓存
GRAPH_MEMORY_CACHE = {}

# 控制日志
VERBOSE = False


# =========================
# 1. 通用工具
# =========================
def log(*args):
    if VERBOSE:
        print(*args)


def ensure_dirs():
    CLIP_CACHE_DIR.mkdir(exist_ok=True)
    GRAPH_CACHE_DIR.mkdir(exist_ok=True)


def geo_distance_km(p1, p2):
    lat1, lon1 = map(math.radians, p1)
    lat2, lon2 = map(math.radians, p2)
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return 6371.0 * 2 * math.asin(math.sqrt(a))


def destination_point(lat, lon, bearing_deg, distance_km):
    R = 6371.0
    bearing = math.radians(bearing_deg)
    lat1 = math.radians(lat)
    lon1 = math.radians(lon)
    d_div_r = distance_km / R

    lat2 = math.asin(
        math.sin(lat1) * math.cos(d_div_r) +
        math.cos(lat1) * math.sin(d_div_r) * math.cos(bearing)
    )

    lon2 = lon1 + math.atan2(
        math.sin(bearing) * math.sin(d_div_r) * math.cos(lat1),
        math.cos(d_div_r) - math.sin(lat1) * math.sin(lat2)
    )
    return math.degrees(lat2), math.degrees(lon2)


def interpolate_points(start, end, n_legs):
    if n_legs <= 1:
        return []

    lat1, lon1 = start
    lat2, lon2 = end
    relay_points = []

    for i in range(1, n_legs):
        frac = i / n_legs
        lat = lat1 + (lat2 - lat1) * frac
        lon = lon1 + (lon2 - lon1) * frac
        relay_points.append((lat, lon))

    return relay_points


def format_point_address(point, city_name, point_type):
    lat, lon = point
    return f"{city_name}附近随机{point_type}（{lat:.4f}, {lon:.4f}）"


def km_to_deg_lat(km):
    return km / 111.0


def km_to_deg_lon(km, lat_deg):
    return km / (111.0 * max(0.2, math.cos(math.radians(lat_deg))))


def trip_cache_key(start_row, end_row, buffer_deg):
    s = f"{start_row['蜂巢城市']}_{start_row['城市']}_{end_row['城市']}_{buffer_deg}"
    return str(abs(hash(s)))


def build_bbox_for_trip_from_rows(start_row, end_row, buffer_deg, jitter_km=POINT_JITTER_KM):
    hive_lat = float(start_row["蜂巢城市中心纬度"])
    hive_lon = float(start_row["蜂巢城市中心经度"])

    start_lat = float(start_row["纬度"])
    start_lon = float(start_row["经度"])

    end_lat = float(end_row["纬度"])
    end_lon = float(end_row["经度"])

    points = [
        (hive_lat, hive_lon),
        (start_lat, start_lon),
        (end_lat, end_lon),
    ]

    lats = [p[0] for p in points]
    lons = [p[1] for p in points]

    mid_lat = sum(lats) / len(lats)

    jitter_lat_deg = km_to_deg_lat(jitter_km) + 0.03
    jitter_lon_deg = km_to_deg_lon(jitter_km, mid_lat) + 0.03

    left = min(lons) - buffer_deg - jitter_lon_deg
    right = max(lons) + buffer_deg + jitter_lon_deg
    bottom = min(lats) - buffer_deg - jitter_lat_deg
    top = max(lats) + buffer_deg + jitter_lat_deg

    return (left, bottom, right, top)


def clip_cache_paths(key):
    pbf_path = CLIP_CACHE_DIR / f"clip_{key}.osm.pbf"
    osm_path = CLIP_CACHE_DIR / f"clip_{key}.osm"
    return pbf_path, osm_path


def graph_cache_path(key):
    return GRAPH_CACHE_DIR / f"graph_{key}.graphml"


# =========================
# 2. 读取 Sheet1
# =========================
def load_sheet1(file_path, sheet_name):
    if not file_path.exists():
        raise FileNotFoundError(f"找不到 Excel：{file_path}")

    df = pd.read_excel(file_path, sheet_name=sheet_name)
    required_cols = ["城市", "纬度", "经度", "蜂巢城市", "蜂巢城市中心纬度", "蜂巢城市中心经度"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Sheet1 缺少列：{missing}")

    df = df[required_cols].copy()
    df = df.dropna(subset=required_cols).reset_index(drop=True)
    return df


def random_point_near_city(city_row, max_radius_km=8):
    radius = random.uniform(0.5, max_radius_km)
    bearing = random.uniform(0, 360)
    return destination_point(float(city_row["纬度"]), float(city_row["经度"]), bearing, radius)


def sample_start_end(sheet1_df):
    start_idx, end_idx = random.sample(range(len(sheet1_df)), 2)
    start_row = sheet1_df.iloc[start_idx].copy()
    end_row = sheet1_df.iloc[end_idx].copy()

    start_point = random_point_near_city(start_row, POINT_JITTER_KM)
    end_point = random_point_near_city(end_row, POINT_JITTER_KM)

    start_hive_point = (
        float(start_row["蜂巢城市中心纬度"]),
        float(start_row["蜂巢城市中心经度"])
    )
    return start_row, end_row, start_hive_point, start_point, end_point


# =========================
# 3. 离线裁切（osmium）
# =========================
def run_command(cmd, cwd=None):
    result = subprocess.run(cmd, capture_output=True, text=True, cwd=cwd)
    if result.returncode != 0:
        raise RuntimeError(
            f"命令执行失败：{' '.join(str(x) for x in cmd)}\n"
            f"stdout:\n{result.stdout}\n"
            f"stderr:\n{result.stderr}"
        )
    return result


def clip_osm_by_bbox_osmium(bbox, cache_key):
    ensure_dirs()

    if not OSM_PBF_PATH.exists():
        raise FileNotFoundError(f"找不到全国 OSM PBF：{OSM_PBF_PATH}")

    pbf_clip_path, osm_clip_path = clip_cache_paths(cache_key)

    if osm_clip_path.exists() and osm_clip_path.stat().st_size > 1024:
        return osm_clip_path

    if pbf_clip_path.exists():
        pbf_clip_path.unlink(missing_ok=True)
    if osm_clip_path.exists():
        osm_clip_path.unlink(missing_ok=True)

    left, bottom, right, top = bbox
    bbox_str = f"{left},{bottom},{right},{top}"

    cmd_extract = [
        OSMIUM_CMD,
        "extract",
        "--bbox", bbox_str,
        "-o", str(pbf_clip_path),
        str(OSM_PBF_PATH)
    ]
    log("开始裁切 PBF ...")
    run_command(cmd_extract, cwd=BASE_DIR)

    if not pbf_clip_path.exists() or pbf_clip_path.stat().st_size < 1024:
        raise RuntimeError(f"osmium extract 输出为空：{pbf_clip_path}")

    cmd_cat = [
        OSMIUM_CMD,
        "cat",
        "-o", str(osm_clip_path),
        "-f", "osm",
        str(pbf_clip_path)
    ]
    log("开始转 OSM XML ...")
    run_command(cmd_cat, cwd=BASE_DIR)

    if not osm_clip_path.exists() or osm_clip_path.stat().st_size < 1024:
        raise RuntimeError(f"osmium cat 输出为空：{osm_clip_path}")

    return osm_clip_path


# =========================
# 4. 建图
# =========================
def build_graph_from_osm_xml(xml_path: Path):
    G = ox.graph_from_xml(
        filepath=xml_path,
        bidirectional=False,
        simplify=True,
        retain_all=True,
        encoding="utf-8",
    )

    G = ox.routing.add_edge_speeds(G, fallback=40)

    for u, v, k, data in G.edges(keys=True, data=True):
        if data.get("length") is None or pd.isna(data.get("length")):
            try:
                y1 = G.nodes[u]["y"]
                x1 = G.nodes[u]["x"]
                y2 = G.nodes[v]["y"]
                x2 = G.nodes[v]["x"]
                data["length"] = float(ox.distance.great_circle(y1, x1, y2, x2))
            except Exception:
                data["length"] = 1.0

    for u, v, k, data in G.edges(keys=True, data=True):
        if data.get("speed_kph") is None or pd.isna(data.get("speed_kph")):
            data["speed_kph"] = 40.0

    G = ox.routing.add_edge_travel_times(G)
    return G


def filter_car_graph(G):
    G_car = nx.MultiDiGraph()
    G_car.graph = G.graph.copy()

    for u, v, k, data in G.edges(keys=True, data=True):
        highway = data.get("highway")
        if highway is None:
            continue

        if isinstance(highway, list):
            highway_values = set(highway)
        else:
            highway_values = {highway}

        if highway_values.issubset(EXCLUDE_HIGHWAY):
            continue

        if u not in G_car:
            G_car.add_node(u, **G.nodes[u])
        if v not in G_car:
            G_car.add_node(v, **G.nodes[v])

        G_car.add_edge(u, v, key=k, **data)

    isolated_nodes = [n for n in G_car.nodes if G_car.degree(n) == 0]
    G_car.remove_nodes_from(isolated_nodes)

    return G_car


def load_graph_for_trip(start_row, end_row):
    ensure_dirs()
    last_error = None

    for buffer_deg in BBOX_BUFFER_DEG_LIST:
        cache_key = trip_cache_key(start_row, end_row, buffer_deg)
        bbox = build_bbox_for_trip_from_rows(start_row, end_row, buffer_deg)
        g_cache = graph_cache_path(cache_key)

        try:
            if cache_key in GRAPH_MEMORY_CACHE:
                return GRAPH_MEMORY_CACHE[cache_key], bbox

            if g_cache.exists() and g_cache.stat().st_size > 1024:
                G_car = ox.load_graphml(g_cache)
            else:
                clip_xml_path = clip_osm_by_bbox_osmium(bbox, cache_key)
                G = build_graph_from_osm_xml(clip_xml_path)
                G_car = filter_car_graph(G)

                if len(G_car.nodes) == 0 or len(G_car.edges) == 0:
                    raise RuntimeError("过滤后汽车图为空")

                ox.save_graphml(G_car, g_cache)

            GRAPH_MEMORY_CACHE[cache_key] = G_car
            return G_car, bbox

        except Exception as e:
            last_error = e
            continue

    raise RuntimeError(f"离线路网建图失败，最后一次错误：{last_error}")


def route_leg_stats(G_car, p_from, p_to):
    orig_node = ox.distance.nearest_nodes(G_car, X=p_from[1], Y=p_from[0])
    dest_node = ox.distance.nearest_nodes(G_car, X=p_to[1], Y=p_to[0])

    route = ox.routing.shortest_path(G_car, orig_node, dest_node, weight="travel_time")
    if route is None or len(route) <= 1:
        raise RuntimeError("未找到可通行路线")

    total_length_m = 0.0
    total_time_s = 0.0

    for u, v in zip(route[:-1], route[1:]):
        edges = G_car.get_edge_data(u, v)
        if not edges:
            continue

        best_edge = min(
            edges.values(),
            key=lambda d: d.get("travel_time", float("inf"))
        )
        total_length_m += best_edge.get("length", 0.0)
        total_time_s += best_edge.get("travel_time", 0.0)

    return total_length_m / 1000.0, total_time_s / 3600.0


# =========================
# 5. 无人机与货车
# =========================
def calc_leg_charges(distance_km, max_range_km):
    if distance_km <= 0:
        return 0
    return max(0, math.ceil(distance_km / max_range_km) - 1)


def simulate_drone(start_hive_point, start_point, end_point, start_row, end_row, weight_kg):
    if weight_kg not in DRONE_CONFIG:
        raise ValueError(f"不支持重量：{weight_kg}kg")

    max_range_km = DRONE_CONFIG[weight_kg]["max_range_km"]
    cruise_speed_kmh = DRONE_CONFIG[weight_kg]["cruise_speed_kmh"]

    leg1_km = geo_distance_km(start_hive_point, start_point)
    leg2_km = geo_distance_km(start_point, end_point)

    charge_count_leg1 = calc_leg_charges(leg1_km, max_range_km)
    charge_count_leg2 = calc_leg_charges(leg2_km, max_range_km)
    charge_count_total = charge_count_leg1 + charge_count_leg2

    relay_leg1 = interpolate_points(start_hive_point, start_point, max(1, math.ceil(leg1_km / max_range_km)))
    relay_leg2 = interpolate_points(start_point, end_point, max(1, math.ceil(leg2_km / max_range_km)))

    flight_time_h = (leg1_km + leg2_km) / cruise_speed_kmh
    charge_time_h = charge_count_total * (DRONE_CHARGE_MIN / 60)
    handling_time_h = 2 * (DRONE_HANDLING_MIN / 60)
    takeoff_landing_h = (
        (max(1, math.ceil(leg1_km / max_range_km)) + max(1, math.ceil(leg2_km / max_range_km)))
        * (TAKEOFF_LANDING_MIN / 60)
    )

    total_time_h = flight_time_h + charge_time_h + handling_time_h + takeoff_landing_h

    relay_addresses = []
    for i, p in enumerate(relay_leg1, 1):
        relay_addresses.append(f"去起点中继点{i}（{p[0]:.4f}, {p[1]:.4f}）")
    for i, p in enumerate(relay_leg2, 1):
        relay_addresses.append(f"送货中继点{i}（{p[0]:.4f}, {p[1]:.4f}）")

    return {
        "货物重量(kg)": weight_kg,
        "起点城市": str(start_row["城市"]),
        "终点城市": str(end_row["城市"]),
        "起始蜂巢城市": str(start_row["蜂巢城市"]),
        "起始蜂巢地址": f"{start_row['蜂巢城市']}中心位置（{start_hive_point[0]:.4f}, {start_hive_point[1]:.4f}）",
        "起点地址": format_point_address(start_point, str(start_row["城市"]), "起点"),
        "终点地址": format_point_address(end_point, str(end_row["城市"]), "终点"),
        "蜂巢到起点距离(km)": round(leg1_km, 2),
        "起点到终点距离(km)": round(leg2_km, 2),
        "无人机总飞行距离(km)": round(leg1_km + leg2_km, 2),
        "充电次数": int(charge_count_total),
        "中继点数量": len(relay_addresses),
        "中继点地址列表": " | ".join(relay_addresses) if relay_addresses else "无",
        "无人机飞行时间(h)": round(flight_time_h, 3),
        "无人机充电时间(h)": round(charge_time_h, 3),
        "无人机总运输时间(h)": round(total_time_h, 3),
    }


def calc_truck_refuels(total_road_km):
    if total_road_km <= 0:
        return 0
    return max(0, math.ceil(total_road_km / TRUCK_SINGLE_TANK_RANGE_KM) - 1)


def simulate_truck(start_hive_point, start_point, end_point, start_row, end_row, G_car, bbox):
    road_km_1, road_drive_h_1 = route_leg_stats(G_car, start_hive_point, start_point)
    road_km_2, road_drive_h_2 = route_leg_stats(G_car, start_point, end_point)

    total_road_km = road_km_1 + road_km_2
    total_drive_h = road_drive_h_1 + road_drive_h_2

    refuel_count = calc_truck_refuels(total_road_km)
    refuel_time_h = refuel_count * (TRUCK_REFUEL_MIN / 60)

    base_time_h = total_drive_h + (TRUCK_PICKUP_MIN + TRUCK_DELIVERY_MIN) / 60 + refuel_time_h
    rest_count = int(base_time_h // 24)
    rest_time_h = rest_count * (TRUCK_REST_PER_DAY_MIN / 60)

    total_time_h = base_time_h + rest_time_h

    return {
        "起点城市": str(start_row["城市"]),
        "终点城市": str(end_row["城市"]),
        "起始蜂巢城市": str(start_row["蜂巢城市"]),
        "起始蜂巢地址": f"{start_row['蜂巢城市']}中心位置（{start_hive_point[0]:.4f}, {start_hive_point[1]:.4f}）",
        "起点地址": format_point_address(start_point, str(start_row["城市"]), "起点"),
        "终点地址": format_point_address(end_point, str(end_row["城市"]), "终点"),
        "蜂巢到起点公路距离(km)": round(road_km_1, 2),
        "起点到终点公路距离(km)": round(road_km_2, 2),
        "货车总公路距离(km)": round(total_road_km, 2),
        "单箱油续航(km)": round(TRUCK_SINGLE_TANK_RANGE_KM, 2),
        "加油次数": int(refuel_count),
        "加油时间(h)": round(refuel_time_h, 3),
        "休息次数": int(rest_count),
        "休息时间(h)": round(rest_time_h, 3),
        "货车纯行驶时间(h)": round(total_drive_h, 3),
        "货车总运输时间(h)": round(total_time_h, 3),
        "汽车bbox": str(bbox),
    }


# =========================
# 6. 批量模拟提速版
# =========================
def simulate_compare_many_fast(sheet1_df, n=50):
    if len(sheet1_df) < 2:
        raise RuntimeError("城市数量不足，无法批量模拟。")

    # 第一步：先生成这一批订单
    orders = []
    pair_groups = {}

    for i in range(n):
        start_row, end_row, start_hive_point, start_point, end_point = sample_start_end(sheet1_df)

        pair_key = (str(start_row["蜂巢城市"]), str(start_row["城市"]), str(end_row["城市"]))

        order = {
            "订单编号": i + 1,
            "start_row": start_row,
            "end_row": end_row,
            "start_hive_point": start_hive_point,
            "start_point": start_point,
            "end_point": end_point,
            "pair_key": pair_key
        }
        orders.append(order)

        if pair_key not in pair_groups:
            pair_groups[pair_key] = []
        pair_groups[pair_key].append(order)

    print(f"本次共生成 {n} 单，涉及 {len(pair_groups)} 组城市组合。")

    # 第二步：预加载每组城市组合的图
    graph_map = {}
    for idx, (pair_key, group_orders) in enumerate(pair_groups.items(), start=1):
        sample_order = group_orders[0]
        start_row = sample_order["start_row"]
        end_row = sample_order["end_row"]

        print(f"预加载图 {idx}/{len(pair_groups)}: {pair_key}")
        G_car, bbox = load_graph_for_trip(start_row, end_row)
        graph_map[pair_key] = (G_car, bbox)

    # 第三步：逐单模拟，但图已经预加载
    all_rows = []
    for idx, order in enumerate(orders, start=1):
        if idx % 10 == 0 or idx == n:
            print(f"正在计算订单 {idx}/{n}")

        start_row = order["start_row"]
        end_row = order["end_row"]
        start_hive_point = order["start_hive_point"]
        start_point = order["start_point"]
        end_point = order["end_point"]
        pair_key = order["pair_key"]

        G_car, bbox = graph_map[pair_key]

        try:
            truck_result = simulate_truck(
                start_hive_point, start_point, end_point,
                start_row, end_row, G_car, bbox
            )

            for weight in [20, 30]:
                drone_result = simulate_drone(
                    start_hive_point, start_point, end_point,
                    start_row, end_row, weight
                )

                drone_time = drone_result["无人机总运输时间(h)"]
                truck_time = truck_result["货车总运输时间(h)"]
                drone_distance = drone_result["无人机总飞行距离(km)"]
                truck_distance = truck_result["货车总公路距离(km)"]

                row = {
                    "订单编号": order["订单编号"],
                    "起始蜂巢城市": drone_result["起始蜂巢城市"],
                    "起点城市": drone_result["起点城市"],
                    "终点城市": drone_result["终点城市"],
                    "货物重量(kg)": weight,
                    "起始蜂巢地址": drone_result["起始蜂巢地址"],
                    "起点地址": drone_result["起点地址"],
                    "终点地址": drone_result["终点地址"],
                    "蜂巢到起点距离(km)": drone_result["蜂巢到起点距离(km)"],
                    "起点到终点距离(km)": drone_result["起点到终点距离(km)"],
                    "无人机总飞行距离(km)": drone_distance,
                    "充电次数": drone_result["充电次数"],
                    "中继点数量": drone_result["中继点数量"],
                    "中继点地址列表": drone_result["中继点地址列表"],
                    "无人机飞行时间(h)": drone_result["无人机飞行时间(h)"],
                    "无人机充电时间(h)": drone_result["无人机充电时间(h)"],
                    "无人机总运输时间(h)": drone_time,
                    "货车总公路距离(km)": truck_distance,
                    "路网绕行系数(货车/无人机)": round(truck_distance / drone_distance, 3) if drone_distance > 0 else None,
                    "加油次数": truck_result["加油次数"],
                    "加油时间(h)": truck_result["加油时间(h)"],
                    "休息次数": truck_result["休息次数"],
                    "休息时间(h)": truck_result["休息时间(h)"],
                    "货车纯行驶时间(h)": truck_result["货车纯行驶时间(h)"],
                    "货车总运输时间(h)": truck_time,
                    "时间差(货车-无人机, h)": round(truck_time - drone_time, 3),
                    "无人机比货车快(小时)": round(truck_time - drone_time, 3),
                    "是否无人机更快": "是" if drone_time < truck_time else "否",
                }
                all_rows.append(row)

        except Exception as e:
            print(f"订单 {order['订单编号']} 失败：{e}")
            continue

    return pd.DataFrame(all_rows)


# =========================
# 7. 汇总函数
# =========================
def summarize_batch_result(batch_df):
    if batch_df.empty:
        return pd.DataFrame()

    summary = batch_df.groupby("货物重量(kg)").agg(
        平均无人机时间_h=("无人机总运输时间(h)", "mean"),
        平均货车时间_h=("货车总运输时间(h)", "mean"),
        平均快多少_h=("无人机比货车快(小时)", "mean"),
        无人机更快占比=("是否无人机更快", lambda s: (s == "是").mean()),
        平均无人机距离_km=("无人机总飞行距离(km)", "mean"),
        平均货车距离_km=("货车总公路距离(km)", "mean"),
        平均充电次数=("充电次数", "mean"),
        平均加油次数=("加油次数", "mean"),
    ).reset_index()

    return summary


def add_distance_bin(batch_df, step_km=DISTANCE_BIN_STEP_KM):
    df = batch_df.copy()
    if df.empty:
        df["距离区间"] = pd.Series(dtype="object")
        return df

    max_distance = df["无人机总飞行距离(km)"].max()
    upper = max(step_km, int(math.ceil(max_distance / step_km) * step_km))
    bins = list(range(0, upper + step_km, step_km))
    labels = [f"{bins[i]}-{bins[i+1]}km" for i in range(len(bins) - 1)]

    df["距离区间"] = pd.cut(
        df["无人机总飞行距离(km)"],
        bins=bins,
        labels=labels,
        right=False,
        include_lowest=True,
    )
    return df


def summarize_distance_advantage(batch_df, step_km=DISTANCE_BIN_STEP_KM):
    df = add_distance_bin(batch_df, step_km=step_km)
    if df.empty:
        return pd.DataFrame()

    distance_summary = df.groupby(["货物重量(kg)", "距离区间"], observed=False).agg(
        样本数=("订单编号", "count"),
        平均无人机距离_km=("无人机总飞行距离(km)", "mean"),
        平均货车距离_km=("货车总公路距离(km)", "mean"),
        平均无人机时间_h=("无人机总运输时间(h)", "mean"),
        平均货车时间_h=("货车总运输时间(h)", "mean"),
        平均无人机快多少_h=("无人机比货车快(小时)", "mean"),
        无人机更快占比=("是否无人机更快", lambda s: (s == "是").mean()),
    ).reset_index()

    distance_summary = distance_summary[distance_summary["样本数"] > 0].copy()
    distance_summary["区间比较结论"] = distance_summary["平均无人机快多少_h"].apply(
        lambda x: "无人机更快" if x > 0 else "货车更快"
    )
    distance_summary["无人机更快占比"] = distance_summary["无人机更快占比"].round(3)
    return distance_summary


def build_weight_conclusion(distance_summary_df):
    if distance_summary_df.empty:
        return pd.DataFrame(columns=[
            "货物重量(kg)", "结论文本", "无人机更快的距离区间", "货车更快的距离区间",
            "无人机仍占优的最远平均距离(km)", "优势最明显区间", "该区间平均快多少(h)"
        ])

    rows = []
    for weight, grp in distance_summary_df.groupby("货物重量(kg)"):
        grp = grp.sort_values("平均无人机距离_km").reset_index(drop=True)

        faster_grp = grp[grp["平均无人机快多少_h"] > 0]
        slower_grp = grp[grp["平均无人机快多少_h"] <= 0]

        faster_bins = "、".join(faster_grp["距离区间"].astype(str).tolist()) if not faster_grp.empty else "无"
        slower_bins = "、".join(slower_grp["距离区间"].astype(str).tolist()) if not slower_grp.empty else "无"

        farthest_adv_km = round(faster_grp["平均无人机距离_km"].max(), 2) if not faster_grp.empty else None

        best_row = grp.loc[grp["平均无人机快多少_h"].idxmax()]
        best_bin = str(best_row["距离区间"])
        best_adv_h = round(float(best_row["平均无人机快多少_h"]), 3)

        if faster_grp.empty:
            conclusion_text = (
                f"{weight}kg情况下，在本次样本中，所有距离区间内货车平均都更快，"
                f"无人机没有出现明显的时间优势。"
            )
        else:
            conclusion_text = (
                f"{weight}kg情况下，在本次样本中，无人机在 {faster_bins} 这些距离区间内平均比货车更快；"
                f"其中最后一个仍保持优势的区间，其平均任务距离约为 {farthest_adv_km} km。"
                f"优势最明显的区间是 {best_bin}，在该区间内无人机平均比货车快 {best_adv_h} 小时。"
            )

        rows.append({
            "货物重量(kg)": weight,
            "结论文本": conclusion_text,
            "无人机更快的距离区间": faster_bins,
            "货车更快的距离区间": slower_bins,
            "无人机仍占优的最远平均距离(km)": farthest_adv_km,
            "优势最明显区间": best_bin,
            "该区间平均快多少(h)": best_adv_h,
        })

    return pd.DataFrame(rows)


def print_weight_conclusion(conclusion_df):
    if conclusion_df.empty:
        print("\n没有可输出的距离比较结论。")
        return

    print("\n===== 距离比较结论 =====")
    for _, row in conclusion_df.iterrows():
        print(row["结论文本"])


def save_analysis_result(batch_df, summary_df, distance_summary_df, conclusion_df):
    out_file = BASE_DIR / "无人机_vs_货车_批量模拟_提速版_含距离比较.xlsx"
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        batch_df.to_excel(writer, sheet_name="模拟明细", index=False)
        summary_df.to_excel(writer, sheet_name="总体汇总", index=False)
        distance_summary_df.to_excel(writer, sheet_name="距离分段比较", index=False)
        conclusion_df.to_excel(writer, sheet_name="文字结论", index=False)
    print(f"\n结果已保存：{out_file}")
    return out_file


# =========================
# 8. 主程序
# =========================
if __name__ == "__main__":
    pd.set_option("display.max_colwidth", 300)

    print("===== 批量模拟提速版（含距离比较） =====")

    # Excel 只读一次
    sheet1_df = load_sheet1(EXCEL_PATH, SHEET_NAME)

    # 跑 50 单示例，可改成 100、200
    batch_df = simulate_compare_many_fast(sheet1_df, n=50)

    print("\n===== 总体汇总结果 =====")
    summary_df = summarize_batch_result(batch_df)
    print(summary_df)

    print("\n===== 距离分段比较 =====")
    distance_summary_df = summarize_distance_advantage(batch_df, step_km=DISTANCE_BIN_STEP_KM)
    print(distance_summary_df)

    conclusion_df = build_weight_conclusion(distance_summary_df)
    print_weight_conclusion(conclusion_df)

    save_analysis_result(batch_df, summary_df, distance_summary_df, conclusion_df)
