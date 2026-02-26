import pandas as pd
import requests
import urllib.parse

# 1. 你的 API 密钥 (建议后续改为从 os.environ 获取)
API_KEY = "AIzaSyDzAxmeKWeWQGK3G5VVnEDLM0IY-RTjzrw"

# 2. 提取起点/终点坐标：乌普萨拉仓库
# 根据你的坐标文件: "Lager – Björkgatan 65, Uppsala",59.8542194,17.6650221
WAREHOUSE_LAT = "59.8542194"
WAREHOUSE_LNG = "17.6650221"
WAREHOUSE_COORD = f"{WAREHOUSE_LAT},{WAREHOUSE_LNG}"

# 3. 模拟匹配好的单条路线数据 (这里以 Abbe 的前几个店为例)
# 在实际运行中，你需要用 pandas 读取 nya rutter.xlsx 和 坐标.xlsx 进行 merge
# 例如: df_coords = pd.read_csv('坐标.csv')
driver_stores = [
    {"name": "Ica Supermarket Hagsätra", "lat": "59.2629", "lng": "18.0135"}, # 替换为真实匹配的坐标
    {"name": "Ica Nära Stuvsta", "lat": "59.2568", "lng": "17.9859"},
    {"name": "Coop Stuvsta", "lat": "59.2565", "lng": "17.9862"}
    # ... 添加 Abbe 路线上的所有店 (最多23个)
]

def optimize_route(stores):
    """
    调用 Google Maps Directions API 优化途经点顺序
    """
    # 将门店坐标拼接成 API 需要的 waypoints 格式
    # 关键参数: optimize:true，这会让 Google 自动为你求解最短路径(TSP)
    waypoints_list = [f"{store['lat']},{store['lng']}" for store in stores]
    waypoints_str = "optimize:true|" + "|".join(waypoints_list)

    url = "https://maps.googleapis.com/maps/api/directions/json"
    params = {
        "origin": WAREHOUSE_COORD,
        "destination": WAREHOUSE_COORD, # 假设司机最后需要返回仓库
        "waypoints": waypoints_str,
        "key": API_KEY
    }

    print("正在请求 Google Maps API 进行路径优化...")
    response = requests.get(url, params=params)
    data = response.json()

    if data['status'] == 'OK':
        # 提取优化后的索引顺序
        optimized_order = data['routes'][0]['waypoint_order']
        print(f"API 返回的优化顺序索引: {optimized_order}")
        
        # 根据优化后的顺序重排门店
        optimized_stores = [stores[i] for i in optimized_order]
        return optimized_stores
    else:
        print(f"API 错误: {data['status']}")
        if 'error_message' in data:
            print(data['error_message'])
        return None

def generate_google_maps_url(optimized_stores):
    """
    生成供司机点击的 Google Maps 导航链接
    """
    base_url = "https://www.google.com/maps/dir/?api=1"
    origin_param = f"&origin={WAREHOUSE_COORD}"
    destination_param = f"&destination={WAREHOUSE_COORD}"
    
    # 将优化后的坐标拼接为 URL 参数
    waypoints = "|".join([f"{store['lat']},{store['lng']}" for store in optimized_stores])
    # 注意：URL 中的 | 需要被转义为 %7C
    waypoints_param = "&waypoints=" + urllib.parse.quote(waypoints)
    
    # 注意：Google Maps Web URL 限制了 waypoints 的数量（通常最多 9-10 个），
    # 但通过程序化分段或特定的客户端深度链接可以绕过。
    # 这里先生成标准的 Web 端/App 唤醒链接
    final_url = base_url + origin_param + destination_param + waypoints_param
    return final_url

# === 运行测试 ===
optimized = optimize_route(driver_stores)
if optimized:
    print("\n优化后的拜访顺序:")
    for idx, store in enumerate(optimized, 1):
        print(f"{idx}. {store['name']}")
    
    nav_link = generate_google_maps_url(optimized)
    print(f"\n✅ 为司机生成的导航链接:\n{nav_link}")