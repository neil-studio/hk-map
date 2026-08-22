#!/usr/bin/env python3
"""
同步脚本：将 ORS 真实步行路网数据（距离、时长、折线轨迹）全量注入到运营沙盒地图系统中
"""

import json, math, time, os, sys
import requests

SANDBOX_DIR = "/Users/nb/google/Antigravity/工作/运营/价单/sandbox"
COORDS_PATH = os.path.join(SANDBOX_DIR, "hk_project_coords.json")
MAP_HTML_PATH = "/Users/nb/.gemini/antigravity/scratch/hk-property-map/hk_property_mtr_map.html"

ORS_TOKEN = "eyJvcmciOiI1YjNjZTM1OTc4NTExMTAwMDFjZjYyNDgiLCJpZCI6ImIwNjljZjY4YTgzMDQ1MTY5ZTMwZTZjOGExNDc2ODE3IiwiaCI6Im11cm11cjY0In0="
ORS_BASE = "https://api.openrouteservice.org/v2/directions/foot-walking"

MTR_STATIONS = [
    {"name_en":"Kennedy Town","name_zh":"堅尼地城","lat":22.2814,"lng":114.1286,"lines":["港島綫"]},
    {"name_en":"HKU","name_zh":"香港大學","lat":22.2840,"lng":114.1347,"lines":["港島綫"]},
    {"name_en":"Sai Ying Pun","name_zh":"西營盤","lat":22.2856,"lng":114.1427,"lines":["港島綫"]},
    {"name_en":"Sheung Wan","name_zh":"上環","lat":22.2865,"lng":114.1520,"lines":["港島綫","荃灣綫"]},
    {"name_en":"Central","name_zh":"中環","lat":22.2820,"lng":114.1583,"lines":["港島綫","荃灣綫"]},
    {"name_en":"Admiralty","name_zh":"金鐘","lat":22.2793,"lng":114.1654,"lines":["港島綫","荃灣綫","南港島綫","東鐵綫"]},
    {"name_en":"Wan Chai","name_zh":"灣仔","lat":22.2776,"lng":114.1731,"lines":["港島綫"]},
    {"name_en":"Causeway Bay","name_zh":"銅鑼灣","lat":22.2802,"lng":114.1840,"lines":["港島綫"]},
    {"name_en":"Tin Hau","name_zh":"天后","lat":22.2825,"lng":114.1918,"lines":["港島綫"]},
    {"name_en":"Fortress Hill","name_zh":"炮台山","lat":22.2879,"lng":114.1937,"lines":["港島綫"]},
    {"name_en":"North Point","name_zh":"北角","lat":22.2912,"lng":114.2001,"lines":["港島綫","將軍澳綫"]},
    {"name_en":"Quarry Bay","name_zh":"鰂魚涌","lat":22.2878,"lng":114.2097,"lines":["港島綫","將軍澳綫"]},
    {"name_en":"Tai Koo","name_zh":"太古","lat":22.2845,"lng":114.2165,"lines":["港島綫"]},
    {"name_en":"Sai Wan Ho","name_zh":"西灣河","lat":22.2828,"lng":114.2222,"lines":["港島綫"]},
    {"name_en":"Shau Kei Wan","name_zh":"筲箕灣","lat":22.2790,"lng":114.2283,"lines":["港島綫"]},
    {"name_en":"Heng Fa Chuen","name_zh":"杏花邨","lat":22.2766,"lng":114.2396,"lines":["港島綫"]},
    {"name_en":"Chai Wan","name_zh":"柴灣","lat":22.2649,"lng":114.2371,"lines":["港島綫"]},
    {"name_en":"Tsim Sha Tsui","name_zh":"尖沙咀","lat":22.2973,"lng":114.1726,"lines":["荃灣綫"]},
    {"name_en":"Jordan","name_zh":"佐敦","lat":22.3050,"lng":114.1716,"lines":["荃灣綫"]},
    {"name_en":"Yau Ma Tei","name_zh":"油麻地","lat":22.3131,"lng":114.1708,"lines":["荃灣綫","觀塘綫"]},
    {"name_en":"Mong Kok","name_zh":"旺角","lat":22.3193,"lng":114.1694,"lines":["荃灣綫","觀塘綫"]},
    {"name_en":"Prince Edward","name_zh":"太子","lat":22.3245,"lng":114.1683,"lines":["荃灣綫","觀塘綫"]},
    {"name_en":"Sham Shui Po","name_zh":"深水埗","lat":22.3309,"lng":114.1622,"lines":["荃灣綫"]},
    {"name_en":"Cheung Sha Wan","name_zh":"長沙灣","lat":22.3358,"lng":114.1558,"lines":["荃灣綫"]},
    {"name_en":"Lai Chi Kok","name_zh":"荔枝角","lat":22.3375,"lng":114.1481,"lines":["荃灣綫"]},
    {"name_en":"Mei Foo","name_zh":"美孚","lat":22.3386,"lng":114.1377,"lines":["荃灣綫","屯馬綫"]},
    {"name_en":"Lai King","name_zh":"荔景","lat":22.3483,"lng":114.1263,"lines":["荃灣綫","東涌綫"]},
    {"name_en":"Kwai Fong","name_zh":"葵芳","lat":22.3570,"lng":114.1279,"lines":["荃灣綫"]},
    {"name_en":"Kwai Hing","name_zh":"葵興","lat":22.3632,"lng":114.1313,"lines":["荃灣綫"]},
    {"name_en":"Tai Wo Hau","name_zh":"大窩口","lat":22.3708,"lng":114.1253,"lines":["荃灣綫"]},
    {"name_en":"Tsuen Wan","name_zh":"荃灣","lat":22.3733,"lng":114.1177,"lines":["荃灣綫"]},
    {"name_en":"Whampoa","name_zh":"黃埔","lat":22.3048,"lng":114.1897,"lines":["觀塘綫"]},
    {"name_en":"Ho Man Tin","name_zh":"何文田","lat":22.3094,"lng":114.1830,"lines":["觀塘綫","屯馬綫"]},
    {"name_en":"Kowloon Tong","name_zh":"九龍塘","lat":22.3368,"lng":114.1760,"lines":["東鐵綫","觀塘綫"]},
    {"name_en":"Wong Tai Sin","name_zh":"黃大仙","lat":22.3424,"lng":114.1936,"lines":["觀塘綫"]},
    {"name_en":"Diamond Hill","name_zh":"鑽石山","lat":22.3401,"lng":114.2015,"lines":["觀塘綫","屯馬綫"]},
    {"name_en":"Choi Hung","name_zh":"彩虹","lat":22.3349,"lng":114.2088,"lines":["觀塘綫"]},
    {"name_en":"Kowloon Bay","name_zh":"九龍灣","lat":22.3235,"lng":114.2138,"lines":["觀塘綫"]},
    {"name_en":"Ngau Tau Kok","name_zh":"牛頭角","lat":22.3158,"lng":114.2193,"lines":["觀塘綫"]},
    {"name_en":"Kwun Tong","name_zh":"觀塘","lat":22.3123,"lng":114.2264,"lines":["觀塘綫"]},
    {"name_en":"Lam Tin","name_zh":"藍田","lat":22.3068,"lng":114.2327,"lines":["觀塘綫"]},
    {"name_en":"Yau Tong","name_zh":"油塘","lat":22.2977,"lng":114.2374,"lines":["觀塘綫","將軍澳綫"]},
    {"name_en":"Tiu Keng Leng","name_zh":"調景嶺","lat":22.3046,"lng":114.2527,"lines":["觀塘綫","將軍澳綫"]},
    {"name_en":"Tseung Kwan O","name_zh":"將軍澳","lat":22.3073,"lng":114.2600,"lines":["將軍澳綫"]},
    {"name_en":"Hang Hau","name_zh":"坑口","lat":22.3155,"lng":114.2646,"lines":["將軍澳綫"]},
    {"name_en":"Po Lam","name_zh":"寶琳","lat":22.3228,"lng":114.2581,"lines":["將軍澳綫"]},
    {"name_en":"LOHAS Park","name_zh":"康城","lat":22.2955,"lng":114.2695,"lines":["將軍澳綫"]},
    {"name_en":"Hong Kong","name_zh":"香港","lat":22.2849,"lng":114.1585,"lines":["東涌綫","機場快綫"]},
    {"name_en":"Kowloon","name_zh":"九龍","lat":22.3050,"lng":114.1615,"lines":["東涌綫","機場快綫"]},
    {"name_en":"Olympic","name_zh":"奧運","lat":22.3178,"lng":114.1601,"lines":["東涌綫"]},
    {"name_en":"Nam Cheong","name_zh":"南昌","lat":22.3265,"lng":114.1537,"lines":["東涌綫","屯馬綫"]},
    {"name_en":"Tsing Yi","name_zh":"青衣","lat":22.3584,"lng":114.1074,"lines":["東涌綫","機場快綫"]},
    {"name_en":"Sunny Bay","name_zh":"欣澳","lat":22.3317,"lng":114.0291,"lines":["東涌綫","迪士尼綫"]},
    {"name_en":"Tung Chung","name_zh":"東涌","lat":22.2892,"lng":114.0001,"lines":["東涌綫"]},
    {"name_en":"Ocean Park","name_zh":"海洋公園","lat":22.2489,"lng":114.1743,"lines":["南港島綫"]},
    {"name_en":"Wong Chuk Hang","name_zh":"黃竹坑","lat":22.2480,"lng":114.1680,"lines":["南港島綫"]},
    {"name_en":"Lei Tung","name_zh":"利東","lat":22.2422,"lng":114.1563,"lines":["南港島綫"]},
    {"name_en":"South Horizons","name_zh":"海怡半島","lat":22.2429,"lng":114.1490,"lines":["南港島綫"]},
    {"name_en":"Exhibition Centre","name_zh":"會展","lat":22.2816,"lng":114.1753,"lines":["東鐵綫"]},
    {"name_en":"Hung Hom","name_zh":"紅磡","lat":22.3029,"lng":114.1818,"lines":["東鐵綫","屯馬綫"]},
    {"name_en":"Mong Kok East","name_zh":"旺角東","lat":22.3221,"lng":114.1722,"lines":["東鐵綫"]},
    {"name_en":"Tai Wai","name_zh":"大圍","lat":22.3726,"lng":114.1785,"lines":["東鐵綫","屯馬綫"]},
    {"name_en":"Sha Tin","name_zh":"沙田","lat":22.3827,"lng":114.1876,"lines":["東鐵綫"]},
    {"name_en":"Fo Tan","name_zh":"火炭","lat":22.3955,"lng":114.1984,"lines":["東鐵綫"]},
    {"name_en":"Racecourse","name_zh":"馬場","lat":22.4012,"lng":114.1836,"lines":["東鐵綫"]},
    {"name_en":"University","name_zh":"大學","lat":22.4133,"lng":114.2100,"lines":["東鐵綫"]},
    {"name_en":"Tai Po Market","name_zh":"大埔墟","lat":22.4446,"lng":114.1706,"lines":["東鐵綫"]},
    {"name_en":"Tai Wo","name_zh":"太和","lat":22.4510,"lng":114.1611,"lines":["東鐵綫"]},
    {"name_en":"Fanling","name_zh":"粉嶺","lat":22.4922,"lng":114.1384,"lines":["東鐵綫"]},
    {"name_en":"Sheung Shui","name_zh":"上水","lat":22.5013,"lng":114.1276,"lines":["東鐵綫"]},
    {"name_en":"Lo Wu","name_zh":"羅湖","lat":22.5282,"lng":114.1133,"lines":["東鐵綫"]},
    {"name_en":"Lok Ma Chau","name_zh":"落馬洲","lat":22.5147,"lng":114.0653,"lines":["東鐵綫"]},
    {"name_en":"Wu Kai Sha","name_zh":"烏溪沙","lat":22.4293,"lng":114.2436,"lines":["屯馬綫"]},
    {"name_en":"Ma On Shan","name_zh":"馬鞍山","lat":22.4250,"lng":114.2316,"lines":["屯馬綫"]},
    {"name_en":"Heng On","name_zh":"恆安","lat":22.4178,"lng":114.2256,"lines":["屯馬綫"]},
    {"name_en":"Tai Shui Hang","name_zh":"大水坑","lat":22.4084,"lng":114.2225,"lines":["屯馬綫"]},
    {"name_en":"Shek Mun","name_zh":"石門","lat":22.3884,"lng":114.2086,"lines":["屯馬綫"]},
    {"name_en":"City One","name_zh":"第一城","lat":22.3832,"lng":114.2037,"lines":["屯馬綫"]},
    {"name_en":"Sha Tin Wai","name_zh":"沙田圍","lat":22.3764,"lng":114.1952,"lines":["屯馬綫"]},
    {"name_en":"Che Kung Temple","name_zh":"車公廟","lat":22.3747,"lng":114.1861,"lines":["屯馬綫"]},
    {"name_en":"Hin Keng","name_zh":"顯徑","lat":22.3654,"lng":114.1708,"lines":["屯馬綫"]},
    {"name_en":"Kai Tak","name_zh":"啟德","lat":22.3307,"lng":114.1994,"lines":["屯馬綫"]},
    {"name_en":"Sung Wong Toi","name_zh":"宋皇臺","lat":22.3261,"lng":114.1894,"lines":["屯馬綫"]},
    {"name_en":"To Kwa Wan","name_zh":"土瓜灣","lat":22.3170,"lng":114.1870,"lines":["屯馬綫"]},
    {"name_en":"Austin","name_zh":"柯士甸","lat":22.3040,"lng":114.1667,"lines":["屯馬綫"]},
    {"name_en":"East Tsim Sha Tsui","name_zh":"尖東","lat":22.2949,"lng":114.1751,"lines":["屯馬綫"]},
    {"name_en":"Tsuen Wan West","name_zh":"荃灣西","lat":22.3684,"lng":114.1097,"lines":["屯馬綫"]},
    {"name_en":"Kam Sheung Road","name_zh":"錦上路","lat":22.4348,"lng":114.0684,"lines":["屯馬綫"]},
    {"name_en":"Yuen Long","name_zh":"元朗","lat":22.4461,"lng":114.0352,"lines":["屯馬綫"]},
    {"name_en":"Long Ping","name_zh":"朗屏","lat":22.4477,"lng":114.0255,"lines":["屯馬綫"]},
    {"name_en":"Tin Shui Wai","name_zh":"天水圍","lat":22.4480,"lng":114.0045,"lines":["屯馬綫"]},
    {"name_en":"Siu Hong","name_zh":"兆康","lat":22.4115,"lng":113.9782,"lines":["屯馬綫"]},
    {"name_en":"Tuen Mun","name_zh":"屯門","lat":22.3952,"lng":113.9733,"lines":["屯馬綫"]},
    {"name_en":"Airport","name_zh":"機場","lat":22.3160,"lng":113.9365,"lines":["機場快綫"]},
    {"name_en":"AsiaWorld-Expo","name_zh":"博覽館","lat":22.3207,"lng":113.9418,"lines":["機場快綫"]},
    {"name_en":"Disneyland Resort","name_zh":"迪士尼","lat":22.3153,"lng":114.0451,"lines":["迪士尼綫"]},
]

def dist_m(lat1, lon1, lat2, lon2):
    R = 6371000
    dLat = math.radians(lat2 - lat1)
    dLon = math.radians(lon2 - lon1)
    a = math.sin(dLat/2)**2 + math.cos(math.radians(lat1))*math.cos(math.radians(lat2))*math.sin(dLon/2)**2
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))

def ors_walk(from_lat, from_lng, to_lat, to_lng):
    params = {"start": f"{from_lng},{from_lat}", "end": f"{to_lng},{to_lat}"}
    headers = {"Authorization": ORS_TOKEN, "Accept": "application/json, application/geo+json"}
    try:
        resp = requests.get(ORS_BASE, params=params, headers=headers, timeout=12)
        if resp.status_code == 200:
            data = resp.json()
            feat = data["features"][0]
            seg = feat["properties"]["segments"][0]
            return {
                "distance": round(seg["distance"]),
                "duration": round(seg["duration"]),
                "geometry": feat["geometry"]
            }
    except Exception as e:
        pass
    return None

def main():
    print("1️⃣ 加载沙盒项目数据...")
    with open(COORDS_PATH, "r", encoding="utf-8") as f:
        sandbox_projects = json.load(f)
    print(f"   沙盒项目总数: {len(sandbox_projects)}")

    print("2️⃣ 加载 302 个已计算好的 ORS 真实路网数据...")
    with open(MAP_HTML_PATH, "r", encoding="utf-8") as f:
        html_c = f.read()
    idx1 = html_c.find("const DATA = ") + len("const DATA = ")
    idx2 = html_c.find(";\n\nlet map,", idx1)
    if idx2 == -1: idx2 = html_c.find(";\nlet map,", idx1)
    ors_data = json.loads(html_c[idx1:idx2])
    print(f"   已计算库总数: {len(ors_data)}")

    updated = 0
    calculated_on_fly = 0

    for i, p in enumerate(sandbox_projects):
        p_name = p.get("name", "").strip()
        p_lat = p.get("lat")
        p_lng = p.get("lng")
        
        # 1. 尝试匹配已有 ORS 结果
        match = next((d for d in ors_data if d["name"].strip() == p_name), None)
        if not match and p_lat and p_lng:
            closest = min(ors_data, key=lambda d: dist_m(p_lat, p_lng, d["lat"], d["lng"]))
            if dist_m(p_lat, p_lng, closest["lat"], closest["lng"]) < 80:
                match = closest
        
        if match and match.get("walk_distance"):
            p["nearest_mtr"] = {
                "address": match.get("address", ""),
                "nearest_mtr_cn": match.get("station_zh", ""),
                "nearest_mtr_en": match.get("station_en", ""),
                "nearest_mtr_lines": "/".join(match.get("lines", [])),
                "nearest_mtr_lat": match.get("station_lat"),
                "nearest_mtr_lng": match.get("station_lng"),
                "dist_straight_m": match.get("straight_distance"),
                "dist_walk_m": match.get("walk_distance"),
                "walk_time_min": match.get("walk_duration"),
                "route_geometry": match.get("route_geometry")
            }
            updated += 1
        elif p_lat and p_lng:
            # 2. 对少量未入库的沙盒楼盘，在线计算 ORS 真实人行道路线
            closest_st = min(MTR_STATIONS, key=lambda s: dist_m(p_lat, p_lng, s["lat"], s["lng"]))
            straight_m = round(dist_m(p_lat, p_lng, closest_st["lat"], closest_st["lng"]))
            route = ors_walk(p_lat, p_lng, closest_st["lat"], closest_st["lng"])
            time.sleep(1.6)
            
            if route:
                walk_m = route["distance"]
                walk_min = round(route["duration"] / 60, 1)
                route_geo = route["geometry"]
            else:
                walk_m = round(straight_m * 1.3)
                walk_min = round(walk_m / 80.0, 1)
                route_geo = None
            
            p["nearest_mtr"] = {
                "address": "",
                "nearest_mtr_cn": closest_st["name_zh"],
                "nearest_mtr_en": closest_st["name_en"],
                "nearest_mtr_lines": "/".join(closest_st["lines"]),
                "nearest_mtr_lat": closest_st["lat"],
                "nearest_mtr_lng": closest_st["lng"],
                "dist_straight_m": straight_m,
                "dist_walk_m": walk_m,
                "walk_time_min": walk_min,
                "route_geometry": route_geo
            }
            calculated_on_fly += 1
            updated += 1
            print(f"   ⚡ 在线测算补充: {p_name} → {closest_st['name_zh']} ({walk_m}m, {walk_min}min)")

    print(f"\n3️⃣ 保存更新后的 {COORDS_PATH}...")
    with open(COORDS_PATH, "w", encoding="utf-8") as f:
        json.dump(sandbox_projects, f, ensure_ascii=False, indent=2)
    print(f"   ✅ 完成！共 {len(sandbox_projects)} 个楼盘，成功挂载真实路线数据 {updated} 个 (其中在线补充 {calculated_on_fly} 个)")

if __name__ == "__main__":
    main()
