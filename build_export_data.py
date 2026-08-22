#!/usr/bin/env python3
"""
销控数据与成交数据增强处理脚本
自动遍历全部 Excel 销控表，提取：
1. 项目在售户型、面积段、价格段、招标标识 (未售房源)
2. 项目所有户型最近成交的 2 套真实成交记录 (楼栋单位号、面积、总价、单价、成交日期)
3. 发展商与落成日期
更新至 sandbox/hk_project_coords.json
"""

import os, json, re, glob
import openpyxl

SANDBOX_DIR = "/Users/nb/google/Antigravity/工作/运营/价单/sandbox"
PROD_DIR = "/Users/nb/google/Antigravity/工作/运营/价单_正式版"
FILES_DIR = os.path.join(PROD_DIR, "files")
COORDS_FILE = os.path.join(SANDBOX_DIR, "hk_project_coords.json")
DATA_FILE = os.path.join(PROD_DIR, "data.json")

KNOWN_DEVELOPERS = [
    ("新鸿基地产", ["新鸿基地产", "新鸿基", "SHKP", "Sun Hung Kai"]),
    ("恒基兆业地产", ["恒基兆业", "恒基地产", "恒基", "Henderson"]),
    ("长实集团", ["长实集团", "长江实业", "长实", "CK Asset"]),
    ("会德丰地产", ["会德丰地产", "会德丰", "Wheelock"]),
    ("新世界发展", ["新世界发展", "新世界", "New World"]),
    ("信和置业", ["信和置业", "信和", "Sino"]),
    ("嘉里建设", ["嘉里建设", "嘉里", "Kerry"]),
    ("华懋集团", ["华懋集团", "华懋", "Chinachem"]),
    ("太古地产", ["太古地产", "太古", "Swire"]),
    ("嘉华国际", ["嘉华国际", "嘉华", "K. Wah"]),
    ("九龙建业", ["九龙建业", "九建", "Kowloon Development"]),
    ("兴胜创建", ["兴胜创建", "Hanison"]),
    ("路劲地产", ["路劲地产", "路劲", "Road King"]),
    ("中国海外发展", ["中国海外", "中海", "China Overseas"]),
    ("保利置业", ["保利置业", "保利", "Poly Property"]),
    ("万科香港", ["万科香港", "万科", "Vanke"]),
    ("香港兴业", ["香港兴业", "HKR International"]),
    ("英皇国际", ["英皇国际", "英皇", "Emperor"]),
    ("远东发展", ["远东发展", "远展", "Far East"]),
    ("爪哇控股", ["爪哇控股", "爪哇"]),
    ("南丰集团", ["南丰集团", "南丰", "Nan Fung"]),
    ("资本策略", ["资本策略", "CSI Properties"]),
    ("宏安地产", ["宏安地产", "宏安", "Wang On"]),
]

def clean_price(val):
    if val is None or val == '-' or val == '暂无' or str(val).strip() == '':
        return None
    s = str(val).replace(',', '').replace('$', '').replace('HKD', '').replace('港币', '').strip()
    try:
        f = float(s)
        if f > 2000000000 and len(s) >= 14:
            f = float(s[:7])
        return f if f > 0 else None
    except:
        return None

def fmt_currency(p):
    if not p:
        return "暂无"
    p_wan = p / 10000
    if p_wan >= 10000:
        return f"{p_wan/10000:.2f} 亿".replace(".00", "")
    return f"{round(p_wan, 1)} 万".replace(".0 万", " 万")

def fmt_uprice(u):
    if not u or u == '-' or u == '暂无':
        return "-"
    try:
        f = float(str(u).replace(',', '').replace('$', '').strip())
        return f"${round(f):,}/呎"
    except:
        return str(u)

def extract_developer(text):
    if not text:
        return "香港品牌发展商"
    for canonical, aliases in KNOWN_DEVELOPERS:
        for alias in aliases:
            if alias in text:
                return canonical
    m = re.search(r"由([^\s，。、（]+?)(?:倾力|重磅|精心)?(?:打造|发展|兴建|营建|推出)", text)
    if m:
        dev = m.group(1).replace("**", "").strip()
        if len(dev) <= 12:
            return dev
    return "香港品牌发展商"

def extract_completion_date(text):
    if not text:
        return "现楼 / 详见项目资料"
    if "现楼" in text:
        return "现楼发售 (即买即住)"
    m = re.search(r"(202[4-9]年[^\s，。、（]*)", text)
    if m:
        return m.group(1).replace("**", "")
    return "现楼 / 详见项目资料"

def parse_excel_sales_and_tx():
    files = glob.glob(os.path.join(FILES_DIR, "*.xlsx"))
    project_sales = {}
    project_tx = {}
    
    for fpath in files:
        fname = os.path.basename(fpath).replace(".xlsx", "")
        parts = fname.split("-")
        pname = parts[-1].strip()
        
        try:
            wb = openpyxl.load_workbook(fpath, data_only=True)
            sheet = wb.active
            rows = list(sheet.iter_rows(values_only=True))
            if len(rows) < 3:
                continue
            
            header = rows[1]
            header_strs = [str(h).strip() if h is not None else "" for h in header]
            
            def find_col(possible_names):
                for name in possible_names:
                    for idx, h in enumerate(header_strs):
                        if name in h:
                            return idx
                return -1

            col_block = find_col(["楼栋", "座数", "座向", "大厦"])
            col_floor = find_col(["楼层", "层数"])
            col_unit = find_col(["房号", "单位", "室"])
            col_layout = find_col(["户型", "间隔", "单位类型", "房间"])
            col_sqft = find_col(["实用面积", "面积", "呎数"])
            col_status = find_col(["销控状态", "状态", "销售状态"])
            col_price = find_col(["折实总价", "总价", "售价", "成交价"])
            col_uprice = find_col(["折实呎价", "实用呎价", "呎价"])
            col_date = find_col(["成交日期", "日期"])
            col_tender = find_col(["是否招标", "招标"])

            if col_layout == -1 or col_sqft == -1 or col_status == -1:
                continue

            unsold_layout_data = {}
            sold_layout_data = {}

            for r in rows[2:]:
                if len(r) <= max(col_layout, col_sqft, col_status):
                    continue
                
                status = str(r[col_status]).strip() if r[col_status] is not None else ""
                raw_layout = str(r[col_layout]).strip() if r[col_layout] is not None else "特色单位"
                if not raw_layout or raw_layout == "None" or raw_layout == "-":
                    raw_layout = "特色单位"
                
                layout = raw_layout
                if "开放" in layout or "0房" in layout: layout = "开放式"
                elif "1" in layout or "一房" in layout or "1房" in layout: layout = "1 房"
                elif "2" in layout or "两房" in layout or "2房" in layout or "二房" in layout: layout = "2 房"
                elif "3" in layout or "三房" in layout or "3房" in layout: layout = "3 房"
                elif "4" in layout or "四房" in layout or "4房" in layout: layout = "4 房"
                elif "5" in layout or "五房" in layout or "5房" in layout: layout = "5 房 / 大宅"
                elif "洋房" in layout or "独立屋" in layout or "别墅" in layout: layout = "独立洋房"
                elif "复式" in layout: layout = "复式大宅"
                elif "特色" in layout or "天台" in layout or "花园" in layout: layout = "特色单位"
                
                sqft = r[col_sqft]
                price = clean_price(r[col_price]) if col_price != -1 and col_price < len(r) else None
                uprice = r[col_uprice] if col_uprice != -1 and col_uprice < len(r) else None
                date = str(r[col_date]).strip() if col_date != -1 and col_date < len(r) and r[col_date] is not None else ""
                if date in ["-", "None", "暂无"]: date = ""

                # 1. 已售单位 -> 存入 sold_layout_data 用于最近成交
                if status in ["已售", "已售出", "已签约", "成交"]:
                    block = str(r[col_block]).strip() if col_block != -1 and r[col_block] is not None and str(r[col_block]) != 'None' else ''
                    floor = str(r[col_floor]).strip() if col_floor != -1 and r[col_floor] is not None and str(r[col_floor]) != 'None' else ''
                    unit = str(r[col_unit]).strip() if col_unit != -1 and r[col_unit] is not None and str(r[col_unit]) != 'None' else ''
                    
                    unit_parts = []
                    if block: unit_parts.append(block if ('座' in block or '栋' in block) else f'{block}座')
                    if floor: unit_parts.append(f'{floor}楼')
                    if unit: unit_parts.append(f'{unit}室')
                    unit_name = ' '.join(unit_parts) if unit_parts else '精选单位'

                    if layout not in sold_layout_data:
                        sold_layout_data[layout] = []
                    
                    sold_layout_data[layout].append({
                        'unit_name': unit_name,
                        'layout': layout,
                        'sqft': f'{int(sqft)} 呎' if isinstance(sqft, (int, float)) and sqft > 0 else '详见销控',
                        'price': fmt_currency(price),
                        'unit_price': fmt_uprice(uprice),
                        'date': date or '-'
                    })
                    continue

                # 2. 未售房源 -> 统计在售销控汇总
                is_tender = False
                if col_tender != -1 and col_tender < len(r) and r[col_tender] is not None:
                    if str(r[col_tender]).strip() in ["是", "招标", "True", "true"]:
                        is_tender = True
                if col_price != -1 and col_price < len(r) and str(r[col_price]).strip() in ["招标", "招标单位"]:
                    is_tender = True
                
                if layout not in unsold_layout_data:
                    unsold_layout_data[layout] = {"sqfts": [], "prices": [], "tender_count": 0, "unsold": 0}
                
                unsold_layout_data[layout]["unsold"] += 1
                if isinstance(sqft, (int, float)) and sqft > 0:
                    unsold_layout_data[layout]["sqfts"].append(sqft)
                
                if is_tender:
                    unsold_layout_data[layout]["tender_count"] += 1
                elif price:
                    unsold_layout_data[layout]["prices"].append(price)

            # 整理在售房源列表
            summary_list = []
            layout_order = ["开放式", "1 房", "2 房", "3 房", "4 房", "5 房 / 大宅", "复式大宅", "特色单位", "独立洋房"]
            sorted_keys = sorted(unsold_layout_data.keys(), key=lambda k: layout_order.index(k) if k in layout_order else 99)
            
            for l_name in sorted_keys:
                d = unsold_layout_data[l_name]
                if d["unsold"] == 0:
                    continue
                
                if d["sqfts"]:
                    min_s, max_s = int(min(d["sqfts"])), int(max(d["sqfts"]))
                    sqft_str = f"{min_s} 呎" if min_s == max_s else f"{min_s} - {max_s} 呎"
                else:
                    sqft_str = "详见销控"
                
                if d["prices"]:
                    min_p = min(d["prices"])
                    max_p = max(d["prices"])
                    p_str1 = fmt_currency(min_p)
                    p_str2 = fmt_currency(max_p)
                    price_str = p_str1 if p_str1 == p_str2 else f"{p_str1} - {p_str2}"
                    if d["tender_count"] > 0:
                        price_str += " (含招标)"
                elif d["tender_count"] > 0:
                    price_str = "招标"
                else:
                    price_str = "待定/详见价单"
                
                summary_list.append({
                    "layout": l_name,
                    "unsold_count": d["unsold"],
                    "sqft_range": sqft_str,
                    "price_range": price_str
                })
            
            project_sales[pname] = summary_list

            # 整理各户型最近2套成交列表
            recent_tx_list = []
            sorted_sold_layouts = sorted(sold_layout_data.keys(), key=lambda k: layout_order.index(k) if k in layout_order else 99)
            for l_name in sorted_sold_layouts:
                tx_items = sold_layout_data[l_name]
                # 按日期倒序排列，取最近 2 套
                tx_items_sorted = sorted(tx_items, key=lambda x: str(x['date']) if x['date'] and x['date'] != '-' else '0000-00-00', reverse=True)
                for tx in tx_items_sorted[:2]:
                    recent_tx_list.append(tx)

            project_tx[pname] = recent_tx_list

        except Exception as e:
            print(f"⚠️ 解析 {fname} 失败: {e}")
            
    print(f"✅ 成功解析在售项目: {len(project_sales)} 个，成交记录项目: {len(project_tx)} 个")
    return project_sales, project_tx

def main():
    print("1️⃣ 开始解析销控表与成交记录...")
    sales_data, tx_data = parse_excel_sales_and_tx()

    print("2️⃣ 读取沙盒数据 hk_project_coords.json...")
    with open(COORDS_FILE, "r", encoding="utf-8") as f:
        projects = json.load(f)

    data_json_map = {}
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                dj = json.load(f)
                for p in dj.get("projects", []):
                    data_json_map[p["name"].strip()] = p
        except Exception as e:
            print("读取 data.json 异常:", e)

    print("3️⃣ 整合发展商、落成日期、在售明细、最近成交记录...")
    tx_attached = 0
    for p in projects:
        pname = p["name"].strip()
        dj_p = data_json_map.get(pname, {})
        
        full_text = " ".join([
            p.get("mainland_selling_points", ""),
            p.get("reason", ""),
            dj_p.get("mainland_selling_points", ""),
            dj_p.get("reason", "")
        ])

        p["developer"] = extract_developer(full_text)
        p["completion_date"] = extract_completion_date(full_text)

        # 匹配在售数据
        sc = sales_data.get(pname)
        if not sc:
            for k, val in sales_data.items():
                if k in pname or pname in k:
                    sc = val
                    break
        p["sales_control_summary"] = sc if sc else []

        # 匹配最近成交数据
        tx = tx_data.get(pname)
        if not tx:
            for k, val in tx_data.items():
                if k in pname or pname in k:
                    tx = val
                    break
        p["recent_transactions"] = tx if tx else []
        if p["recent_transactions"]:
            tx_attached += 1

    print(f"4️⃣ 保存更新至 {COORDS_FILE} (共 {len(projects)} 个楼盘，挂载最近成交数据 {tx_attached} 个)...")
    with open(COORDS_FILE, "w", encoding="utf-8") as f:
        json.dump(projects, f, ensure_ascii=False, indent=2)
    print("🎉 数据升级完成！")

if __name__ == "__main__":
    main()
