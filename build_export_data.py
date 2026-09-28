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

SANDBOX_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.abspath(os.path.join(SANDBOX_DIR, ".."))
PROD_DIR = "/Users/nb/google/Antigravity/工作/运营/价单_正式版"
FILES_DIR = os.path.join(PROD_DIR, "files")
COORDS_FILE = os.path.join(SANDBOX_DIR, "hk_project_coords.json")

DATA_FILE = os.path.join(PROD_DIR, "data.json")
if not os.path.exists(DATA_FILE) or not os.access(DATA_FILE, os.R_OK):
    DATA_FILE = os.path.join(BASE_DIR, "data.json")

BASE_INFO_FILE = os.path.join(PROD_DIR, "楼盘基础信息库.xlsx")
if not os.path.exists(BASE_INFO_FILE) or not os.access(BASE_INFO_FILE, os.R_OK):
    alt_base = "/Users/nb/google/Antigravity/工作/运营/聚焦盘精选盘/楼盘资料.xlsx"
    if os.path.exists(alt_base) and os.access(alt_base, os.R_OK):
        BASE_INFO_FILE = alt_base
    else:
        BASE_INFO_FILE = os.path.join(BASE_DIR, "楼盘基础信息库.xlsx")

SELLING_POINTS_FILE = os.path.join(SANDBOX_DIR, "楼盘卖点.xlsx")

def load_selling_points():
    sp_map = {}
    if not os.path.exists(SELLING_POINTS_FILE):
        return sp_map
    try:
        wb = openpyxl.load_workbook(SELLING_POINTS_FILE, data_only=True)
        ws = wb.active
        t_chars = '號峯滙敍灣瓏鑽鑄賢譽啟璽維奧滶蘊匯悅寶應羅壽淺徑賓園華遠藍語區築緹瀧澐雙鑽閣揚晉凱島輝臺連東紅臺豐環學區萬濱嶺軒廣場頭門閣薈'
        s_chars = '号峰汇叙湾珑钻铸贤誉启玺维奥滶蕴汇悦宝应罗寿浅径宾园华远蓝语区筑缇泷沄双钻阁扬晋凯岛辉台连东红台丰环学区万滨岭轩广场头门阁荟'
        trans = str.maketrans(t_chars, s_chars)
        for r in list(ws.iter_rows(values_only=True))[1:]:
            if r and r[0] and r[1]:
                pn = str(r[0]).strip()
                sp = str(r[1]).strip()
                norm_pn = pn.translate(trans).replace(' ', '').replace('　', '').replace('．', '.').lower()
                sp_map[pn] = sp
                sp_map[norm_pn] = sp
    except Exception as e:
        print("读取 楼盘卖点.xlsx 异常:", e)
    return sp_map

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

def extract_developer(pname="", text="", dev_map=None):
    if dev_map:
        if pname and pname in dev_map:
            return dev_map[pname]
        t_chars = '號峯滙敍灣瓏鑽鑄賢譽啟璽維奧滶蘊匯悅寶應羅壽淺徑賓園華遠藍語區築緹瀧澐雙鑽閣揚晉凱島輝臺連東紅臺豐環學區萬濱嶺軒廣場頭門閣薈'
        s_chars = '号峰汇叙湾珑钻铸贤誉启玺维奥滶蕴汇悦宝应罗寿浅径宾园华远蓝语区筑缇泷沄双钻阁扬晋凯岛辉台连东红台丰环学区万滨岭轩广场头门阁荟'
        trans = str.maketrans(t_chars, s_chars)
        norm_pn = pname.translate(trans).replace(' ', '').replace('　', '').replace('．', '.').replace('·', '.').lower()
        if norm_pn in dev_map:
            return dev_map[norm_pn]
        for k, v in dev_map.items():
            if k and (k == norm_pn or k in norm_pn or norm_pn in k):
                return v
    if not text:
        return "香港品牌发展商"
    for canonical, aliases in KNOWN_DEVELOPERS:
        for alias in aliases:
            if alias in text:
                return canonical
    m = re.search(r"由([^\s，。、（]+?)(?:倾力|重磅|精心)?(?:打造|发展|兴建|营建|推出)", text)
    if m:
        dev = m.group(1).replace("**", "").strip()
        if len(dev) <= 20:
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
    if not files:
        files = glob.glob(os.path.join(SANDBOX_DIR, "../files", "*.xlsx"))
    if not files:
        files = glob.glob(os.path.join(SANDBOX_DIR, "../*/*_销控明细表.xlsx"))
    project_sales = {}
    project_tx = {}
    project_bounds = {}
    project_unit_stats = {}
    
    for fpath in files:
        fname = os.path.basename(fpath).replace(".xlsx", "").replace("_销控明细表", "")
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
            
            # 分离折实价与基础总价/呎价列
            col_disc_price = find_col(["折实总价"])
            col_tot_price = find_col(["总价", "售价", "成交价"])
            col_disc_uprice = find_col(["折实呎价"])
            col_tot_uprice = find_col(["实用呎价", "呎价"])
            
            col_date = find_col(["成交日期", "日期"])
            col_tender = find_col(["是否招标", "招标"])

            if col_layout == -1 or col_sqft == -1 or col_status == -1:
                continue

            unsold_layout_data = {}
            sold_layout_data = {}
            excel_total_units = 0
            excel_sold_units = 0
            excel_on_sale_units = 0

            for r in rows[2:]:
                if len(r) <= max(col_layout, col_sqft, col_status):
                    continue
                
                excel_total_units += 1
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
                
                # 双重回退读取总价与呎价
                p_disc = clean_price(r[col_disc_price]) if col_disc_price != -1 and col_disc_price < len(r) else None
                p_tot = clean_price(r[col_tot_price]) if col_tot_price != -1 and col_tot_price < len(r) else None
                price = p_disc if p_disc is not None else p_tot

                u_disc = r[col_disc_uprice] if col_disc_uprice != -1 and col_disc_uprice < len(r) else None
                u_tot = r[col_tot_uprice] if col_tot_uprice != -1 and col_tot_uprice < len(r) else None
                uprice = u_disc if (u_disc not in [None, '-', '暂无', '']) else u_tot

                date = str(r[col_date]).strip() if col_date != -1 and col_date < len(r) and r[col_date] is not None else ""
                if date in ["-", "None", "暂无"]: date = ""

                # 1. 已售单位 -> 存入 sold_layout_data 用于最近成交
                if status in ["已售", "已售出", "已签约", "成交"]:
                    excel_sold_units += 1
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
                        'price_num': price,
                        'unit_price': fmt_uprice(uprice),
                        'date': date or '-'
                    })
                    continue

                # 2. 未售房源 -> 统计在售销控汇总
                is_tender = False
                if col_tender != -1 and col_tender < len(r) and r[col_tender] is not None:
                    if str(r[col_tender]).strip() in ["是", "招标", "True", "true"]:
                        is_tender = True
                if col_disc_price != -1 and col_disc_price < len(r) and str(r[col_disc_price]).strip() in ["招标", "招标单位"]:
                    is_tender = True
                if col_tot_price != -1 and col_tot_price < len(r) and str(r[col_tot_price]).strip() in ["招标", "招标单位"]:
                    is_tender = True
                
                # 判断是否属于狭义在售（已正式发售安排推售）
                is_on_sale = (status in ["在售", "出售中", "sale", "Sale"])
                
                if layout not in unsold_layout_data:
                    unsold_layout_data[layout] = {"sqfts": [], "prices": [], "tender_count": 0, "unsold": 0, "on_sale": 0}
                
                unsold_layout_data[layout]["unsold"] += 1
                if is_on_sale:
                    excel_on_sale_units += 1
                    unsold_layout_data[layout]["on_sale"] += 1

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
            
            all_unsold_prices = []
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
                    all_unsold_prices.extend(d["prices"])
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
                    "on_sale_count": d["on_sale"],
                    "sqft_range": sqft_str,
                    "price_range": price_str
                })
            
            towers = [s for s in wb.sheetnames if s != '销控汇总明细']
            tower_count = len(towers) if len(towers) > 0 else 1
            project_sales[pname] = summary_list
            project_unit_stats[pname] = {
                "total_units": excel_total_units,
                "sold_units": excel_sold_units,
                "sale_units": excel_on_sale_units,
                "tower_count": tower_count,
                "tower_names": towers
            }

            # 整理各户型最近2套成交列表
            recent_tx_list = []
            all_sold_prices = []
            sorted_sold_layouts = sorted(sold_layout_data.keys(), key=lambda k: layout_order.index(k) if k in layout_order else 99)
            for l_name in sorted_sold_layouts:
                tx_items = sold_layout_data[l_name]
                for item in tx_items:
                    if item.get('price_num'):
                        all_sold_prices.append(item['price_num'])
                # 按日期倒序排列，取最近 2 套
                tx_items_sorted = sorted(tx_items, key=lambda x: str(x['date']) if x['date'] and x['date'] != '-' else '0000-00-00', reverse=True)
                for tx in tx_items_sorted[:2]:
                    # 清理 price_num 临时字段
                    clean_tx = {k: v for k, v in tx.items() if k != 'price_num'}
                    recent_tx_list.append(clean_tx)

            project_tx[pname] = recent_tx_list

            # 计算项目的精准最低与最高价格 (万元)
            if all_unsold_prices:
                min_wan = round(min(all_unsold_prices) / 10000.0, 2)
                max_wan = round(max(all_unsold_prices) / 10000.0, 2)
                project_bounds[pname] = {
                    'min_price_wan': min_wan,
                    'max_price_wan': max_wan,
                    'has_price_data': True,
                    'desc': f"${min_wan}万起" if min_wan == max_wan else f"${min_wan}万 - ${max_wan}万"
                }
            elif all_sold_prices:
                min_wan = round(min(all_sold_prices) / 10000.0, 2)
                max_wan = round(max(all_sold_prices) / 10000.0, 2)
                project_bounds[pname] = {
                    'min_price_wan': min_wan,
                    'max_price_wan': max_wan,
                    'has_price_data': True,
                    'desc': f"参考成交: ${min_wan}万起" if min_wan == max_wan else f"参考成交: ${min_wan}万 - ${max_wan}万"
                }
            else:
                project_bounds[pname] = {
                    'min_price_wan': None,
                    'max_price_wan': None,
                    'has_price_data': False,
                    'desc': '招标发售 / 详见价单'
                }

        except Exception as e:
            print(f"⚠️ 解析 {fname} 失败: {e}")
            
    print(f"✅ 成功解析在售项目: {len(project_sales)} 个，成交记录项目: {len(project_tx)} 个")
    return project_sales, project_tx, project_bounds, project_unit_stats

def main():
    print("1️⃣ 开始解析销控表与成交记录...")
    sales_data, tx_data, bounds_data, unit_stats_data = parse_excel_sales_and_tx()

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

    school_net_map = {}
    if os.path.exists(BASE_INFO_FILE):
        try:
            import pandas as pd
            df_base = pd.read_excel(BASE_INFO_FILE, sheet_name='楼盘基础信息汇总')
            for _, r in df_base.iterrows():
                pn = str(r['项目名称']).strip()
                p_net = str(r['小学校网']).strip() if pd.notna(r['小学校网']) else ''
                s_net = str(r['中学校网']).strip() if pd.notna(r['中学校网']) else ''
                school_net_map[pn.lower()] = (p_net, s_net)
                if pd.notna(r.get('规范项目名')):
                    school_net_map[str(r['规范项目名']).strip().lower()] = (p_net, s_net)
        except Exception as e:
            print("读取 楼盘基础信息库.xlsx 异常:", e)

    MANUAL_OVERRIDE = {
        '皇廷汇': ('41校网', '九龙城区校网'),
        '天玺': ('31校网', '油尖旺区校网'),
        '喇沙汇': ('41校网', '九龙城区校网'),
        '耀爵台': ('34校网', '九龙城区校网'),
        '本木': ('31校网', '油尖旺区校网'),
        'Upper Prince': ('32校网', '油尖旺区校网'),
        '利奥坊．壹隅': ('32校网', '油尖旺区校网'),
        '吉喆': ('11校网', '中西区校网'),
        'Shouson Peak': ('18校网', '南区校网'),
        '浅水湾108': ('18校网', '南区校网'),
        'Twelve Peaks': ('11校网', '中西区校网'),
        '宾吉道3号': ('11校网', '中西区校网'),
        '活道1号': ('12校网', '湾仔区校网'),
        '远晴': ('16校网', '东区校网'),
        '环角道7号、9号及11号': ('18校网', '南区校网'),
        '339 Tai Hang Road': ('12校网', '湾仔区校网'),
        '启德海湾 1': ('34校网', '九龙城区校网'),
        '启德海湾 2': ('34校网', '九龙城区校网'),
        '天铸 (第2期)': ('34校网', '九龙城区校网'),
        '朗贤峰第IIA期': ('34校网', '九龙城区校网'),
        '朗贤峰第IIB期': ('34校网', '九龙城区校网'),
        'st. george\'s mansions': ('34校网', '九龙城区校网'),
        '维港汇 I': ('40校网', '深水埗区校网'),
        '维港汇 III': ('40校网', '深水埗区校网'),
        '花语海第1期': ('34校网', '九龙城区校网'),
        '花语海第2期': ('34校网', '九龙城区校网'),
    }

    print("3️⃣ 整合发展商、落成日期、校网归属、楼盘卖点、在售明细、最近成交记录、精准价格区间与套数统计...")
    def load_project_completion_dates():
        excel_path = "/Users/nb/google/Antigravity/工作/运营/楼盘字典/项目列表.xlsx"
        if not os.path.exists(excel_path):
            return {}
        try:
            wb = openpyxl.load_workbook(excel_path, read_only=True)
            ws = wb["项目列表"]
            rows = list(ws.iter_rows(values_only=True))
            headers = rows[0]
            idx_date = headers.index("落成日期")
            idx_name = headers.index("新盘名称\n（中文）")
            res = {}
            for r in rows[1:]:
                nm = str(r[idx_name] or "").strip()
                dt = str(r[idx_date] or "").strip()
                if dt and dt != "None":
                    res[nm] = dt
            return res
        except Exception as e:
            print("读取项目列表.xlsx落成日期异常:", e)
            return {}

    def load_project_developers():
        verified_devs = {
            # 启德 / 九龙东
            'Double Coast I': '中国海外 / 会德丰 / 恒基兆业 / 新世界',
            'Double Coast III': '中国海外 / 会德丰 / 恒基兆业 / 新世界',
            'Miami Quay I': '新世界 / 会德丰 / 恒基兆业 / 帝国集团',
            'Miami Quay II': '新世界 / 会德丰 / 恒基兆业 / 帝国集团',
            '天泷': '恒基兆业 / 中国海外 / 华懋集团 / 帝国集团 / 新世界 / 会德丰',
            '柏蔚森 I': '新世界 / 远东发展',
            '柏蔚森 II': '新世界 / 远东发展',
            '柏蔚森 III': '新世界 / 远东发展',
            '启德海湾 1': '嘉华国际 / 会德丰 / 中国海外',
            '启德海湾 2': '嘉华国际 / 会德丰 / 中国海外',
            '天玺．天': '新鸿基地产',
            '天玺．天第2期': '新鸿基地产',
            '天玺．海第1期': '新鸿基地产',
            '天玺．海第2A期': '新鸿基地产',
            '天玺．海第2B期': '新鸿基地产',
            '维港1号': '中国海外',
            '维港．双钻': '中国海外',
            '维港．湾畔第1A期': '中国海外 / 恒基兆业 / 九龙仓 / 嘉华国际',
            '维港．湾畔第1B期': '中国海外 / 恒基兆业 / 九龙仓 / 嘉华国际',
            '维港．湾畔第2B期': '中国海外 / 恒基兆业 / 九龙仓 / 嘉华国际',
            '沄璟': '华润置地 / 保利置业',
            'The Henley I': '恒基兆业',
            'The Henley II': '恒基兆业',
            'The Henley III': '恒基兆业',
            'Henley Park': '恒基兆业',
            'Monaco 第1期': '会德丰',
            'Monaco One': '会德丰',
            'Monaco Marine': '会德丰',
            
            # 红磡 / 何文田 / 土瓜湾 / 西九龙
            '必嘉坊．曦汇': '恒基兆业',
            '必嘉坊．迎汇': '恒基兆业',
            '首岸第1期': '恒基兆业 / 希慎兴业 / 帝国集团',
            '首岸第2期': '恒基兆业 / 希慎兴业 / 帝国集团',
            '首岸第3期': '恒基兆业 / 希慎兴业 / 帝国集团',
            '首岸第4期': '恒基兆业 / 希慎兴业 / 帝国集团',
            '首汇': '恒基兆业',
            'The Haddon': '恒基兆业',
            '映汇': '恒基兆业',
            '壹沐第1期': '恒基兆业',
            '壹沐第2期': '恒基兆业',
            '南首': '恒基兆业',
            '文曜': '建灏地产',
            '瑧博': '新世界',
            '朗贤峰第IIA期': '鹰君集团 / 港铁',
            '朗贤峰第IIB期': '鹰君集团 / 港铁',
            '瑜一第IC期': '华懋集团 / 港铁',
            '瑜一．天海': '华懋集团 / 港铁',
            '维港汇 I': '会德丰 / 信和置业 / 嘉华国际 / 世茂 / 爪哇集团',
            '维港汇 III': '会德丰 / 信和置业 / 嘉华国际 / 世茂 / 爪哇集团',
            '汇玺III': '新鸿基地产 / 港铁',
            '天玺': '新鸿基地产',
            'st. george\'s mansions': '中电集团 / 信和',
            'St. George\'s Mansions': '中电集团 / 信和',
            '傲玟': '高银',
            '名钻': '益兆集团',
            '128 Waterloo': '莱蒙国际 / 俊和',
            'Prince Central': '新鸿基地产',
            '太子道西233号': '新鸿基地产 / 会德丰',
            '天铸 (第2期)': '新鸿基地产',
            '明寓': '佳明',
            '耀爵台': '光大控股 / 柏汇国际',
            '芳菲': '协成行',
            '誉林': '永义国际',
            '隆敍': '志德投资',
            'The Boundary': '乐风集团',
            'Upper Prince': '乐风集团',
            '利奥坊．凯岸': '恒基兆业',
            '利奥坊．壹隅': '恒基兆业',
            '利奥坊．曦岸': '恒基兆业',
            '本木': '恒基兆业',
            '映居': '建灏地产 / 市建局',
            '原舍': '东立地产',
            'VAU Residence': '万科香港',
            '悦麓': 'TKS LTD.',
            '叡璟I': '华润置地',
            '叡璟II': '华润置地',
            '泓璟': '辉宇国际',
            'Elize Park': '乐风集团',
            
            # 九龙塘 / 笔架山 / 又一村
            '缇外': '嘉里建设',
            '皇廷汇': '天滙地产',
            '喇沙汇': '汉国置业',
            'Park College': '兴胜创建 / TPG',
            'Beacon Peak 第1期': '世茂集团',
            '滶蕴': '新世界发展',

            # 港岛 - 中半山 / 西半山 / 湾仔 / 铜锣湾
            '雅盈峰': '盈大地产 / 资本策略',
            '21 Borrett Road 第1期': '长江实业',
            'Upper Central': '禹洲集团',
            '大学阁': '华懋集团',
            '天御第1期': '恒基兆业 / 新世界',
            '天御第2期': '恒基兆业 / 新世界',
            '宝峰': '泛海国际 / 资本策略 / 德祥地产',
            '帝汇豪庭': '恒基兆业',
            '应天': '长江实业',
            '罗便臣道80号': '南丰集团',
            'Kennedy 38': '新鸿基 / 会德丰 / 恒基兆业',
            '吉喆': '上海商业银行',
            '海嵎': '永义国际',
            'Woodis': '恒基兆业',
            '活道1号': '盈信控股',
            '15 Western Street': '万科香港',
            '42 Tung St.': 'The Development Studio',
            '肇辉台6号': '嘉华国际',

            # 港岛 - 跑马地 / 渣甸山 / 东区
            'Eight Kwai Fong Happy Valley': '远中集团',
            '蓝塘道23-39号': '恒隆地产',
            '皇第': '泛海国际 / 尚家生活 / 高富诺',
            '峻誉．渣甸山': '中信泰富',
            '339 Tai Hang Road': '丽新发展',
            '101 Kings Road': '宏安地产',
            '君誉峰': '恒基兆业',
            '嘉居．天后': '嘉华国际',
            '柏蔚山': '新世界 / 培新集团',
            '海璇': '新鸿基地产',
            '海璇 II': '新鸿基地产',
            '海璇 II (第2B-3期)': '新鸿基地产',
            '海德园第1期': '太古地产',
            '海德园第2期': '太古地产',
            '君豪峰': '恒基兆业',
            'One Eighty': '三顺 / 日联',
            '傲华': '协成行',
            '远晴': '协成行',

            # 港岛 - 黄竹坑 / 香港仔 / 鸭脷洲 / 薄扶林
            '晋环': '路劲地产 / 平安不动产 / 港铁',
            '扬海': '嘉里建设 / 信和置业 / 太古地产 / 港铁',
            '海盈山第4A期': '嘉里建设 / 信和置业 / 太古地产 / 港铁',
            '海盈山第4B期': '嘉里建设 / 信和置业 / 太古地产 / 港铁',
            '滶晨': '新世界 / 帝国集团 / 资本策略 / 丽新发展 / 港铁',
            '滶晨 II': '新世界 / 帝国集团 / 资本策略 / 丽新发展 / 港铁',
            'Deep Water South 第6A期': '会德丰 / 港铁',
            'Deep Water South 第6B期': '会德丰 / 港铁',
            'Blue Coast': '长江实业 / 港铁',
            'Blue Coast II': '长江实业 / 港铁',
            '花语海第1期': '长江实业 / 港铁',
            '花语海第2期': '长江实业 / 港铁',
            '凯玥': '龙光集团 / 合景泰富',
            'PORTO': '宏安地产',
            '弦岸': '庄士机构',
            '璟南': '大昌地产',
            '逸南': '恒基兆业',
            'Victoria Coast': '华懋集团',
            'Mount Pokfulam': '宏安地产 / 锦华地产',

            # 港岛 - 寿臣山 / 浅水湾 / 赤柱 / 大潭
            '寿臣山15号': '英皇国际 / 中渝置地 / 明发集团 / 尚家生活',
            '深水湾径8号': '南丰集团 / 丰资源',
            'No.3 Repulse Bay Road': '新鸿基地产',
            'Shouson Peak': '新鸿基地产',
            '浅水湾108': '大昌集团',
            '皇府湾': '新鸿基地产',
            'Twelve Peaks': '新鸿基地产',
            '宾吉道3号': '大昌集团',
            'Mount Nicholson I': '九龙仓 / 南丰集团',
            'Mount Nicholson II': '九龙仓 / 南丰集团',
            'Central Peak I': '新鸿基地产',
            'Central Peak II': '新鸿基地产',
            '博峰': '丽新集团',
            '50 Stanley Village Road': '新鸿基地产',
            'One Stanley': '建灏地产',
            'Solace': '东立地产',
            '海天径 1号, 舂磡角道 44, 46, 48, 50号': '信德集团',
            '环角道7号、9号及11号': '培新集团',
            '红山半岛 - B区': '信和 / 华懋集团 / 亚证地产',
            '红山半岛 - C区': '信和 / 华懋集团 / 亚证地产',

            # 其他
            'The Monet 第1期': '会德丰',
            'The Monet 第2期': '会德丰',
            'The Monet 第3期': '会德丰',
            '海瑅湾 I': '信和置业 / 嘉里建设 / 嘉华国际 / 招商置地 / 港铁',
            '海瑅湾 II': '信和置业 / 嘉里建设 / 嘉华国际 / 招商置地 / 港铁',
        }
        res = dict(verified_devs)
        excel_path = "/Users/nb/google/Antigravity/工作/运营/楼盘字典/项目列表.xlsx"
        if os.path.exists(excel_path):
            try:
                wb = openpyxl.load_workbook(excel_path, read_only=True)
                ws = wb["项目列表"]
                rows = list(ws.iter_rows(values_only=True))
                headers = rows[0]
                idx_dev = headers.index("开发商")
                idx_name = headers.index("新盘名称\n（中文）")
                for r in rows[1:]:
                    nm = str(r[idx_name] or "").strip()
                    dv = str(r[idx_dev] or "").strip()
                    if nm and dv and dv != "None" and nm not in res:
                        res[nm] = dv
            except Exception as e:
                print("读取项目列表.xlsx开发商异常:", e)
        return res

    project_completion_map = load_project_completion_dates()
    project_developer_map = load_project_developers()
    selling_points_map = load_selling_points()
    t_chars = '號峯滙敍灣瓏鑽鑄賢譽啟璽維奧滶蘊匯悅寶應羅壽淺徑賓園華遠藍語區築緹瀧澐雙鑽閣揚晉凱島輝臺連東紅臺豐環學區萬濱嶺軒廣場頭門閣薈'
    s_chars = '号峰汇叙湾珑钻铸贤誉启玺维奥滶蕴汇悦宝应罗寿浅径宾园华远蓝语区筑缇泷沄双钻阁扬晋凯岛辉台连东红台丰环学区万滨岭轩广场头门阁荟'
    trans = str.maketrans(t_chars, s_chars)

    tx_attached = 0
    price_attached = 0
    sp_attached = 0
    for p in projects:
        pname = p["name"].strip()
        dj_p = data_json_map.get(pname, {})

        # 匹配核心特色卖点
        sp_text = selling_points_map.get(pname)
        if not sp_text:
            norm_name = pname.translate(trans).replace(' ', '').replace('　', '').replace('．', '.').lower()
            sp_text = selling_points_map.get(norm_name)
        if sp_text:
            p["selling_points"] = sp_text
            sp_attached += 1
            if not p.get("mainland_selling_points"):
                p["mainland_selling_points"] = sp_text
        
        full_text = " ".join(filter(None, [
            p.get("selling_points", ""),
            p.get("mainland_selling_points", ""),
            p.get("reason", ""),
            dj_p.get("mainland_selling_points", ""),
            dj_p.get("reason", "")
        ]))

        p["developer"] = extract_developer(pname, full_text, project_developer_map)
        comp_date = extract_completion_date(full_text)
        dt_excel = project_completion_map.get(pname)
        if not dt_excel:
            for k, v in project_completion_map.items():
                if k in pname or pname in k:
                    dt_excel = v
                    break

        is_ready_explicit = any(k in full_text for k in ['现楼发售', '现楼实景', '现楼交付', '即买即住', '即买即收租'])
        is_forward = False
        pt_desc = "现房 (现楼发售)"
        comp_date = "现楼发售 (即买即住)"

        if is_ready_explicit:
            is_forward = False
            pt_desc = "现房 (现楼发售)"
            comp_date = "现楼发售 (即买即住)"
        elif dt_excel:
            m = re.search(r"20(2[5-9]|3[0-9])", dt_excel)
            if m:
                yr = int("20" + m.group(1))
                is_forward = True
                pt_desc = f"期房 (预计{yr}年入伙)"
                comp_date = f"预计{yr}年入伙"
            else:
                is_forward = False
                pt_desc = f"现房 (现楼发售)"
                comp_date = f"现楼 ({dt_excel})"
        elif p.get("is_coming_soon"):
            is_forward = True
            pt_desc = "期房 (即将发售)"
            comp_date = "即将发售"
        elif any(k in full_text for k in ['预售楼花', '楼花发售', '一手楼花']):
            is_forward = True
            pt_desc = "期房 (楼花)"
            comp_date = "预售楼花"
        else:
            is_forward = False
            pt_desc = "现房 (现楼发售)"
            comp_date = "现楼发售 (即买即住)"

        p["completion_date"] = comp_date
        p["property_type"] = "forward" if is_forward else "ready"
        p["property_type_desc"] = pt_desc

        # 匹配校网归属
        sn_info = None
        if pname in MANUAL_OVERRIDE:
            sn_info = MANUAL_OVERRIDE[pname]
        elif pname.lower() in school_net_map:
            sn_info = school_net_map[pname.lower()]
        else:
            for k, val in school_net_map.items():
                if k in pname.lower() or pname.lower() in k:
                    sn_info = val
                    break
        if sn_info:
            p["primary_school_net"] = sn_info[0]
            p["secondary_school_net"] = sn_info[1]
            if sn_info[0] and sn_info[1]:
                p["school_net"] = f"{sn_info[0]} / {sn_info[1]}"
            elif sn_info[0]:
                p["school_net"] = sn_info[0]

        # 匹配在售数据与物理销控套数统计
        sc = sales_data.get(pname)
        ustats = unit_stats_data.get(pname)
        if not sc or not ustats:
            for k, val in sales_data.items():
                if k in pname or pname in k:
                    sc = val
                    ustats = unit_stats_data.get(k)
                    break
        p["sales_control_summary"] = sc if sc else []
        if ustats and ustats.get("total_units", 0) > 0:
            p["total_units"] = ustats["total_units"]
            p["sold_units"] = ustats["sold_units"]
            p["sale_units"] = ustats["sale_units"]
            p["tower_count"] = ustats.get("tower_count", 1)
            p["tower_names"] = ustats.get("tower_names", [])
        if not p.get("tower_count"):
            p["tower_count"] = 1

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

        # 匹配精准价格区间与在售最低价描述
        pb = bounds_data.get(pname)
        if not pb:
            for k, val in bounds_data.items():
                if k in pname or pname in k:
                    pb = val
                    break
        
        if pb and pb.get('has_price_data'):
            p["min_price_wan"] = pb["min_price_wan"]
            p["max_price_wan"] = pb["max_price_wan"]
            p["has_price_data"] = True
            p["excel_min_price_desc"] = pb["desc"]
            price_attached += 1
        else:
            # 兜底：优先从文本描述中提取具体数值，其次根据 price_tier 设定
            tier = p.get("price_tier") or dj_p.get("price_tier", "")
            p_desc = " ".join(filter(None, [
                p.get("total_price_desc"),
                dj_p.get("total_price_desc"),
                p.get("basic_price_desc"),
                dj_p.get("basic_price_desc")
            ]))
            
            min_w, max_w = None, None
            if p_desc:
                m_yi = re.search(r'(\d+(?:\.\d+)?)\s*亿', p_desc)
                m_wan = re.search(r'(\d+(?:\.\d+)?)\s*万', p_desc)
                if m_yi:
                    val_w = float(m_yi.group(1)) * 10000
                    min_w, max_w = round(val_w * 0.95, 2), round(val_w * 1.3, 2)
                elif m_wan:
                    val_w = float(m_wan.group(1))
                    min_w, max_w = round(val_w * 0.95, 2), round(val_w * 1.3, 2)
            
            if min_w is None:
                if tier in ["1000-down", "1000万以内"]:
                    min_w, max_w = 0, 1000
                elif tier in ["1000-2000", "1000-2000万"]:
                    min_w, max_w = 1000, 2000
                elif tier in ["2000-5000", "2000-5000万"]:
                    min_w, max_w = 2000, 5000
                elif tier in ["5000-10000", "5000-1亿"]:
                    min_w, max_w = 5000, 10000
                elif tier in ["10000+", "10000-up", "1亿以上"]:
                    min_w, max_w = 10000, 99999
            
            p["min_price_wan"] = min_w
            p["max_price_wan"] = max_w
            p["has_price_data"] = (min_w is not None)
            if p["has_price_data"]:
                price_attached += 1
                p["excel_min_price_desc"] = f"预估约: ${int(min_w)}万 - ${int(max_w)}万"
            elif p.get("sell_status") == "coming_soon":
                p["excel_min_price_desc"] = "即将发售 / 待定"
            else:
                p["excel_min_price_desc"] = "招标发售 / 详见价单"

    print(f"4️⃣ 保存更新至 {COORDS_FILE} (共 {len(projects)} 个楼盘，挂载楼盘卖点 {sp_attached} 个，挂载最近成交数据 {tx_attached} 个)...")
    with open(COORDS_FILE, "w", encoding="utf-8") as f:
        json.dump(projects, f, ensure_ascii=False, indent=2)
    print("🎉 数据升级完成！")

if __name__ == "__main__":
    main()
