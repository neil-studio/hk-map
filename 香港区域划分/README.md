# 香港 156 个真实微型规划片区 GIS 空间数据集与集成说明

## 📦 软件包内容清单

| 文件名 | 格式 | 适用场景与说明 |
| :--- | :--- | :--- |
| **`hong_kong_micro_districts.geojson`** | GeoJSON (WGS84) | **标准地理数据底表**。涵盖全港 156 个细分微型片区，100% 互斥无重叠、无生硬直线、沿真实道路/山脊/海岸线收敛。适用于 Leaflet、Mapbox、Google Maps、QGIS。 |
| **`hong_kong_micro_districts_gcj02.geojson`** | GeoJSON (GCJ-02) | **国测局火星坐标系版本**。适用于高德地图 (AMap)、腾讯地图、百度地图，防止国内地图坐标偏移。 |
| **`hong_kong_micro_districts.kml`** | KML (XML) | 适用于 **Google Earth (3D)**、**Google My Maps (我的地图)** 免代码直接上传可视化。 |
| **`integration_snippet.js`** | JavaScript | **针对 neil-studio.github.io/hk-map 的即插即用代码**（静默底图渲染 + 顶栏下拉菜单飞行聚焦）。 |
| **`preview.html`** | HTML/JS | **本地独立预览查看器**，双击即可在浏览器中查看全港 156 个片区和全港标杆楼盘测试。 |

---

## 🗺️ 数据来源与核心特性

1. **官方权威数据源**：
   * 香港特区政府城市规划委员会（TPB）法定规划图则（OZP）
   * 香港规划署（PlanD）法定小规划统计区（TPU 2021）
2. **纯粹地理无缝细分**：
   * **港岛南区细分**：浅水湾及舂磡角、深水湾、寿臣山、鸭脷洲、黄竹坑及深湾、香港仔、田湾、赤柱等；
   * **半山及山顶细分**：西半山、中半山、东半山、大坑及渣甸山、山顶（含金马伦山、布力径、聂高信山 Mount Nicholson）；
   * **西九龙细分**：九龙站及柯士甸（含西九文化区）、奥运及大角咀、南昌及西九龙海滨（纯陆地无海水）、昂船洲；
   * **核心区合并**：中环与金钟合并为统一【中环及金钟】；湾仔核心与湾仔北会展合并为统一【湾仔】；
   * **坚尼地城细分**：坚尼地城（市区海滨）与摩星岭（山体豪宅）分离；
3. **空间拓扑质量**：
   * 片区相互重叠数：**严格为 0（100% 全域互斥）**；
   * 人工生硬切割线：**0 条**；
   * 海面多余水域：**已全部剔除贴合海岸线**。

---

## 🚀 快速集成到 GitHub 网站 (neil-studio.github.io/hk-map)

### 步骤 1：放入数据文件
将 `hong_kong_micro_districts.geojson` 复制到您的网站仓库根目录中。

### 步骤 2：引入 JS 代码
将 `integration_snippet.js` 中的代码直接复制到您的页面脚本中，或者在 HTML 底部引入：
```html
<script src="./integration_snippet.js"></script>
```

### 步骤 3：绑定顶栏下拉菜单
将您现有的下拉菜单 `<select id="anchor-district">` 设置为：
```html
<select class="anchor-select-glass" id="anchor-district" onchange="onAnchorDistrictSelect(this.value)">
  <option value="">🏙️ 核心豪宅商圈...</option>
</select>
```

---

## 📊 属性字段结构说明
每个片区 Polygon 包含以下 GeoJSON 属性（`feature.properties`）：
```json
{
  "id": "TPU-111-KT",
  "plan_no": "TPU-111-KT",
  "micro_district": "坚尼地城",
  "micro_district_en": "Kennedy Town",
  "main_area": "港岛中西区",
  "color": "#7c3aed",
  "source": "香港特区政府规划署 (PlanD) 官方法定小规划统计区 (TPU)"
}
```
