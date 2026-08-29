// =========================================================================
// 🗺️ 全港 156 个微型片区底图加载 + 顶栏下拉菜单聚焦联动（0 鼠标弹窗干扰）
// 适用项目：Leaflet.js 地图网站（如 neil-studio.github.io/hk-map）
// =========================================================================
const districtBoundsMap = {}; // 存储每个片区的地理边界坐标，供下拉框飞行使用

function initDistrictsAndDropdown() {
  fetch('./hong_kong_micro_districts.geojson')
    .then(res => res.json())
    .then(data => {
      // 1. 在地图上静默绘制 156 个片区红线（关闭鼠标交互，点击直接穿透到楼盘）
      L.geoJSON(data, {
        interactive: false, // 关键：关闭交互，鼠标事件直接穿透至楼盘标记
        style: function(feature) {
          return {
            color: '#e11d48',       // 真实规划红线
            weight: 1.2,            // 边线粗细
            opacity: 0.75,          // 边线不透明度
            fillColor: feature.properties.color || '#3b82f6',
            fillOpacity: 0.08       // 低透明度底色，不遮挡底层地图与楼盘
          };
        },
        onEachFeature: function(feature, layer) {
          const name = feature.properties.micro_district;
          districtBoundsMap[name] = layer.getBounds();
        }
      }).addTo(map);

      // 2. 自动填充顶栏现有的 "🏙️ 核心豪宅商圈..." 下拉菜单
      populateAnchorDistrictDropdown(data.features);
    })
    .catch(err => console.error("加载片区数据失败:", err));
}

// 自动按大区分组填充下拉框
function populateAnchorDistrictDropdown(features) {
  const select = document.getElementById('anchor-district');
  if (!select) return;

  select.innerHTML = '<option value="">🏙️ 核心豪宅商圈 / 片区 (156个)...</option>';

  // 按大区（港岛中西区、港岛南区、九龙西等）归类分组
  const groups = {};
  features.forEach(f => {
    const main = f.properties.main_area || '其他片区';
    const name = f.properties.micro_district;
    if (!groups[main]) groups[main] = [];
    if (!groups[main].includes(name)) groups[main].push(name);
  });

  for (const [mainArea, names] of Object.entries(groups)) {
    const optgroup = document.createElement('optgroup');
    optgroup.label = mainArea;
    names.sort().forEach(name => {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      optgroup.appendChild(opt);
    });
    select.appendChild(optgroup);
  }
}

// 3. 监听下拉菜单选择事件：平滑飞行聚焦到该片区红线
function onAnchorDistrictSelect(selectedDistrict) {
  if (!selectedDistrict) return;
  const bounds = districtBoundsMap[selectedDistrict];
  if (bounds) {
    map.fitBounds(bounds, {
      padding: [40, 40],
      maxZoom: 15,
      animate: true,
      duration: 1.0
    });
  }
}

// 页面加载完成后自动执行
document.addEventListener('DOMContentLoaded', initDistrictsAndDropdown);
