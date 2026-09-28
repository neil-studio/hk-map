#!/usr/bin/env python3
"""
楼盘高清原图处理与缩略图流水线脚本
1. 递归扫描 images/projects/original/（支持项目子文件夹或根目录直接放图）
2. 智能匹配 hk_project_coords.json 中的楼盘项目（自动兼容期数如 The Monet 第1/2/3期）
3. 自动生成保持原始比例、无损完整展示的极速 WebP 缩略图到 images/projects/thumb/
4. 建立 project_images_manifest.json 供前端毫秒级加载
"""

import os, glob, json
from PIL import Image, ImageOps

SANDBOX_DIR = os.path.dirname(os.path.abspath(__file__))
ORIGINAL_DIR = os.path.join(SANDBOX_DIR, "images", "projects", "original")
THUMB_DIR = os.path.join(SANDBOX_DIR, "images", "projects", "thumb")
MANIFEST_FILE = os.path.join(SANDBOX_DIR, "project_images_manifest.json")
COORDS_FILE = os.path.join(SANDBOX_DIR, "hk_project_coords.json")

os.makedirs(ORIGINAL_DIR, exist_ok=True)
os.makedirs(THUMB_DIR, exist_ok=True)

SUPPORTED_EXTS = ('.jpg', '.jpeg', '.png', '.webp', '.bmp', '.heic', '.tif', '.tiff')

def load_all_project_names():
    if not os.path.exists(COORDS_FILE):
        return []
    try:
        with open(COORDS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return [p["name"] for p in data if "name" in p]
    except Exception as e:
        print("读取 hk_project_coords.json 异常:", e)
        return []

def match_projects(candidate_name, all_projects):
    c_clean = candidate_name.replace(' ', '').replace('_', '').replace('-', '').replace('·', '').replace('．', '').lower()
    matches = []
    # 1. 精确匹配
    for p in all_projects:
        p_clean = p.replace(' ', '').replace('_', '').replace('-', '').replace('·', '').replace('．', '').lower()
        if p_clean == c_clean:
            matches.append(p)
    if matches:
        return matches

    # 2. 包含匹配 (例如 the monet 匹配 The Monet 第1期, The Monet 第2期...)
    for p in all_projects:
        p_clean = p.replace(' ', '').replace('_', '').replace('-', '').replace('·', '').replace('．', '').lower()
        if c_clean in p_clean:
            matches.append(p)
    if matches:
        return matches

    # 3. 反向包含匹配
    for p in all_projects:
        p_clean = p.replace(' ', '').replace('_', '').replace('-', '').replace('·', '').replace('．', '').lower()
        if p_clean in c_clean:
            matches.append(p)
    return matches

def process_images():
    print(f"🔍 扫描高清原图目录: {ORIGINAL_DIR}")
    all_projects = load_all_project_names()

    # 递归查找所有图片
    image_entries = []
    for root, dirs, files in os.walk(ORIGINAL_DIR):
        for f in files:
            if f.startswith('.') or f == 'Icon' or f.endswith('.gitkeep'):
                continue
            ext = os.path.splitext(f)[1].lower()
            if ext in SUPPORTED_EXTS:
                full_path = os.path.join(root, f)
                rel_dir = os.path.relpath(root, ORIGINAL_DIR)
                folder_name = rel_dir if rel_dir != '.' else ''
                base_name = os.path.splitext(f)[0].strip()
                # 候选名称优先取子文件夹名，否则取文件名
                candidate = folder_name if folder_name else base_name
                image_entries.append((candidate, full_path, f))

    if not image_entries:
        print("ℹ️ original 目录中暂无图片，请将楼盘高清原图放入该目录（支持子文件夹或直接放图）。")
        with open(MANIFEST_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=2)
        return

    manifest = {}
    processed_count = 0

    for candidate, orig_path, fname in sorted(image_entries, key=lambda x: x[0]):
        matched = match_projects(candidate, all_projects)
        if not matched:
            matched = [candidate]

        try:
            with Image.open(orig_path) as img:
                img = ImageOps.exif_transpose(img)
                w, h = img.size

                # 处理透明度
                if img.mode in ("RGBA", "P", "LA"):
                    img = img.convert("RGBA")
                    bg = Image.new("RGBA", img.size, (15, 23, 42, 255))
                    img = Image.alpha_composite(bg, img).convert("RGB")
                else:
                    img = img.convert("RGB")

                # 生成缩略图：最大边长 360px，比例完整保留
                thumb_img = img.copy()
                max_edge = 360
                if max(w, h) > max_edge:
                    thumb_img.thumbnail((max_edge, max_edge), Image.Resampling.LANCZOS)

                orig_rel = os.path.relpath(orig_path, SANDBOX_DIR)
                orig_size_kb = os.path.getsize(orig_path) / 1024

                for proj_name in matched:
                    thumb_filename = f"{proj_name}.webp"
                    thumb_dest = os.path.join(THUMB_DIR, thumb_filename)
                    thumb_img.save(thumb_dest, "WEBP", quality=85, method=6)
                    thumb_size_kb = os.path.getsize(thumb_dest) / 1024

                    manifest[proj_name] = {
                        "original": orig_rel,
                        "thumb": f"images/projects/thumb/{thumb_filename}",
                        "width": w,
                        "height": h,
                        "aspect_ratio": round(w / h, 3)
                    }
                    processed_count += 1
                    print(f"  ✅ 映射楼盘 [{proj_name}]: {w}x{h} ({orig_size_kb:.1f} KB) ➔ 缩略图 ({thumb_size_kb:.1f} KB)")

        except Exception as e:
            print(f"  ❌ 处理图片失败 [{orig_path}]: {e}")

    with open(MANIFEST_FILE, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print(f"\n🎉 处理完毕！共生成 {processed_count} 个楼盘封面映射，已更新索引至 project_images_manifest.json")

if __name__ == "__main__":
    process_images()
