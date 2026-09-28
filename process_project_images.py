#!/usr/bin/env python3
"""
楼盘高清原图处理与缩略图流水线脚本
1. 遍历 images/projects/original/ 中的高清大图
2. 自动生成保持完整长宽比的轻量 WebP 缩略图到 images/projects/thumb/
3. 建立 images_manifest.json 供前端极速识别并加载
"""

import os, glob, json
from PIL import Image, ImageOps

SANDBOX_DIR = os.path.dirname(os.path.abspath(__file__))
ORIGINAL_DIR = os.path.join(SANDBOX_DIR, "images", "projects", "original")
THUMB_DIR = os.path.join(SANDBOX_DIR, "images", "projects", "thumb")
MANIFEST_FILE = os.path.join(SANDBOX_DIR, "project_images_manifest.json")

os.makedirs(ORIGINAL_DIR, exist_ok=True)
os.makedirs(THUMB_DIR, exist_ok=True)

SUPPORTED_EXTS = ('.jpg', '.jpeg', '.png', '.webp', '.bmp', '.heic', '.tif', '.tiff')

def process_images():
    print(f"🔍 扫描高清原图目录: {ORIGINAL_DIR}")
    files = [f for f in os.listdir(ORIGINAL_DIR) if os.path.splitext(f)[1].lower() in SUPPORTED_EXTS]
    
    if not files:
        print("ℹ️ original 目录中暂无图片，请将楼盘高清原图放入该目录（命名为：楼盘名称.jpg）。")
        # 即使无图，也写入空 manifest
        with open(MANIFEST_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=2)
        return

    manifest = {}
    success_count = 0

    for fname in sorted(files):
        base_name, ext = os.path.splitext(fname)
        proj_name = base_name.strip()
        orig_path = os.path.join(ORIGINAL_DIR, fname)
        thumb_name = f"{proj_name}.webp"
        thumb_path = os.path.join(THUMB_DIR, thumb_name)

        try:
            with Image.open(orig_path) as img:
                # 自动根据 EXIF 纠正方向
                img = ImageOps.exif_transpose(img)
                w, h = img.size

                # 转换为 RGB
                if img.mode in ("RGBA", "P", "LA"):
                    img = img.convert("RGBA")
                    bg = Image.new("RGBA", img.size, (15, 23, 42, 255)) # 暗色底
                    img = Image.alpha_composite(bg, img).convert("RGB")
                else:
                    img = img.convert("RGB")

                # 生成缩略图：最大边长 360px，保持原始长宽比完整不变形
                max_edge = 360
                if max(w, h) > max_edge:
                    img.thumbnail((max_edge, max_edge), Image.Resampling.LANCZOS)

                # 保存为优化的 WebP
                img.save(thumb_path, "WEBP", quality=85, method=6)
                orig_size_kb = os.path.getsize(orig_path) / 1024
                thumb_size_kb = os.path.getsize(thumb_path) / 1024

                manifest[proj_name] = {
                    "original": f"images/projects/original/{fname}",
                    "thumb": f"images/projects/thumb/{thumb_name}",
                    "width": w,
                    "height": h,
                    "aspect_ratio": round(w / h, 3)
                }
                success_count += 1
                print(f"  ✅ [{proj_name}]: {w}x{h} ({orig_size_kb:.1f} KB) ➔ 缩略图 {thumb_size_kb:.1f} KB")

        except Exception as e:
            print(f"  ❌ 处理图片失败 [{fname}]: {e}")

    with open(MANIFEST_FILE, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print(f"\n🎉 处理完毕！共处理 {success_count} 张楼盘图片，已更新索引至 project_images_manifest.json")

if __name__ == "__main__":
    process_images()
