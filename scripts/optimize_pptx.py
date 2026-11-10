#!/usr/bin/env python3
"""
optimize_pptx.py — 体积治理：媒体降采样 + 无损重打包。

背景（实测）：学校模板单张装饰图可达 42MB，两套模板分别 243MB / 104MB；
学生要微信/邮件/U 盘/教务系统传输，交付前必须压到 <15MB。

做法：只替换 ppt/media 下的大图字节（同格式重编码，不改扩展名、不改关系），
其余 zip 条目原样保留，因此不会破坏 pptx 结构。

用法:
    python3 optimize_pptx.py <in.pptx> -o <out.pptx> [--max-width 1920] [--quality 82] [--min-kb 800]
"""
import argparse
import io
import sys
import zipfile
from pathlib import Path

from PIL import Image

RASTER_EXT = (".png", ".jpg", ".jpeg")


def optimize(src: Path, dst: Path, max_width=1920, quality=82, min_kb=800):
    zin = zipfile.ZipFile(src)
    stats, total_before, total_after = [], 0, 0
    dst.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            total_before += len(data)
            new_data = data
            if (info.filename.startswith("ppt/media/")
                    and info.filename.lower().endswith(RASTER_EXT)
                    and len(data) > min_kb * 1024):
                try:
                    im = Image.open(io.BytesIO(data))
                    orig_size = im.size
                    if im.width > max_width:
                        im = im.resize((max_width, round(im.height * max_width / im.width)),
                                       Image.LANCZOS)
                    buf = io.BytesIO()
                    if info.filename.lower().endswith((".jpg", ".jpeg")):
                        if im.mode not in ("RGB", "L"):
                            im = im.convert("RGB")
                        im.save(buf, "JPEG", quality=quality, optimize=True, progressive=True)
                    else:
                        im.save(buf, "PNG", optimize=True)
                    cand = buf.getvalue()
                    if len(cand) < len(data):
                        new_data = cand
                    stats.append({
                        "file": Path(info.filename).name,
                        "beforeMB": round(len(data) / 1048576, 2),
                        "afterMB": round(len(new_data) / 1048576, 2),
                        "size": f"{orig_size[0]}x{orig_size[1]}",
                    })
                except Exception as e:  # 矢量图/损坏图直接跳过
                    stats.append({"file": Path(info.filename).name, "skipped": str(e)[:60]})
            total_after += len(new_data)
            zout.writestr(info, new_data)
    zin.close()
    return {
        "input": str(src),
        "output": str(dst),
        "inputMB": round(total_before / 1048576, 1),
        "outputMB": round(total_after / 1048576, 1),
        "saved": f"{(1 - total_after / total_before) * 100:.0f}%",
        "media": stats,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("--max-width", type=int, default=1920)
    ap.add_argument("--quality", type=int, default=82)
    ap.add_argument("--min-kb", type=int, default=800)
    args = ap.parse_args()
    rep = optimize(Path(args.input), Path(args.output), args.max_width, args.quality, args.min_kb)
    print(f"{rep['inputMB']} MB → {rep['outputMB']} MB (省 {rep['saved']})")
    for m in rep["media"][:10]:
        if "skipped" in m:
            print(f"  skip {m['file']}")
        else:
            print(f"  {m['file']}: {m['beforeMB']} → {m['afterMB']} MB  ({m['size']})")
    if rep["outputMB"] > 15:
        print(f"[warn] 仍 {rep['outputMB']} MB，可降低 --max-width 或 --quality")
        sys.exit(1)


if __name__ == "__main__":
    main()
