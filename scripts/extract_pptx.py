#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""extract_pptx.py — 提取 .pptx 逐页内容为结构化 Markdown，供学术 PPT 审查使用。

用法:
    python extract_pptx.py <file.pptx> [-o output.md] [--imgs <dir>]

输出: 每页一节，含标题、层级正文、表格、备注、视觉元素计数与文本量统计。
      --imgs 将嵌入图片按页导出到指定目录（P03_00_显示尺寸.png 命名，
      自动跳过显示尺寸过小的图标/logo），供多模态核验图形页。
局限: SmartArt、艺术字、嵌入图片内的文字无法提取；此类页面若文本为空，
     会被标注为"需人工核对视觉信息"。

依赖: python-pptx（缺失时先安装: pip install python-pptx）
"""
import argparse
import os
import sys


def build_slide_section(slide, idx, total, img_dir=None, img_counter=None):
    """提取单页内容，返回 (markdown 行列表, 统计信息 dict)。"""
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    lines = []
    stats = {"chars": 0, "pics": 0, "charts": 0, "tables": 0, "notes": 0}

    def save_picture(shape):
        """导出嵌入图片，返回文件名；显示尺寸过小（图标/logo）则跳过。"""
        if img_dir is None:
            return None
        try:
            img = shape.image
        except Exception:
            return None
        disp_px = ((shape.width or 0) // 9525) * ((shape.height or 0) // 9525)
        if disp_px < 30000:  # 约 173×173 px 以下视为装饰元素
            return None
        img_counter[0] += 1
        fn = f"P{idx:02d}_{img_counter[0]:02d}_{disp_px // 1000}k.{img.ext}"
        with open(os.path.join(img_dir, fn), "wb") as f:
            f.write(img.blob)
        return fn

    # 标题占位符（记录 shape_id，正文提取时跳过以免重复）
    title_shape = None
    title_text = None
    title_id = None
    try:
        if slide.shapes.title is not None:
            title_shape = slide.shapes.title
            title_text = (title_shape.text or "").strip()
            title_id = title_shape.shape_id
    except Exception:
        pass
    lines.append(f"## Slide {idx}/{total}")
    lines.append(f"标题: {title_text or '(无标题占位符)'}")

    body_lines = []
    pic_files = []

    def collect(shapes):
        for shape in shapes:
            if title_id is not None and getattr(shape, "shape_id", None) == title_id:
                continue
            try:
                st = shape.shape_type
            except Exception:
                st = None
            if st == MSO_SHAPE_TYPE.GROUP:
                collect(shape.shapes)
                continue
            if st == MSO_SHAPE_TYPE.PICTURE:
                stats["pics"] += 1
                saved = save_picture(shape)
                if saved:
                    pic_files.append(saved)
                continue
            if getattr(shape, "has_chart", False):
                stats["charts"] += 1
                body_lines.append("[图表对象]")
                continue
            if getattr(shape, "has_table", False):
                stats["tables"] += 1
                for row in shape.table.rows:
                    cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                    body_lines.append("| " + " | ".join(cells) + " |")
                continue
            if getattr(shape, "has_text_frame", False):
                for para in shape.text_frame.paragraphs:
                    text = (para.text or "").strip()
                    if text:
                        indent = "  " * (para.level or 0)
                        body_lines.append(f"{indent}- {text}")
                        stats["chars"] += len(text)

    collect(slide.shapes)

    if body_lines:
        lines.extend(body_lines)
    else:
        lines.append("(无文本内容 —— 可能为整页图片/SmartArt，需人工核对视觉信息)")

    lines.append(
        f"[视觉元素] 图片×{stats['pics']} 表格×{stats['tables']} 图表×{stats['charts']} · 正文字数≈{stats['chars']}"
    )
    if pic_files:
        lines.append(f"[已导出图片] {' '.join(pic_files)}（用 Read 工具核验后纳入审查结论）")
    if stats["chars"] > 400:
        lines.append(f"[审查提示] 文本量偏高（{stats['chars']}字），信息密度可能淹没论证主线")

    notes = ""
    if slide.has_notes_slide:
        notes = (slide.notes_slide.notes_text_frame.text or "").strip()
    if notes:
        stats["notes"] = len(notes)
        lines.append(f"[备注] {notes}")

    lines.append("")
    return lines, stats


def main():
    ap = argparse.ArgumentParser(description="提取 pptx 逐页内容为结构化 Markdown")
    ap.add_argument("pptx", help=".pptx 文件路径")
    ap.add_argument("-o", "--output", help="输出文件路径（默认打印到 stdout）")
    ap.add_argument("--imgs", help="导出嵌入图片的目标目录（可选；纯图形页的审查需核验图片）")
    args = ap.parse_args()

    if args.imgs:
        os.makedirs(args.imgs, exist_ok=True)

    try:
        from pptx import Presentation
    except ImportError:
        sys.stderr.write("[ERROR] 缺少依赖 python-pptx，请先在隔离环境中安装:\n"
                         "  pip install python-pptx\n")
        sys.exit(2)

    try:
        prs = Presentation(args.pptx)
    except Exception as exc:
        sys.stderr.write(f"[ERROR] 无法打开文件: {exc}\n")
        sys.exit(3)

    total = len(prs.slides)
    out = [f"# PPT 内容提取 · 共 {total} 页 · 文件: {args.pptx}", ""]

    all_stats = []
    shared_counter = [0]
    for idx, slide in enumerate(prs.slides, 1):
        section, stats = build_slide_section(
            slide, idx, total, img_dir=args.imgs, img_counter=shared_counter
        )
        out.extend(section)
        all_stats.append(stats)

    # 全局统计，辅助发现结构性问题
    no_text = sum(1 for s in all_stats if s["chars"] == 0 and s["tables"] == 0)
    no_notes = sum(1 for s in all_stats if s["notes"] == 0)
    total_chars = sum(s["chars"] for s in all_stats)
    out.append("---")
    out.append(f"[全局统计] 总正文字数≈{total_chars} · 无文本页×{no_text} · 无备注页×{no_notes}/{total}")
    if total > 0:
        avg = total_chars / total
        out.append(f"[全局统计] 平均每页≈{avg:.0f}字")
        if no_text > total * 0.3:
            out.append("[审查提示] 超过三成页面无文本——大量图形页，文本审查结论需谨慎外推")

    text = "\n".join(out)
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"[OK] 已写入 {args.output}")
    else:
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
        print(text)


if __name__ == "__main__":
    main()
