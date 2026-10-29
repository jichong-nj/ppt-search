#!/usr/bin/env python3
"""生成测试用 pptx 夹具：deck-a(3页,中文,表格,备注,图片) deck-b(2页,中英混合)。
布局：<out>/deck-a/source.pptx、<out>/deck-b/source.pptx（与 decksRoot 约定一致）。"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

out_dir = Path(sys.argv[1] if len(sys.argv) > 1 else ".")

# 一个 1x1 红色 PNG
PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d4948445200000001000000010802000000907753"
    "de0000000c4944415408d763f8cfc0000007f001ef88e8700000000049454e44ae426082"
)
png_path = out_dir / "dot.png"
png_path.parent.mkdir(parents=True, exist_ok=True)
png_path.write_bytes(PNG)


def add_slide(prs, title, body_lines, notes=None):
    layout = prs.slide_layouts[5] if title is None else prs.slide_layouts[1]
    slide = prs.slides.add_slide(layout)
    if title is not None:
        slide.shapes.title.text = title
        tf = slide.placeholders[1].text_frame
        for i, line in enumerate(body_lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
    else:
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(2))
        tf = box.text_frame
        for i, line in enumerate(body_lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
    if notes is not None:
        slide.notes_slide.notes_text_frame.text = notes
    return slide


prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

add_slide(prs, "年度总结", ["本年度营收增长 25%", "机器学习平台上线", "团队扩张至 50 人"], notes="演讲提示：先讲营收")
add_slide(prs, "产品路线图", ["Q1 完成检索引擎", "Q2 完成 MCP 接口", "Q3 上线合成能力"])
s3 = add_slide(prs, "数据概览", ["详见下表", "营收 1.2 亿"])
table = s3.shapes.add_table(2, 2, Inches(1), Inches(3.5), Inches(6), Inches(1.5)).table
table.cell(0, 0).text = "指标"
table.cell(0, 1).text = "数值"
table.cell(1, 0).text = "营收"
table.cell(1, 1).text = "1.2亿"
s3.shapes.add_picture(str(png_path), Inches(8), Inches(4))
(out_dir / "deck-a").mkdir(parents=True, exist_ok=True)
prs.save(out_dir / "deck-a" / "source.pptx")

prs2 = Presentation()
prs2.slide_width = Inches(10)
prs2.slide_height = Inches(7.5)
add_slide(prs2, "Quarterly Revenue", ["revenue grew 25 percent", "machine learning segment doubled", "old-brand needs replace"])
add_slide(prs2, "技术架构 Architecture", ["检索引擎 BM25", "存储 JSON", "old-brand 出现"])
(out_dir / "deck-b").mkdir(parents=True, exist_ok=True)
prs2.save(out_dir / "deck-b" / "source.pptx")
print("fixtures ok")
