#!/usr/bin/env python3
# 生成富样本 rich.pptx：文本/项目符号/颜色、柱状图、折线图、饼图、环形图、表格、图片、形状、组合、渐变
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "fixtures", "rich.pptx")
DOT = os.path.join(HERE, "fixtures", "dot.png")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# --- slide 1: 标题 + 正文（多 run 样式） ---
s = prs.slides.add_slide(prs.slide_layouts[1])
s.shapes.title.text = "渲染模型测试"
tf = s.placeholders[1].text_frame
tf.text = "第一条要点"
p = tf.add_paragraph()
r = p.add_run(); r.text = "第二条：粗体"; r.font.bold = True; r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
r2 = p.add_run(); r2.text = " + 斜体"; r2.font.italic = True
p3 = tf.add_paragraph(); p3.text = "第三条右对齐"; p3.alignment = PP_ALIGN.RIGHT

# --- slide 2: 柱状图 + 折线图 ---
s2 = prs.slides.add_slide(prs.slide_layouts[5])
cd = CategoryChartData(); cd.categories = ["一月", "二月", "三月"]
cd.add_series("销量", (1.2, 3.4, 2.8)); cd.add_series("退货", (0.2, 0.4, 0.8))
s2.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.4), Inches(5.8), Inches(4.5), cd)
cd2 = CategoryChartData(); cd2.categories = ["Q1", "Q2", "Q3", "Q4"]
cd2.add_series("营收", (10, 22, 15, 30))
s2.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(6.8), Inches(1.4), Inches(6), Inches(4.5), cd2)

# --- slide 3: 饼图 + 环形图 + 表格 ---
s3 = prs.slides.add_slide(prs.slide_layouts[5])
cd3 = CategoryChartData(); cd3.categories = ["甲", "乙", "丙"]
cd3.add_series("占比", (45, 30, 25))
s3.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(0.5), Inches(1.4), Inches(4), Inches(3.2), cd3)
cd4 = CategoryChartData(); cd4.categories = ["东", "南", "西", "北"]
cd4.add_series("区域", (12, 8, 20, 5))
s3.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(4.9), Inches(1.4), Inches(4), Inches(3.2), cd4)
tbl = s3.shapes.add_table(3, 3, Inches(9.2), Inches(1.4), Inches(3.6), Inches(1.6)).table
tbl.cell(0, 0).text = "地区"; tbl.cell(0, 1).text = "数值"; tbl.cell(0, 2).text = "备注"
tbl.cell(1, 0).text = "华东"; tbl.cell(1, 1).text = "123"
tbl.cell(2, 0).text = "华南"; tbl.cell(2, 1).text = "456"

# --- slide 4: 图片 + 形状 + 组合 + 渐变 ---
s4 = prs.slides.add_slide(prs.slide_layouts[5])
if os.path.exists(DOT):
    s4.shapes.add_picture(DOT, Inches(0.5), Inches(1.4), Inches(2), Inches(2))
rr = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3), Inches(1.4), Inches(3), Inches(1.2))
rr.fill.solid(); rr.fill.fore_color.rgb = RGBColor(0x44, 0x72, 0xC4)
rr.text_frame.text = "圆角矩形"
ell = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3), Inches(3), Inches(2), Inches(2))
ell.fill.gradient()
ell.line.color.rgb = RGBColor(0xFF, 0x00, 0x00); ell.line.width = Pt(2.5)
ar = s4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(6.5), Inches(3), Inches(2.5), Inches(1))
ar.fill.solid(); ar.fill.fore_color.rgb = RGBColor(0x70, 0xAD, 0x47)
grp = s4.shapes.add_group_shape()
g1 = grp.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(1), Inches(1))
g1.fill.solid(); g1.fill.fore_color.rgb = RGBColor(0xFF, 0xC0, 0x00)
g2 = grp.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(1.2), Inches(0), Inches(1), Inches(1))
grp.left, grp.top, grp.width, grp.height = Inches(8), Inches(1.4), Inches(4), Inches(2)
tb = s4.shapes.add_textbox(Inches(6.5), Inches(4.5), Inches(5), Inches(1))
tb.text_frame.text = "普通文本框 中文 English 混排换行测试，用于验证 measureText 换行逻辑的正确性。"

prs.save(OUT)
print("saved", OUT)
