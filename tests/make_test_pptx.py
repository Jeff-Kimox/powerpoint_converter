import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from PIL import Image, ImageDraw

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

# 1 title slide
s = prs.slides.add_slide(prs.slide_layouts[0])
s.shapes.title.text = "Quarterly Review"
s.placeholders[1].text = "Finance Team - Q3 2026"

# 2 bullets with nesting
s = prs.slides.add_slide(prs.slide_layouts[1])
s.shapes.title.text = "Key Points"
tf = s.placeholders[1].text_frame
tf.text = "Revenue grew 12%"
for t, lvl in [("Driven by new clients", 1), ("Nairobi region strongest", 2), ("Costs flat", 0)]:
    p = tf.add_paragraph(); p.text = t; p.level = lvl
p = tf.add_paragraph(); r = p.add_run(); r.text = "See the "; r2 = p.add_run(); r2.text = "full report"; r2.hyperlink.address = "https://example.com"; r2.font.bold = True

# 3 picture + text
s = prs.slides.add_slide(prs.slide_layouts[5])
s.shapes.title.text = "Product Photo"
img = Image.new("RGB", (800, 500), (30, 120, 200)); ImageDraw.Draw(img).rectangle((50,50,750,450), outline="white", width=8)
buf = io.BytesIO(); img.save(buf, "PNG"); buf.seek(0)
pic = s.shapes.add_picture(buf, Inches(1), Inches(1.8), width=Inches(7))
pic.crop_left = 0.1
tb = s.shapes.add_textbox(Inches(9), Inches(2), Inches(3.5), Inches(2)); tb.text_frame.text = "Caption text beside the image"

# 4 table
s = prs.slides.add_slide(prs.slide_layouts[5])
s.shapes.title.text = "Budget Table"
gf = s.shapes.add_table(3, 3, Inches(1), Inches(2), Inches(10), Inches(2))
t = gf.table
for r, row in enumerate([["Item","Q2","Q3"],["Rent","100","110"],["Salaries","500","520"]]):
    for c, v in enumerate(row): t.cell(r,c).text = v
t.cell(1,0).merge(t.cell(2,0))

# 5 chart
s = prs.slides.add_slide(prs.slide_layouts[5])
s.shapes.title.text = "Sales Chart"
cd = CategoryChartData(); cd.categories = ["Jan","Feb","Mar"]; cd.add_series("2025",(10,20,15)); cd.add_series("2026",(12,25,22))
s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(2), Inches(8), Inches(4), cd)
s.notes_slide.notes_text_frame.text = "Mention that March beat the target.\nAsk for questions."

# 6 text box numbered list, group, no title
s = prs.slides.add_slide(prs.slide_layouts[6])
g = s.shapes.add_group_shape()
b = g.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1)); b.text_frame.text = "Text inside a group"
tb = s.shapes.add_textbox(Inches(1), Inches(3), Inches(6), Inches(2))
tb.text_frame.text = "Plain paragraph one"
tb.text_frame.add_paragraph().text = "Plain paragraph two"

# 7 hidden
s = prs.slides.add_slide(prs.slide_layouts[1]); s.shapes.title.text = "Hidden slide"; s._element.set("show","0")
prs.save("test_deck.pptx")
