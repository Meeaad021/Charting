import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches
from pptx.dml.color import RGBColor

# Input/Output files
excel_file = "data.xlsx"
output_ppt = "excel_barcharts.pptx"

#Specify color
BAR_COLOUR = RGBColor(0, 112, 192)


# Load all sheets as dict of DataFrames
sheets = pd.read_excel(excel_file, sheet_name=None)

# Start new PowerPoint
prs = Presentation()

for sheet_name, df in sheets.items():
    # Skip if sheet is empty or missing columns
    if df.empty or len(df.columns) < 2:
        continue

    # Clean data
    df = df.iloc[:, :2].dropna()
    df = df[~df.iloc[:, 0].astype(str).str.startswith("Base")]
    df.iloc[:, 1] = pd.to_numeric(df.iloc[:, 1], errors="coerce")
    df = df.dropna()  

    if df.empty:
        continue

    # Add a slide
    slide_layout = prs.slide_layouts[5]  # Title Only
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = f"Bar Chart - {sheet_name}"

    # Prepare chart data (col1 = categories, col2 = totals)
    chart_data = CategoryChartData()
    chart_data.categories = df.iloc[:, 0].astype(str)  # ensure text labels
    chart_data.add_series("Total", df.iloc[:, 1].astype(float))  # ensure numeric

    # Add a clustered bar chart
    x, y, cx, cy = Inches(1), Inches(2), Inches(6), Inches(4)
    chart_shapes = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED, x, y, cx, cy, chart_data
    )
    chart = chart_shapes.chart

    #Remove gridlines
    value_axis = chart.value_axis
    category_axis = chart.category_axis
    value_axis.has_major_gridlines = False
    value_axis.has_minor_gridlines = False
    category_axis.has_major_gridlines = False
    category_axis.has_minor_gridlines = False

    # Set bar color
    for series in chart.series:
        for point in series.points:
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = BAR_COLOUR

# Save presentation
prs.save(output_ppt)
print(f"✅ Presentation created: {output_ppt}")
