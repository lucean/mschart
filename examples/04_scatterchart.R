library(xml2)
library(writexl)
library(stats)
library(data.table)
library(grDevices)
library(cellranger)
library(htmltools)
library(devtools)
library(officer)

load_all()


# example chart_01 -------
chart_01 <- ms_scatterchart(
  data = mtcars, x = "disp",
  y = "drat"
)
chart_01 <- chart_settings(chart_01, scatterstyle = "marker")


# example chart_02 -------
chart_02 <- ms_scatterchart(
  data = iris, x = "Sepal.Length",
  y = "Petal.Length", group = "Species"
)
chart_02 <- chart_settings(chart_02, scatterstyle = "marker")
chart_02 <- chart_ax_x(chart_02)
chart_02 <- chart_ax_y(chart_02, major_unit = 1.0, minor_unit = 0.5)

doc <- read_pptx()
doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
doc <- ph_with(doc, value = chart_02, location = ph_location_fullsize())

print(doc, target = "example.pptx")