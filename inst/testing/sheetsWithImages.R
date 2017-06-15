setwd("~/MyPackages/hwriteXLSX/inst/testing")

# Sys.Date() - as.Date("1899-12-30") # 42900
A1 <- cell(1, 1, 42900, numberFormat="yyyy-mm-dd;@", comment="date")
B3 <- cell(2, 3, "abc", bold=TRUE, color="red")
cells <- list(Sheet1=c(A1,B3))
json1 <- jsonlite::toJSON(cells, null="null", auto_unbox = TRUE)

image1File <- system.file("images", "Lissajous.png", package="hwriteXLSX")
image2File <- system.file("images", "power.png", package="hwriteXLSX")

image1 <- list(file=image1File, left=3, top=2, width=400, height=400)
image2 <- list(file=image2File, left=8, top=2, width=400, height=400)
images <- list(Sheet1 = list(image1), Sheet2 = list(image1, image2))
json2 <- jsonlite::toJSON(images, auto_unbox = TRUE)

json2xlsx(json1, json2, outfile="zz.xlsx", overwrite = TRUE)
# repair car mêmes filenames
