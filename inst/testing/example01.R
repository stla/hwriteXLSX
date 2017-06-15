setwd("~/MyPackages/hwriteXLSX/inst/testing")
library(hwriteXLSX)

sheet1 <- createSheet(iris, "iris")
sheet2 <- createSheet(mtcars, "mtcars")
worksheet <- c(sheet1, sheet2)

image1File <- system.file("images", "Lissajous.png", package="hwriteXLSX")
image2File <- system.file("images", "power.png", package="hwriteXLSX")
image3File <- system.file("images", "timeseries.png", package="hwriteXLSX")
image1 <- list(file=image1File, left=6, top=2, width=400, height=400)
image2 <- list(file=image2File, left=6, top=23, width=500, height=400)
image3 <- list(file=image3File, left=13, top=2, width=500, height=400)
images <- list(iris = list(image1, image2), mtcars = list(image3))

hwriteXLSX("example01.xlsx", worksheet, images, overwrite = TRUE)
