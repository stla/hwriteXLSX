setwd("~/MyPackages/hwriteXLSX/inst/testing")
library(hwriteXLSX)

colorMatrix <- matrix(grDevices::colours(), ncol=9)
f <- function(i, j){
  cell(j, i, value = colorMatrix[i,j], fill = colorMatrix[i,j])
}
sheet <- list(Colors = do.call(Vectorize(f), expand.grid(i=1:73, j=1:9)))

hwriteXLSX("fill.xlsx", sheet, overwrite = TRUE)

