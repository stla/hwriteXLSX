setwd("~/MyPackages/hwriteXLSX/inst/testing")
library(hwriteXLSX)

dat <-
  list(
    A = list(1, 2, 3),
    B = list(NULL, "a", 1000),
    C = list(NA, "b")
  )
attr(dat$C[[2]], "colspan") <- 2
attr(dat$C[[2]], "rowspan") <- 2

sheet <- createSheet(dat, "Sheet1")
hwriteXLSX("span.xlsx", sheet, overwrite = TRUE)

dat <-
  list(
    A = list("long cell", 2, 3),
    B = list(NULL, "a", 1000),
    C = list(NA, "b")
  )
attr(dat$A[[1]], "colspan") <- 3
attr(dat$A[[1]], "fill") <- "pink"
sheet <- createSheet(dat, "Sheet1")
hwriteXLSX("span2.xlsx", sheet, overwrite = TRUE)
