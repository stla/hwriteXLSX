setwd("~/MyPackages/hwriteXLSX/inst/testing")
library(hwriteXLSX)

dat <-
  list(
    A = list(1, 2, 3),
    B = list(NULL, "a", 1000),
    C = list(NA, "b")
  )
# does not work: attr(dat$B[[1]], "comment") <- "empty cell"
attr(dat$B[[2]], "comment") <- "hello"
attr(dat$C[[1]], "comment") <- "empty cell"

sheet <- createSheet(dat, "Sheet1")
hwriteXLSX("example02.xlsx", sheet, overwrite = TRUE)
