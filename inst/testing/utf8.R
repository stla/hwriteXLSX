dyn.load("C:/HaskellProjects/utf8dll/utf8dll.dll")
.C("HsStart")
input <- jsonlite::toJSON("{\"a\":\"µ\"}")
.C("readText", input=input, output="")
dyn.unload("C:/HaskellProjects/utf8dll/utf8dll.dll")


setwd("~/MyPackages/hwriteXLSX/inst/testing")
library(hwriteXLSX)

A1 <- cell(1, 1, 9.9, numberFormat="2Decimal")
B3 <- cell(2, 3, "µ", bold=TRUE, color="red", comment="Stéphane")
sheet <- list(Sheet1 = c(A1, B3))
# JSON string ready for json2xlsx
json <- jsonlite::toJSON(sheet, null="null", auto_unbox = TRUE)
#json <- iconv(json, "latin1", "UTF-8") maintenant dans le code
json2xlsx(json, outfile="xlsx.xlsx", overwrite = TRUE)



dat <-
  list(
    A = list(1, 2, 3),
    B = list(NULL, "Stéphane", 1000),
    C = list(NA, "b")
  )
attr(dat$B[[2]], "comment") <- "hello"
attr(dat$C[[1]], "comment") <- "µ"

sheet <- createSheet(dat, "Sheet1")
hwriteXLSX("utf8.xlsx", sheet, overwrite = TRUE)
