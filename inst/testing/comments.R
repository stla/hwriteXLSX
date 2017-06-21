setwd("~/MyPackages/hwriteXLSX/inst/testing")

library(hwriteXLSX)

A1 <- cell(1, 1, 9.9, numberFormat="2Decimal")
B3 <- cell(2, 3, "µµ", bold=TRUE, color="red",
           comment="greek", commentAuthor = "Stéphane")
sheet <- list(Sheet1 = c(A1, B3))
# JSON string ready for json2xlsx
json <- jsonlite::toJSON(sheet, null="null", auto_unbox = TRUE)
json2xlsx(json, outfile="comments.xlsx", overwrite = TRUE)
