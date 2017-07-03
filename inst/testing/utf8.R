dyn.load("C:/HaskellProjects/utf8dll/utf8dll.dll") -> dll
.C("HsStart")
input <- "{\"a\":\"?\"}"
Encoding(input) <- "UTF-8"
input <- "{\"a\":\"µ\"}"
input <- "{\"a\":\"\u00b5\"}"
input <- jsonlite::toJSON(list(a="µ"), auto_unbox = TRUE)
.C("readText", input=input, output="")
.C("readJson1", input=input, result="") # failure
.C("readJson1", input=iconv(input, "latin1", "UTF-8"), result="") # failure!!
.C("readJson2", input=input, result="") # success
.C("readJson3", input=input, result="")
dyn.unload("C:/HaskellProjects/utf8dll/utf8dll.dll")

x <- "{\"a\":\"µ\"}"
y <- jsonlite::toJSON(list(a="µ"), auto_unbox = TRUE)
y <- as.character(y)
# x and y are identical:
identical(x,y)
# but:
system(x)
system(y)

# however we get two different results:
iconv(x, "latin1", "UTF-8")
iconv(y, "latin1", "UTF-8")
# x and y have not the same encoding
Encoding(x)
Encoding(y)
# but even after setting the same encoding, the difference remains:
Encoding(x) <- "UTF-8"
iconv(x, "latin1", "UTF-8")
# another way to see this difference:
system(x)
system(y)


setwd("~/MyPackages/hwriteXLSX/inst/testing")
library(hwriteXLSX)

A1 <- cell(1, 1, 9.9, numberFormat="2Decimal")
B3 <- cell(2, 3, "µ", bold=TRUE, color="red", comment="Stéphane")
sheet <- list(Sheet1 = c(A1, B3))
# JSON string ready for json2xlsx
json <- jsonlite::toJSON(sheet, null="null", auto_unbox = TRUE)
#json <- iconv(json, "latin1", "UTF-8") maintenant dans le code
hwriteXLSX:::.json2xlsx(json, "{}", "{}", outfile="xlsx.xlsx")



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

##
hwriteXLSX:::json2xlsx2(jsonlite::toJSON(sheet, auto_unbox = TRUE, null="null"),
                        "{}", "xlsx.xlsx", overwrite = TRUE)


##
sheet <- "{\"Sheet1\":{\"A1\":{\"value\":\"A\",\"format\":{\"numberFormat\":null,\"font\":{\"name\":\"Verdana\",\"bold\":true,\"color\":null}},\"comment\":null},\"A2\":{\"value\":1,\"format\":{\"numberFormat\":null,\"font\":{\"name\":null,\"bold\":null,\"color\":null}},\"comment\":null},\"A3\":{\"value\":2,\"format\":{\"numberFormat\":null,\"font\":{\"name\":null,\"bold\":null,\"color\":null}},\"comment\":null},\"A4\":{\"value\":3,\"format\":{\"numberFormat\":null,\"font\":{\"name\":null,\"bold\":null,\"color\":null}},\"comment\":null},\"B1\":{\"value\":\"B\",\"format\":{\"numberFormat\":null,\"font\":{\"name\":\"Verdana\",\"bold\":true,\"color\":null}},\"comment\":null},\"B2\":{\"value\":null,\"format\":{\"numberFormat\":null,\"font\":{\"name\":null,\"bold\":null,\"color\":null}},\"comment\":null},\"B3\":{\"value\":\"Stéphane\",\"format\":{\"numberFormat\":null,\"font\":{\"name\":null,\"bold\":null,\"color\":null}},\"comment\":\"hello\"},\"B4\":{\"value\":1000,\"format\":{\"numberFormat\":null,\"font\":{\"name\":null,\"bold\":null,\"color\":null}},\"comment\":null},\"C1\":{\"value\":\"C\",\"format\":{\"numberFormat\":null,\"font\":{\"name\":\"Verdana\",\"bold\":true,\"color\":null}},\"comment\":null},\"C2\":{\"value\":null,\"format\":{\"numberFormat\":null,\"font\":{\"name\":null,\"bold\":null,\"color\":null}},\"comment\":\"µ\"},\"C3\":{\"value\":\"b\",\"format\":{\"numberFormat\":null,\"font\":{\"name\":null,\"bold\":null,\"color\":null}},\"comment\":null}}}"

hwriteXLSX:::json2xlsx2(sheet, "{}", "xlsx.xlsx", overwrite = TRUE)
