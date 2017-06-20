context("Dates")

test_that("Dates with cellDate", {
  A1 <- cellDate(1, 1, "06-14-2017")
  A2 <- cellDate(1, 2, "2017/06/14")
  sheet <- list(Sheet1 = c(A1, A2))
  json <- jsonlite::toJSON(sheet, null="null", auto_unbox = TRUE)
  tmpXlsx <- tempfile(fileext = ".xlsx")
  expect_silent(json2xlsx(json, outfile=tmpXlsx))
  dat <- as.data.frame(readxl::read_xlsx(tmpXlsx, sheet=1, col_names=FALSE))
  dates <- as.character(dat[,1])
  expect_identical(dates, rep("2017-06-14", 2))
})

test_that("Dates with createSheet", {
  dat <- data.frame(Date1 = as.Date("2017-06-14"),
                    Date2 = as.POSIXct("2017-06-14"))
  sheet <- createSheet(dat, "Sheet1")
  tmpXlsx <- tempfile(fileext = ".xlsx")
  expect_silent(hwriteXLSX(tmpXlsx, sheet))
  dat2 <- as.data.frame(readxl::read_xlsx(tmpXlsx, sheet=1))
  expect_identical(as.character(dat2$Date1[1]), "2017-06-14")
  expect_identical(as.character(dat2$Date2[1]), "2017-06-14")
})
