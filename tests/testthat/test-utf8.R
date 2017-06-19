context("UTF-8")

test_that("utf8 - json2xlsx", {
  # utf8 in cell value
  sheet <- "{\"Sheet1\":{\"A1\":{\"value\":\"µ\"}}}"
  tmpXlsx <- tempfile(fileext = ".xlsx")
  expect_silent(json2xlsx(sheet, outfile=tmpXlsx))
  dat <- readxl::read_xlsx(tmpXlsx, sheet = 1, col_names = FALSE)
  expect_equal(as.data.frame(dat)[1,1], "µ")
  # utf8 in cell comment
  sheet <- "{\"Sheet1\":{\"A1\":{\"value\":\"µ\",\"comment\":\"µ\"}}}"
  tmpXlsx <- tempfile(fileext = ".xlsx")
  expect_silent(json2xlsx(sheet, outfile=tmpXlsx))
  txl <- suppressWarnings(tidyxl::tidy_xlsx(tmpXlsx, sheets=1))
  comment <- txl$data$Sheet1$comment
  expect_equal(comment, "µ")
})

test_that("utf8 - hwriteXLSX", {
  dat0 <- data.frame(A=1:2, B=c("Stéphane", "µ"), stringsAsFactors=FALSE)
  sheet <- createSheet(dat0, "Sheet1")
  tmpXlsx <- tempfile(fileext = ".xlsx")
  expect_silent(hwriteXLSX(tmpXlsx, sheet))
  dat <- readxl::read_xlsx(tmpXlsx, sheet = 1, col_names = TRUE)
  expect_equal(dat0, as.data.frame(dat))
})
