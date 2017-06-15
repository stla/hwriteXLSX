#' @title General cell
#' @description Creates a general cell.
#'
#' @param col integer, the column index
#' @param row integer, the row index
#' @param value cell value
#' @param comment cell comment
#' @param numberFormat number format; see Details
#' @param fontname font name; see Details
#' @param bold set \code{TRUE} for bold font
#' @param color color name; see Details
#'
#' @return A named list representing a cell.
#' @export
#'
#' @details xxx
#'
#' @seealso \code{\link{cellDate}} to create a cell for a date
#'
#' @examples
#' A1 <- cell(1, 1, 9.9, numberFormat="Nf2Decimal")
#' B3 <- cell(2, 3, "abc", bold=TRUE, color="red", comment="this cell is red")
#' sheet <- list(Sheet1 = c(A1, B3))
#' # JSON string ready for json2xlsx
#' jsonlite::toJSON(sheet, null="null", auto_unbox = TRUE)
cell <- function(col, row, value, comment=NULL, numberFormat=NULL, fontname=NULL, bold=NULL, color=NULL){
  cellRef <- paste0(openxlsx:::convert_to_excel_ref(col, LETTERS), row)
  # ou bien paste0(cellranger::num_to_letter(col), row)
  setNames(list(createCell(
    value = value,
    comment = comment,
    format = createFormat(numberFormat = numberFormat,
                          font = createFont(name = fontname,
                                            bold = bold,
                                            color = color)))), cellRef)
}

#' @title Date cell
#' @description Creates a cell for a date.
#'
#' @param col integer, the column index
#' @param row integer, the row index
#' @param date an input which will be parsed to a date through the
#' \code{\link[anytime]{anydate}} function
#' @param comment cell comment
#' @param fontname font name; see \code{\link{cell}}
#' @param bold set \code{TRUE} for bold font
#' @param color color name; see \code{\link{cell}}
#' @param ... arguments passed to \code{\link[anytime]{anydate}}
#'
#' @return A named list representing a cell.
#' @export
#' @importFrom anytime anydate
#'
#' @details A date in Excel is stored as an integer: the number of days between
#' this date and a certain origin. The \code{cellDate} function calculates this
#' integer and then passes it to the function \code{\link{cell}} with a date format.
#'
#' @examples
#' A1 <- cellDate(1, 1, "2017-06-14")
#' unlist(A1)
#' A1 <- cellDate(1, 1, as.Date(42900, origin="1899-12-30"))
#' unlist(A1)
#' A1 <- cellDate(1, 1, "2017/06/14")
#' unlist(A1)
#' A1 <- cellDate(1, 1, "20170614")
#' unlist(A1)
#' A1 <- cellDate(1, 1, "06-14-2017")
#' unlist(A1)
#' A1 <- cellDate(1, 1, "2017-06-14 04:05:06")
#' unlist(A1)
#' A1 <- cellDate(1, 1, "2017-06-14 04:05:06", tz="America/Los_Angeles")
#' unlist(A1)
#' A1 <- cellDate(1, 1, Sys.Date())
#' unlist(A1)
#' A1 <- cellDate(1, 1, "x")
#' unlist(A1)
cellDate <- function(col, row, date, comment=NULL, fontname=NULL, bold=NULL,
                     color=NULL, ...){
  if(is.na(date) || is.null(date)){
    dateValue <- NULL
  }else{
    dateParsed <- anytime::anydate(date, ...)
    if(is.na(dateParsed)){
      warning(sprintf("Value `%s` cannnot be parsed to a date.", date))
      dateValue <- NULL
    }else{
      dateValue <- as.integer(dateParsed - as.Date("1899-12-30"))
    }
  }
  cellRef <- paste0(openxlsx:::convert_to_excel_ref(col, LETTERS), row)
  setNames(list(createCell(
    value = dateValue,
    comment = comment,
    format = createFormat(numberFormat = "yyyy-mm-dd;@",
                          font = createFont(name = fontname,
                                            bold = bold,
                                            color = color)))), cellRef)
}


# createSheetCells <- function(dat, sheetname, guessDates=FALSE){
#   # et les comments ?..
#   # ou bien, à la place de dat:
#   #   fonction (j,i) -> cell(j, i, ...) # genre Haskell FormattedCellMap
#   #   problème des keys... NULL cell ? pas de pb:value=NULL
#   # comment avoir toutes les keys ?.. excel B2 au lieu de (2,2)
# }


#' Title
#'
#' @param worksheet xxxx
#' @param file name of the xlsx file to be written
#' @param guessDates whether to guess dates
#'
#' @return No returned value.
#' @export
#'
#' @examples
#' xxx
hwriteXLSX <- function(worksheet_cells, worksheet_images, file, guessDates=FALSE){
  # je n'autorise pas les flat list ?..
  # ou si, avec une fonction createWorksheet, as.worksheet... ?
  # aussi createSheet, addSheet ?..
  if("data.frame")
  invisible()
}
