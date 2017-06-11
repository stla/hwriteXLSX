#' Create a cell
#'
#' @param col integer, the column index
#' @param row integer, the row index
#' @param value cell value
#' @param comment cell comment
#' @param numberFormat number format; see Details
#' @param fontname font name; see Details
#' @param bold set \code{TRUE} for bold font
#' @param color color name
#'
#' @return A named list representing a cell.
#' @export
#'
#' @examples
#' B3 <- cell(2, 3, "abc", bold=TRUE, color="red")
#' sheet <- setNames(list(B3), "Sheet1")
#' # JSON string ready for json2xlsx
#' jsonlite::toJSON(sheet, null="null", auto_unbox = TRUE)
cell <- function(col, row, value, comment=NULL, numberFormat=NULL, fontname=NULL, bold=NULL, color=NULL){
  cellRef <- paste0(openxlsx:::convert_to_excel_ref(col, LETTERS), row)
  setNames(list(createCell(
    value = value,
    comment = comment,
    format = createFormat(numberFormat = numberFormat,
                          font = createFont(name = fontname,
                                            bold = bold,
                                            color = color)))), cellRef)
}
