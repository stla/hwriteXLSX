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


#' @title Create a sheet
#' @description Create a sheet from an appropriate data input.
#'
#' @param dat appropriate data input (see examples and vignette)
#' @param sheetname name of the sheet
#'
#' @return A named list containing one element: the list of cells
#' @export
#' @import dict
#'
#' @examples
#' sheet <- createSheet(mtcars[1:2, 1:2], "Sheet1")
#' # write to xlsx.xlsx:
#' \donttest{
#' hwriteXLSX(file="xlsx.xlsx", worksheet=sheet)}
createSheet <- function(dat, sheetname){
  D <- dict()
  for(j in seq_along(dat)){
    D[[c(1,j)]] <- cell(j, 1, names(dat)[j], fontname="Verdana", bold=TRUE)
    column <- dat[[j]]
    for(i in seq_along(column)){
        value <- column[[i]]
        if(is_date(value)){
          D[[c(i+1,j)]] <- cellDate(j, i+1, value, comment=attr(value, "comment"))
        }else{
          D[[c(i+1,j)]] <- cell(j, i+1, value, comment=attr(value, "comment"))
        }
    }
  }
  return(setNames(list(do.call(c, D$values())), sheetname))
}


#' @title Write XLSX file
#' @description Write a XLSX file from a series of sheets, allowing inclusion of images.
#'
#' @param worksheet a sheet created by \code{\link{createSheet}}, or a series of
#' such sheets concatenated by \code{c}.
#' @param images a named list defining the images to be included in the worksheet;
#' see vignette and examples
#' @param file name of the xlsx file to be written
#' @param overwrite logical, whether to overwrite the output file if it already exists
#'
#' @return No returned value.
#' @export
#' @importFrom jsonlite toJSON
#'
#' @examples
#' sheet1 <- createSheet(iris, "iris")
#' sheet2 <- createSheet(mtcars, "mtcars")
#' worksheet <- c(sheet1, sheet2)
#' \donttest{
#' hwriteXLSX("xlsx.xlsx", worksheet)}
hwriteXLSX <- function(file, worksheet, images=NULL, overwrite=FALSE){
  if(!overwrite && file.exists(file)){
    stop(sprintf("File `%s` already exists.", file))
  }
  if(is.null(worksheet)){
    jsonCells <- "{}"
  }else{
    jsonCells <- jsonlite::toJSON(worksheet, null="null", auto_unbox = TRUE)
  }
  if(is.null(images)){
    jsonImages <- "{}"
  }else{
    jsonImages <- jsonlite::toJSON(images, null="null", auto_unbox = TRUE)
  }
  json2xlsx(jsonCells, jsonImages, outfile=file, overwrite=FALSE)
}
