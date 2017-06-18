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
#' A1 <- cell(1, 1, 9.9, numberFormat="2Decimal")
#' B3 <- cell(2, 3, "abc", bold=TRUE, color="red", comment="this cell is red")
#' sheet <- list(Sheet1 = c(A1, B3))
#' # JSON string ready for json2xlsx
#' json <- jsonlite::toJSON(sheet, null="null", auto_unbox = TRUE)
#' \donttest{json2xlsx(json, outfile="xlsx.xlsx")}
#' # predefined number formats
#' numericFormats <-
#'   c(
#'     "General",
#'     "Zero",
#'     "2Decimal",
#'     "Max3Decimal",
#'     "ThousandSeparator2Decimal",
#'     "Percent",
#'     "Percent2Decimal",
#'     "Exponent2Decimal",
#'     "SingleSpacedFraction",
#'     "DoubleSpacedFraction",
#'     "ThousandsNegativeParens",
#'     "ThousandsNegativeRed",
#'     "Thousands2DecimalNegativeParens",
#'     "Tousands2DecimalNEgativeRed",
#'     "Exponent1Decimal",
#'     "TextPlaceHolder"
#'   )
#' dateFormats <-
#'   c(
#'     "MmSs",
#'     "OptHMmSs",
#'     "MmSs1Decimal",
#'     "MmDdYy",
#'     "DMmmYy",
#'     "DMmm",
#'     "MmmYy",
#'     "HMm12Hr",
#'     "HMmSs12Hr",
#'     "HMm",
#'     "HMmSs",
#'     "MdyHMm"
#'   )
#' A <- lapply(seq_along(numericFormats),
#'             function(i) cell(1, i, value=numericFormats[i]))
#' B <- lapply(seq_along(numericFormats),
#'             function(i) cell(2, i, value=-9999.1234,
#'                              numberFormat=numericFormats[i]))
#' sheet1 <- list(numericFormats = do.call(c, c(A,B)))
#' A <- lapply(seq_along(dateFormats),
#'             function(i) cell(1, i, value=dateFormats[i]))
#' B <- lapply(seq_along(dateFormats),
#'             function(i) cell(2, i, value=10000,
#'                              numberFormat=dateFormats[i]))
#' sheet2 <- list(dateFormats = do.call(c, c(A,B)))
#' json <- jsonlite::toJSON(c(sheet1, sheet2), null="null", auto_unbox = TRUE)
#' \donttest{json2xlsx(json, outfile="numberFormats.xlsx")}
cell <- function(col, row, value, comment=NULL, numberFormat=NULL, fontname=NULL, bold=NULL, color=NULL){
  #cellRef <- paste0(openxlsx:::convert_to_excel_ref(col, LETTERS), row)
  # ou bien paste0(cellranger::num_to_letter(col), row)
  cellRef <- paste0(int_to_letter(col), row)
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
#' @seealso \code{\link{hwriteXLSX}}
#'
#' @examples
#' # Create a sheet from a dataframe
#' sheet <- createSheet(mtcars[1:2, 1:2], "Sheet1")
#' # write to xlsx.xlsx:
#' \donttest{hwriteXLSX(file="xlsx.xlsx", worksheet=sheet)}
#' # Create a sheet from a list of lists
#' dat <-
#'   list(
#'     A = list(1, 2, 3),
#'     B = list(NULL, "a", 1000),
#'     C = list(NA, "b")
#'   )
#' attr(dat$B[[2]], "color") <- "red"
#' attr(dat$B[[2]], "comment") <- "this cell is red"
#' attr(dat$C[[1]], "comment") <- "this cell is empty"
#' sheet <- createSheet(dat, "Sheet1")
createSheet <- function(dat, sheetname){
  D <- dict()
  for(j in seq_along(dat)){
    D[[c(1,j)]] <- cell(j, 1, names(dat)[j], fontname="Verdana", bold=TRUE)
    column <- dat[[j]]
    for(i in seq_along(column)){
        value <- column[[i]]
        if(is_date(value)){
          D[[c(i+1,j)]] <- cellDate(j, i+1, value,
                                    comment = attr(value, "comment"),
                                    color = attr(value, "color"),
                                    fontname = attr(value, "fontname"),
                                    bold = attr(value, "bold"))
        }else{
          D[[c(i+1,j)]] <- cell(j, i+1, value,
                                comment = attr(value, "comment"),
                                color = attr(value, "color"),
                                fontname = attr(value, "fontname"),
                                bold = attr(value, "bold"),
                                numberFormat = attr(value, "numberFormat"))
        }
    }
  }
  return(setNames(list(do.call(c, D$values())), sheetname))
}


#' @title Write XLSX file
#' @description Write a XLSX file from a series of sheets,
#' allowing inclusion of images.
#'
#' @param worksheet a sheet such as one created by \code{\link{createSheet}},
#' or a series of such sheets concatenated by \code{c}.
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
#' \donttest{hwriteXLSX("xlsx.xlsx", worksheet)}
#' # include some images
#' image1File <- system.file("images", "Lissajous.png", package="hwriteXLSX")
#' image2File <- system.file("images", "power.png", package="hwriteXLSX")
#' image3File <- system.file("images", "timeseries.png", package="hwriteXLSX")
#' image1 <- list(file=image1File, left=6, top=2, width=400, height=400)
#' image2 <- list(file=image2File, left=6, top=23, width=500, height=400)
#' image3 <- list(file=image3File, left=13, top=2, width=500, height=400)
#' images <- list(iris = list(image1, image2), mtcars = list(image3))
#' \donttest{hwriteXLSX("xlsx.xlsx", worksheet, images, overwrite = TRUE)}
#' # colors example
#' colorMatrix <- matrix(grDevices::colours(), ncol=9)
#' f <- function(i, j){
#'   cell(j, i, value = colorMatrix[i,j], color = colorMatrix[i,j], bold=TRUE)
#' }
#' sheet <- list(Colors = do.call(Vectorize(f), expand.grid(i=1:73, j=1:9)))
#' \donttest{hwriteXLSX("colors.xlsx", sheet)}
#' # Write a XLSX file from a list of lists
#' dat <-
#'   list(
#'     A = list(1, 2, 3),
#'     B = list(NULL, "a", 1000),
#'     C = list(NA, "b")
#'   )
#' attr(dat$B[[2]], "color") <- "red"
#' attr(dat$B[[2]], "comment") <- "this cell is red"
#' attr(dat$C[[1]], "comment") <- "this cell is empty"
#' sheet <- createSheet(dat, "Sheet1")
#' \donttest{hwriteXLSX("xlsx.xlsx", sheet, overwrite = TRUE)}
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
  json2xlsx(jsonCells, jsonImages, outfile=file, overwrite=TRUE)
}
