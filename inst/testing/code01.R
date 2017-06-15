library(magrittr)
library(jsonAccess)

json <- "{\"Sheet1\":{\"A1\":{\"value\":2,\"format\":{\"numberFormat\":\"Nf2Decimal\",\"font\":{\"bold\":true}}},\"B2\":{\"value\":1000,\"format\":{\"numberFormat\":\"yyyy-mm-dd;@\"}},\"A3\":{\"value\":\"abc\",\"format\":{\"font\":{\"name\":\"Courier\",\"color\":\"FF00FF00\"}}}}}"
cat(jsonPretty(json))
jsonlite::fromJSON(json)


# A1 = list(value = 2, format = ...)
# format = list(numberFormat = "Nf2Decimal", font = ...)
# font = list(bold=TRUE, name="Courier", color = "FF00FF00") # color ARGB

createFont <- function(name=NULL, bold=NULL, color=NULL){
  if(!is.null(bold) && !is.logical(bold)){
    stop("`bold` must be logical.")
  }
  if(!is.null(color)){
    if(! color %in% grDevices::colours()){
      stop("Invalid color name.")
    }
  }
  if(!is.null(name) && !is.character(name)){
    stop("`name` must be a font name")
  }
  list(name=name, bold=bold, color=color)
}

createFormat <- function(numberFormat=NULL, font=NULL){
  list(numberFormat=numberFormat, font=font)
}

createCell <- function(value=NULL, format=NULL, comment=NULL){
  return(list(value=value, format=format, comment=comment))
}

A1 <- createCell(value=2,
                 format=createFormat(numberFormat = "General",
                                     font = createFont(bold=TRUE, color="yellow")))

jsonlite::toJSON(list(A1=A1), auto_unbox = TRUE, null="null")

sheet <- list(Sheet1 = list(A1=A1))
as.character(jsonlite::toJSON(sheet, auto_unbox = TRUE, null="null"))

sheets <- list(Sheet1 = list(A1=A1), Sheet2 = list(A1=A1))
as.character(jsonlite::toJSON(sheets, auto_unbox = TRUE, null="null"))


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

cell(2, 3, "abc", bold=TRUE, color="red")

colorMatrix <- matrix(grDevices::colours(), ncol=9)
f <- function(i, j){
  cell(i, j, value = colorMatrix[i,j], color = colorMatrix[i,j])
}
sheet <- list(Colors = do.call(Vectorize(f), expand.grid(i=1:73,j=1:9)))

