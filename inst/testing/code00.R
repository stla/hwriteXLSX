library(magrittr)
library(jsonAccess)

json <- "{\"Sheet1\":{\"A1\":{\"value\":2,\"format\":{\"numberFormat\":\"Nf2Decimal\",\"font\":{\"bold\":true}}},\"B2\":{\"value\":1000,\"format\":{\"numberFormat\":\"yyyy-mm-dd;@\"}},\"A3\":{\"value\":\"abc\",\"format\":{\"font\":{\"name\":\"Courier\",\"color\":\"FF00FF00\"}}}}}"

cat(jsonPretty(json))
jsonlite::fromJSON(json)

print(data.tree::as.Node(jsonlite::fromJSON(json)))

# A1 = list(value = 2, format = ...)
# format = list(numberFormat = "Nf2Decimal", font = ...)
# font = list(bold=TRUE, name="Courier", color = "FF00FF00") # color ARGB

createFont <- function(name=NULL, bold=NULL, color=NULL){
  if(!is.null(bold) && !is.logical(bold)){
    stop("`bold` must be logical.")
  }
  if(!is.null(color)){
    stopifnot(is.character(color))
    if(color %in% grDevices::colours()){
      RGB <- grDevices::col2rgb(colour) # faire une fonction ARGB
      hexCode <- grDevices::rgb(RGB[1,], RGB[2,], RGB[3,], 255, maxColorValue = 255)
      hexCode <- rawToChar(charToRaw(hexCode), multiple = TRUE)
      color <- paste0("FF", paste0(hexCode[2:7], collapse=""))
    }else{
      if(nchar(color) != 6L){
        stop("Invalid color")
      }
      chars <- rawToChar(charToRaw(color), multiple = TRUE)
      if(!all(chars %in% c(as.character(0:9), LETTERS[1:6]))){
        stop("Invalid color")
      }
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

createCell <- function(value=NULL, format=NULL){
  return(list(value=value, format=format))
}

A1 <- createCell(value=2,
                 format=createFormat(numberFormat = "General",
                                     font = createFont(bold=TRUE, color="yellow")))

jsonlite::toJSON(list(A1=A1), auto_unbox = TRUE, null="null")

sheet <- list(Sheet1 = list(A1=A1))
as.character(jsonlite::toJSON(sheet, auto_unbox = TRUE, null="null"))

sheets <- list(Sheet1 = list(A1=A1), Sheet2 = list(A1=A1))
as.character(jsonlite::toJSON(sheets, auto_unbox = TRUE, null="null"))


cell <- function(col, row, value, numberFormat=NULL, name=NULL, bold=NULL, color=NULL){
  cellRef <- paste0(openxlsx:::convert_to_excel_ref(col, LETTERS), row)
  setNames(list(createCell(
    value = value,
    format = createFormat(numberFormat = numberFormat,
                          font = createFont(name = name,
                                            bold = bold,
                                            color = color)))), cellRef)
}

cell(2, 3, "abc", bold=TRUE, color="red")

