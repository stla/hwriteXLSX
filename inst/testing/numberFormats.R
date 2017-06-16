setwd("~/MyPackages/hwriteXLSX/inst/testing")
library(hwriteXLSX)

numericFormats <-
  c(
    "General",
    "Zero",
    "2Decimal",
    "Max3Decimal",
    "ThousandSeparator2Decimal",
    "Percent",
    "Percent2Decimal",
    "Exponent2Decimal",
    "SingleSpacedFraction",
    "DoubleSpacedFraction",
    "ThousandsNegativeParens",
    "ThousandsNegativeRed",
    "Thousands2DecimalNegativeParens",
    "Tousands2DecimalNEgativeRed",
    "Exponent1Decimal",
    "TextPlaceHolder"
  )
dateFormats <-
  c(
    "MmSs",
    "OptHMmSs",
    "MmSs1Decimal",
    "MmDdYy",
    "DMmmYy",
    "DMmm",
    "MmmYy",
    "HMm12Hr",
    "HMmSs12Hr",
    "HMm",
    "HMmSs",
    "MdyHMm"
  )


# data ImpliedNumberFormat =
#   NfGeneral                         -- ^> 0 General
# | NfZero                            -- ^> 1 0
# | Nf2Decimal                        -- ^> 2 0.00
# | NfMax3Decimal                     -- ^> 3 #,##0
# | NfThousandSeparator2Decimal       -- ^> 4 #,##0.00
# | NfPercent                         -- ^> 9 0%
# | NfPercent2Decimal                 -- ^> 10 0.00%
# | NfExponent2Decimal                -- ^> 11 0.00E+00
# | NfSingleSpacedFraction            -- ^> 12 # ?/?
# | NfDoubleSpacedFraction            -- ^> 13 # ??/??
# | NfMmDdYy                          -- ^> 14 mm-dd-yy
# | NfDMmmYy                          -- ^> 15 d-mmm-yy
# | NfDMmm                            -- ^> 16 d-mmm
# | NfMmmYy                           -- ^> 17 mmm-yy
# | NfHMm12Hr                         -- ^> 18 h:mm AM/PM
# | NfHMmSs12Hr                       -- ^> 19 h:mm:ss AM/PM
# | NfHMm                             -- ^> 20 h:mm
# | NfHMmSs                           -- ^> 21 h:mm:ss
# | NfMdyHMm                          -- ^> 22 m/d/yy h:mm
# | NfThousandsNegativeParens         -- ^> 37 #,##0 ;(#,##0)
# | NfThousandsNegativeRed            -- ^> 38 #,##0 ;[Red](#,##0)
# | NfThousands2DecimalNegativeParens -- ^> 39 #,##0.00;(#,##0.00)
# | NfTousands2DecimalNEgativeRed     -- ^> 40 #,##0.00;[Red](#,##0.00)
# | NfMmSs                            -- ^> 45 mm:ss
# | NfOptHMmSs                        -- ^> 46 [h]:mm:ss
# | NfMmSs1Decimal                    -- ^> 47 mmss.0
# | NfExponent1Decimal                -- ^> 48 ##0.0E+0
# | NfTextPlaceHolder                 -- ^> 49 @
#   | NfOtherImplied Int                -- ^ other (non local-neutral?) built-in format (id < 164)

A <- lapply(seq_along(numericFormats),
            function(i) cell(1, i, value=numericFormats[i]))
B <- lapply(seq_along(numericFormats),
            function(i) cell(2, i, value=-9999.1234,
                             numberFormat=numericFormats[i]))
sheet1 <- list(numericFormats = do.call(c, c(A,B)))
A <- lapply(seq_along(dateFormats),
            function(i) cell(1, i, value=dateFormats[i]))
B <- lapply(seq_along(dateFormats),
            function(i) cell(2, i, value=10000,
                             numberFormat=dateFormats[i]))
sheet2 <- list(dateFormats = do.call(c, c(A,B)))

json <- jsonlite::toJSON(c(sheet1, sheet2), null="null", auto_unbox = TRUE)

json2xlsx(json, outfile="numberFormat.xlsx", overwrite=TRUE)

dd <- readxl::read_xlsx("numberFormat.xlsx", sheet=2, col_names=FALSE)

