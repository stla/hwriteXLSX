setwd("~/MyPackages/hwriteXLSX/inst/testing")
library(hwriteXLSX)


library(data.table)

set.seed(666)
dat <- data.table(x   = rnorm(100),
                  y   = rnorm(100),
                  day = Sys.Date() + 1:100)

library(dict)
D <- dict()

for(j in seq_along(dat)){
  D[[c(1,j)]] <- cell(j, 1, names(dat)[j], fontname="Verdana", bold=TRUE)
  column <- dat[[j]]
  for(i in seq_along(column)){
    if(j<3){
      value <- column[[i]]
      if(abs(value) < 2){
        D[[c(i+1,j)]] <- cell(j, i+1, value)
      }else{
        D[[c(i+1,j)]] <- cell(j, i+1, value, color="red", bold=TRUE, comment="outlier")
      }
    }else{
      D[[c(i+1,j)]] <- cellDate(j, i+1, column[[i]])
    }
  }
}

sheet <- list(Sheet1 = do.call(c, D$values()))

jsonSheet <- jsonlite::toJSON(sheet, null="null", auto_unbox = TRUE)

#json2xlsx(jsonSheet, outfile="example00.xlsx")

#### add plots ####
library(ggplot2)
gg_scatter <- ggplot(dat, aes(x=x, y=y)) + geom_point()
gg_tsx <-
  ggplot(dat, aes(x=day, y=x)) + geom_point() +
  scale_x_date(date_breaks = "1 week", date_minor_breaks = "1 day",
               date_labels = "%Y-%m-%d") +
  theme(axis.text.x = element_text(angle=45, vjust = 0.5),
        plot.margin = margin(t=0.1, r=1, b=0.1, l=0, "cm"))
gg_tsy <-
  ggplot(dat, aes(x=day, y=y)) + geom_point() +
  scale_x_date(date_breaks = "1 week", date_minor_breaks = "1 day",
               date_labels = "%Y-%m-%d") +
  theme(axis.text.x = element_text(angle=45, vjust = 0.5),
        plot.margin = margin(t=0.1, r=1, b=0.1, l=0, "cm"))
# ggsave("scatter.png", gg_scatter, width=6, height=6, units="cm")
# ggsave("ts_x.png", gg_tsx, width=6, height=4, units="cm")
# ggsave("ts_y.png", gg_tsx, width=12, height=8, units="cm")
imageFile1 <- "scatter.png"
imageFile2 <- "ts_x.png"
imageFile3 <- "ts_y.png"
# imageFile1 <- "C:/zTemp/scatter.png"
# imageFile2 <- "C:/zTemp/ts_x.png"
# imageFile3 <- "C:/zTemp/ts_y.png"
ggsave(imageFile1, gg_scatter, width=12, height=12, units="cm")
ggsave(imageFile2, gg_tsx, width=12, height=8, units="cm")
ggsave(imageFile3, gg_tsx, width=12, height=8, units="cm")

# plante si fichier sur disque U ! SOLVED with strict bytestrings !!
image1 <- list(file=imageFile1, left=5, top=1, width=400, height=400)
image2 <- list(file=imageFile2, left=5, top=22, width=400, height=267)
image3 <- list(file=imageFile3, left=5, top=36, width=400, height=267)
images <- list(Sheet1 = list(image1, image2, image3))
jsonImages <- jsonlite::toJSON(images, auto_unbox = TRUE)

json2xlsx(jsonSheet, jsonImages, outfile="example000testU.xlsx", overwrite = TRUE)
