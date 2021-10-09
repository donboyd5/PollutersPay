
# libraries ----

library(tidyverse)
options(tibble.print_max = 75, tibble.print_min = 75) # if more than 60 rows, print 60 - enough for states
# ggplot2 tibble tidyr readr purrr dplyr stringr forcats
library(scales)

library(readxl)
library(btools)
library(kableExtra)

# devtools::install_version("openxlsx", version = "4.2.3", repos = "http://cran.us.r-project.org")  # 4.2.4 broke my code
library(openxlsx)
library(gt)

library(foreign)

# constants ----
aaiidir <- r"(C:\Program Files (x86)\Stock Investor\Professional\)"
ddict <- paste0(aaiidir, "Datadict/")
ddict

ddbfs <- paste0(aaiidir, "Dbfs/")


# look at dbf files ----
pdictdbfs <- list.files(path=ddict, pattern="dbf", full.names=TRUE)
pdictdbfs

pddbfs <- list.files(path=ddbfs, pattern="dbf", full.names=TRUE)
pddbfs

# read.dbf(file=dictdbfs[1]) %>% head

for(f in pdictdbfs){
  df <- read.dbf(file=f)
  print(f)
  print(head(df))
}

for(f in pddbfs){
  df <- read.dbf(file=f)
  print(f)
  print(head(df))
}

fdbf <- function(f){
  df <- read.dbf(file=f, as.is=TRUE)
  flist <- list()
  flist$fname=f
  flist$names=names(df)
  flist$nrow=nrow(df)
  flist$ncol=ncol(df)
  return(flist)
}

fdbf(pddbfs[1])

info <- purrr::map(pddbfs, fdbf)
info

for(f in pddbfs){
  df <- read.dbf(file=f)
  print(f)
  print(head(df))
}







