
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

# data ----
load(here::here("data", "report_data.rdata"), verbose=TRUE)
load(file=here::here("data", "maindata.Rdata"), verbose=TRUE)

# global6518  # global emissions 1965-2018, MtCO2e
# ongl_factors  # conversion factors
# ng_factors  # conversion factors
# coal_factors  # conversion factors
# uninames  # df with uid, uninmame, plus alternative names used in files
# ucombo6518  # df with uid uniname rank ptype co2 fmethane total fmethane_ch4; total includes coal and cement
# usumrank_long  # allows computing coal and cement percents for the period covered
comment(udca_base)
comment(uninames)
comment(ucombo6518)
comment(usumrank_long)
comment(global6518)
comment(global6518_xcement)
comment(global_shares)


# analysis ----
glimpse(entities)

# get all 108 entities
dca_total_max <- 4000 # $ billions
dca_total_annual <- 50 # $ billions
threshold_proportion <- 0.0005

tax_base <- entities  %>%
  select(uid, table_name, ptype, group, label, reachable, dca_base, pct_emissions) %>%
  mutate(payor=str_detect(reachable, "yes") & 
           (!uid %in% coal_drops) &
           (pct_emissions >= threshold_proportion),
         payor_dca_base=sum(ifelse(payor, dca_base, 0)),
         pct_payors=ifelse(payor, dca_base / payor_dca_base, 0),
         dca_max = pct_emissions * dca_total_max,
         dca_annual = pct_payors * dca_total_annual) %>% 
  arrange(payor, -pct_emissions)
glimpse(tax_base)

tax_base %>%
  filter(payor) %>%
  group_by(label) %>%
  summarise(n=n(),
            across(c(pct_emissions, pct_payors, dca_annual, dca_max), sum))

sum(tax_base$dca_annual)
sum(tax_base$dca_max)

tax_base %>%
  filter(payor) %>%
  arrange(desc(dca_annual)) %>%
  select(uid, table_name, group, dca_annual)


# do any of the payors or nonpayors look wrong?
# payors
tax_base %>%
  filter(payor) %>%
  select(uid, table_name, group, reachable, dca_annual) %>%
  arrange(group, desc(dca_annual))
# looks good - maybe Libya National Oil and Sinopec are questionable; maybe Petronas, PetroChina questionable

# nonpayors
tax_base %>%
  filter(!payor) %>%
  select(uid, table_name, group, reachable, pct_emissions) %>%
  arrange(group, desc(pct_emissions))



# WORKBOOK EXCLUDING MAXIMUM DAMAGE ASSESSMENT ----
# workbook prep ----
# data we will use
df <- tax_base %>%
  mutate(blank1="", blank2="") # so we can have a blank column

use_names <- quote(c(table_name, pct_emissions, blank1, pct_payors, blank2, dca_annual))


# define styles ----
curr2 <- createStyle(numFmt = "$0.00")
num2 <- createStyle(numFmt = "0.00")
curr3 <- createStyle(numFmt = "$0.000")
num3 <- createStyle(numFmt = "0.000")

pct1 <- createStyle(numFmt = "0.0%")
pct2 <- createStyle(numFmt = "0.00%")

# for totals rows
totstyle <- createStyle(fontSize = 11, fgFill = "#f0f0f0", border = "TopBottom", borderColour = "#f0f0f0")

bold_style <- createStyle(fontSize = 11, textDecoration = "bold")

# start the workbook ----
outfile <- "PPCompanyTable_v1.xlsx"  # name of the output file
snames <- getSheetNames(here::here("data", "PPCompanyShell.xlsx"))
wb <- loadWorkbook(file = here::here("data", "PPCompanyShell.xlsx"))
# get the sheets we want
sname <- "PPCompanyTable"
cloneWorksheet(wb, sheetName=sname, clonedSheet = "TableShell_nomax")
purrr::walk(snames, function(x) removeWorksheet(wb, x))


last_line_written <- 0


#.. SOE ----
soe <- df %>%
  filter(payor, group %in% c("SOE", "nation")) %>%
  select(!!use_names) %>%
  arrange(desc(dca_annual))

soe_total <- soe %>%
  summarise(table_name=paste0("  Total for ", nrow(soe), " State-Owned Enterprises"), 
            pct_emissions=sum(pct_emissions), 
            blank1=first(blank1),
            pct_payors=sum(pct_payors),
            blank2=first(blank2),
            dca_annual=sum(dca_annual))

# write the details for SOE
soe_headrow <- 9
writeData(wb, sheet = sname, x="State-Owned Enterprises (SOEs)", startCol=1, startRow=soe_headrow, rowNames = FALSE, colNames = FALSE)
last_line_written <- soe_headrow

soe_start <- last_line_written + 1
writeData(wb, sheet = sname, x=soe, startCol=1, startRow=soe_start, rowNames = FALSE, colNames = FALSE)
last_line_written <- soe_start + nrow(soe)

# now write the SOE summary
soe_totalrow <- last_line_written + 1
writeData(wb, sheet = sname, x=soe_total, 
          startCol=1, startRow=soe_totalrow, rowNames = FALSE, colNames = FALSE)
last_line_written <- soe_totalrow


#.. foreign IOC ----
fioc <- df %>%
  filter(payor, group %in% c("fIOC")) %>%
  select(!!use_names) %>%
  arrange(desc(dca_annual))

fioc_total <- fioc %>%
  summarise(table_name=paste0("  Total for ", nrow(fioc), " Foreign Investor-Owned Companies"), 
            pct_emissions=sum(pct_emissions), 
            blank1=first(blank1),
            pct_payors=sum(pct_payors),
            blank2=first(blank2),
            dca_annual=sum(dca_annual))

# write the details for fioc
fioc_headrow <- last_line_written + 3
writeData(wb, sheet = sname, x="Foreign Investor-Owned Companies (IOCs)", startCol=1, startRow=fioc_headrow, rowNames = FALSE, colNames = FALSE)
last_line_written <- fioc_headrow

fioc_start <- last_line_written + 1
writeData(wb, sheet = sname, x=fioc, startCol=1, startRow=fioc_start, rowNames = FALSE, colNames = FALSE)
last_line_written <- fioc_start

# now write the fioc summary
fioc_totalrow <- last_line_written + nrow(fioc) + 1
writeData(wb, sheet = sname, x=fioc_total, 
          startCol=1, startRow=fioc_totalrow, rowNames = FALSE, colNames = FALSE)
last_line_written <- fioc_totalrow


#.. subtotal after SOE and fIOC ----
soefioc <- bind_rows(soe, fioc)
soefioc_total <- soefioc %>%
  summarise(table_name=paste0("  Combined Total for ", nrow(.), " SOEs and Foreign IOCs"), 
            pct_emissions=sum(pct_emissions), 
            blank1=first(blank1),
            pct_payors=sum(pct_payors),
            blank2=first(blank2),
            dca_annual=sum(dca_annual))

soefioc_totalrow <- last_line_written + 2
writeData(wb, sheet = sname, x=soefioc_total, 
          startCol=1, startRow=soefioc_totalrow, rowNames = FALSE, colNames = FALSE)
last_line_written <- soefioc_totalrow



#.. domestic IOC ----
dioc <- df %>%
  filter(payor, group %in% c("dIOC")) %>%
  select(!!use_names) %>%
  arrange(desc(dca_annual))

dioc_total <- dioc %>%
  summarise(table_name=paste0("  Total for ", nrow(dioc), " Domestic Investor-Owned Companies"), 
            pct_emissions=sum(pct_emissions), 
            blank1=first(blank1),
            pct_payors=sum(pct_payors),
            blank2=first(blank2),
            dca_annual=sum(dca_annual))

# write the details for dioc
dioc_headrow <- last_line_written + 3
writeData(wb, sheet = sname, x="Domestic Investor-Owned Companies", startCol=1, startRow=dioc_headrow, rowNames = FALSE, colNames = FALSE)
last_line_written <- dioc_headrow

dioc_start <- last_line_written + 1
writeData(wb, sheet = sname, x=dioc, startCol=1, startRow=dioc_start, rowNames = FALSE, colNames = FALSE)
last_line_written <- dioc_start + nrow(dioc)


# now write the dioc summary
dioc_totalrow <- last_line_written + 1
writeData(wb, sheet = sname, x=dioc_total, 
          startCol=1, startRow=dioc_totalrow, rowNames = FALSE, colNames = FALSE)
last_line_written <- dioc_totalrow


# ..grand totals ----
gtdata <- df %>%
  filter(payor, group %in% c("SOE", "nation", "fIOC", "dIOC"))
grand_total <- gtdata %>%
  summarise(table_name=paste0("    Grand Total for ", nrow(gtdata), " Entities"), 
            pct_emissions=sum(pct_emissions),
            blank1=first(blank1),
            pct_payors=sum(pct_payors),
            blank2=first(blank2),
            dca_annual=sum(dca_annual))
grand_totalrow <- last_line_written + 2
writeData(wb, sheet = sname, x=grand_total, 
          startCol=1, startRow=grand_totalrow, rowNames = FALSE, colNames = FALSE)
last_line_written <- grand_totalrow


# write notes ----
note_intro <- "Notes:"
note1 <- "  * Assumes the Climage Damage Compensation Tax is only applied to polluters accounting for at least 0.05% of historical global emissions."

note_start <- last_line_written + 3

writeData(wb, sheet = sname, x=note_intro, 
          startCol=1, startRow=note_start, rowNames = FALSE, colNames = FALSE)
writeData(wb, sheet = sname, x=note1, 
          startCol=1, startRow=note_start + 1, rowNames = FALSE, colNames = FALSE)


# apply styles ----
pct_cols <- c(2, 4)
dollar_cols <- 6

# default styles for the workbook
first_data_row <- 10
addStyle(wb, sheet = sname, style = pct2, rows=first_data_row:200, cols = pct_cols, gridExpand = TRUE)
addStyle(wb, sheet = sname, style = curr3, rows=first_data_row, cols = dollar_cols)
addStyle(wb, sheet = sname, style = num3, rows=(first_data_row + 1):200, cols = dollar_cols, gridExpand = TRUE)

# styles for each block
addStyle(wb, sheet = sname, style = totstyle, rows = soe_headrow, cols = 1:6, stack=TRUE)
addStyle(wb, sheet = sname, style = bold_style, rows = soe_headrow, cols = 1, stack=TRUE)
addStyle(wb, sheet = sname, style = curr3, rows=soe_totalrow, cols = dollar_cols)
addStyle(wb, sheet = sname, style = totstyle, rows = soe_totalrow, cols = 1:6, stack=TRUE)

addStyle(wb, sheet = sname, style = totstyle, rows = fioc_headrow, cols = 1:6, stack=TRUE)
addStyle(wb, sheet = sname, style = bold_style, rows = fioc_headrow, cols = 1, stack=TRUE)
addStyle(wb, sheet = sname, style = curr3, rows=fioc_totalrow, cols = dollar_cols)
addStyle(wb, sheet = sname, style = totstyle, rows = fioc_totalrow, cols = 1:6, stack=TRUE)

addStyle(wb, sheet = sname, style = curr3, rows=soefioc_totalrow, cols = dollar_cols)
addStyle(wb, sheet = sname, style = totstyle, rows = soefioc_totalrow, cols = 1:6, stack=TRUE)

addStyle(wb, sheet = sname, style = totstyle, rows = dioc_headrow, cols = 1:6, stack=TRUE)
addStyle(wb, sheet = sname, style = bold_style, rows = dioc_headrow, cols = 1, stack=TRUE)
addStyle(wb, sheet = sname, style = curr3, rows=dioc_totalrow, cols = dollar_cols)
addStyle(wb, sheet = sname, style = totstyle, rows = dioc_totalrow, cols = 1:6, stack=TRUE)

addStyle(wb, sheet = sname, style = curr3, rows=grand_totalrow, cols = dollar_cols)
addStyle(wb, sheet = sname, style = totstyle, rows = grand_totalrow, cols = 1:6, stack=TRUE)
addStyle(wb, sheet = sname, style = bold_style, rows = grand_totalrow, cols = 1, stack=TRUE)


# save workbook ----

saveWorkbook(wb, file = here::here("results", outfile), overwrite = TRUE)
openXL(here::here("results", outfile))











# 
# 
# # WORKBOOK INCLUDING MAXIMUM DAMAGE ASSESSMENT ----
# # workbook prep ----
# # data we will use
# df <- tax_base %>%
#   mutate(blank="") # so we can have a blank column
# 
# 
# # define styles ----
# curr2 <- createStyle(numFmt = "$0.00")
# num2 <- createStyle(numFmt = "0.00")
# curr3 <- createStyle(numFmt = "$0.000")
# num3 <- createStyle(numFmt = "0.000")
# 
# pct1 <- createStyle(numFmt = "0.0%")
# pct2 <- createStyle(numFmt = "0.00%")
# 
# # for totals rows
# totstyle <- createStyle(fontSize = 11, fgFill = "#f0f0f0", border = "TopBottom", borderColour = "#f0f0f0")
# 
# bold_style <- createStyle(fontSize = 11, textDecoration = "bold")
# 
# # start the workbook ----
# outfile <- "PPCompanyTable_v1.xlsx"  # name of the output file
# wb <- loadWorkbook(file = here::here("data", "PPCompanyShell.xlsx"))
# # get the sheets we want
# sname <- "PPCompanyTable"
# cloneWorksheet(wb, sheetName=sname, clonedSheet = "TableShell")
# removeWorksheet(wb, "TableShell")
# 
# 
# 
# #.. SOE ----
# soe <- df %>%
#   filter(payor, group %in% c("SOE", "nation")) %>%
#   select(table_name, pct_emissions, dca_max, blank, pct_payors, dca_annual) %>%
#   arrange(desc(dca_annual))
# 
# soe_total <- soe %>%
#   summarise(table_name=paste0("  Total for ", nrow(soe), " State-Owned Enterprises"), 
#             pct_emissions=sum(pct_emissions), 
#             dca_max=sum(dca_max),
#             blank=first(blank),
#             pct_payors=sum(pct_payors),
#             dca_annual=sum(dca_annual))
# 
# # write the details for SOE
# soe_headrow <- 9
# writeData(wb, sheet = sname, x="State-Owned Enterprises", startCol=1, startRow=soe_headrow, rowNames = FALSE, colNames = FALSE)
# 
# soe_start <- soe_headrow + 1
# writeData(wb, sheet = sname, x=soe, startCol=1, startRow=soe_start, rowNames = FALSE, colNames = FALSE)
# 
# # now write the SOE summary
# soe_totalrow <- soe_start + nrow(soe) + 1
# writeData(wb, sheet = sname, x=soe_total, 
#           startCol=1, startRow=soe_totalrow, rowNames = FALSE, colNames = FALSE)
# 
# 
# 
# #.. foreign IOC ----
# fioc <- df %>%
#   filter(payor, group %in% c("fIOC")) %>%
#   select(table_name, pct_emissions, dca_max, blank, pct_payors, dca_annual) %>%
#   arrange(desc(dca_annual))
# 
# fioc_total <- fioc %>%
#   summarise(table_name=paste0("  Total for ", nrow(fioc), " Foreign Investor-Owned Companies"), 
#             pct_emissions=sum(pct_emissions), 
#             dca_max=sum(dca_max),
#             blank=first(blank),
#             pct_payors=sum(pct_payors),
#             dca_annual=sum(dca_annual))
# 
# # write the details for fioc
# fioc_headrow <- soe_totalrow + 3
# writeData(wb, sheet = sname, x="Foreign Investor-Owned Companies", startCol=1, startRow=fioc_headrow, rowNames = FALSE, colNames = FALSE)
# 
# fioc_start <- fioc_headrow + 1
# writeData(wb, sheet = sname, x=fioc, startCol=1, startRow=fioc_start, rowNames = FALSE, colNames = FALSE)
# 
# # now write the fioc summary
# fioc_totalrow <- fioc_start + nrow(fioc) + 1
# writeData(wb, sheet = sname, x=fioc_total, 
#           startCol=1, startRow=fioc_totalrow, rowNames = FALSE, colNames = FALSE)
# 
# 
# #.. domestic IOC ----
# dioc <- df %>%
#   filter(payor, group %in% c("dIOC")) %>%
#   select(table_name, pct_emissions, dca_max, blank, pct_payors, dca_annual) %>%
#   arrange(desc(dca_annual))
# 
# dioc_total <- dioc %>%
#   summarise(table_name=paste0("  Total for ", nrow(dioc), " Domestic Investor-Owned Companies"), 
#             pct_emissions=sum(pct_emissions), 
#             dca_max=sum(dca_max),
#             blank=first(blank),
#             pct_payors=sum(pct_payors),
#             dca_annual=sum(dca_annual))
# 
# # write the details for dioc
# dioc_headrow <- fioc_totalrow + 3
# writeData(wb, sheet = sname, x="Domestic Investor-Owned Companies", startCol=1, startRow=dioc_headrow, rowNames = FALSE, colNames = FALSE)
# 
# dioc_start <- dioc_headrow + 1
# writeData(wb, sheet = sname, x=dioc, startCol=1, startRow=dioc_start, rowNames = FALSE, colNames = FALSE)
# 
# # now write the dioc summary
# dioc_totalrow <- dioc_start + nrow(dioc) + 1
# writeData(wb, sheet = sname, x=dioc_total, 
#           startCol=1, startRow=dioc_totalrow, rowNames = FALSE, colNames = FALSE)
# 
# 
# # ..grand totals ----
# gtdata <- df %>%
#   filter(payor, group %in% c("SOE", "nation", "fIOC", "dIOC"))
# grand_total <- gtdata %>%
#   summarise(table_name=paste0("    Grand Total for ", nrow(gtdata), " Entities"), 
#             pct_emissions=sum(pct_emissions), 
#             dca_max=sum(dca_max),
#             blank=first(blank),
#             pct_payors=sum(pct_payors),
#             dca_annual=sum(dca_annual))
# grand_totalrow <- dioc_totalrow + 2
# writeData(wb, sheet = sname, x=grand_total, 
#           startCol=1, startRow=grand_totalrow, rowNames = FALSE, colNames = FALSE)
# 
# 
# # write notes ----
# note_intro <- "Notes:"
# note1 <- "  *  Maximum Cumulative Damage Tax Based on Assumed $4 Trillion Total Damages."
# note2 <- "  ** Assumes the Climage Damage Compensation Tax is only applied to polluters accounting for at least 0.05% of historical global emissions."
# 
# note_start <- grand_totalrow + 3
# 
# writeData(wb, sheet = sname, x=note_intro, 
#           startCol=1, startRow=note_start, rowNames = FALSE, colNames = FALSE)
# writeData(wb, sheet = sname, x=note1, 
#           startCol=1, startRow=note_start + 1, rowNames = FALSE, colNames = FALSE)
# writeData(wb, sheet = sname, x=note2, 
#           startCol=1, startRow=note_start + 2, rowNames = FALSE, colNames = FALSE)
# 
# 
# # apply styles ----
# 
# # default styles for the workbook
# first_data_row <- 10
# addStyle(wb, sheet = sname, style = pct2, rows=first_data_row:200, cols = c(2, 5), gridExpand = TRUE)
# addStyle(wb, sheet = sname, style = curr3, rows=first_data_row, cols = c(3, 6))
# addStyle(wb, sheet = sname, style = num3, rows=(first_data_row + 1):200, cols = c(3, 6), gridExpand = TRUE)
# 
# # styles for each block
# addStyle(wb, sheet = sname, style = totstyle, rows = soe_headrow, cols = 1:6, stack=TRUE)
# addStyle(wb, sheet = sname, style = bold_style, rows = soe_headrow, cols = 1, stack=TRUE)
# addStyle(wb, sheet = sname, style = curr3, rows=soe_totalrow, cols = c(3, 6))
# addStyle(wb, sheet = sname, style = totstyle, rows = soe_totalrow, cols = 1:6, stack=TRUE)
# 
# addStyle(wb, sheet = sname, style = totstyle, rows = fioc_headrow, cols = 1:6, stack=TRUE)
# addStyle(wb, sheet = sname, style = bold_style, rows = fioc_headrow, cols = 1, stack=TRUE)
# addStyle(wb, sheet = sname, style = curr3, rows=fioc_totalrow, cols = c(3, 6))
# addStyle(wb, sheet = sname, style = totstyle, rows = fioc_totalrow, cols = 1:6, stack=TRUE)
# 
# addStyle(wb, sheet = sname, style = totstyle, rows = dioc_headrow, cols = 1:6, stack=TRUE)
# addStyle(wb, sheet = sname, style = bold_style, rows = dioc_headrow, cols = 1, stack=TRUE)
# addStyle(wb, sheet = sname, style = curr3, rows=dioc_totalrow, cols = c(3, 6))
# addStyle(wb, sheet = sname, style = totstyle, rows = dioc_totalrow, cols = 1:6, stack=TRUE)
# 
# addStyle(wb, sheet = sname, style = curr3, rows=grand_totalrow, cols = c(3, 6))
# addStyle(wb, sheet = sname, style = totstyle, rows = grand_totalrow, cols = 1:6, stack=TRUE)
# addStyle(wb, sheet = sname, style = bold_style, rows = grand_totalrow, cols = 1, stack=TRUE)
# 
# # save workbook ----
# 
# saveWorkbook(wb, file = here::here("results", outfile), overwrite = TRUE)
# openXL(here::here("results", outfile))





