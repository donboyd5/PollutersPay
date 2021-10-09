
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


# get new good data ----
load(file=here::here("data", "constants.rdata"), verbose=TRUE)  # overwrite maindata constants
load(file=here::here("data", "pollpay_data.Rdata"), verbose=TRUE)
xwalk <- readRDS(here::here("data", "xwalk.rds"))

# get sloppy excel revenue and profits data ----
fn <- "top108(19).xlsx"
xldata1 <- read_excel(here::here("data_raw", fn), sheet="receipts_table", range="A3:AU111", col_types = "text")
glimpse(xldata1)

# convert to numeric, put proper uid on the file
xldata <- xldata1 %>%
  select(uid_old=uid, uniname, table_name, ptype, symbol, fd, country, total_revenue_millions,
         dca_base, reachable, starts_with("receipts20"), starts_with("profits20")) %>%
  mutate(uid_old=as.integer(uid_old),
         across(c(total_revenue_millions, dca_base, starts_with("receipts20"), starts_with("profits20")),
                as.numeric)) %>%
  left_join(xwalk %>% select(uid, uname, uid_old, uname_old), by = "uid_old") %>%
  select(-c(uid_old, uniname, uname_old))  # after verifying that all is good

# xldata %>%
#   select(uid, uniname, ptype, fd, total_revenue_millions, receipts2019) %>%
#   filter(row_number() >= 70)

# compute average receipts and profits values for each company
findata <- xldata %>%
  select(uid, uname, receipts2019:receipts2016, profits2019:profits2016) %>%
  pivot_longer(cols = -c(uid, uname)) %>%
  filter(!is.na(value)) %>%
  mutate(fintype=ifelse(str_detect(name, "receipts"), "receipts", "profits"),
         year=str_sub(name, -4, -1) %>% as.integer) %>%
  arrange(uname, uid, fintype, year) %>%
  group_by(uid, uname, fintype) %>%
  summarise(nnotna=sum(!is.na(value)),
            mean=mean(value, na.rm=TRUE),
            min=min(value, na.rm=TRUE),
            max=max(value, na.rm=TRUE),
            fyear=year[which.min(year)],
            lyear=year[which.max(year)],
            .groups = "drop")

finwide <- findata %>%
  pivot_wider(names_from = fintype,
              values_from = -c(uid, uname, fintype),
              names_glue = "{fintype}_{.value}") %>%
  select(uid, uname, starts_with("profits"), starts_with("receipts"))
# note that we only have profits for 36 firms


# start data ----
# now get data needed for tax, to compare to profits
# calc global totals 
# years <- 1965:2018  # 108 entities
# years <- 1993:2018  # we only have 105 entities in this period - lose Cyprus Amax, FSU, Czech coal
years <- 2000:2018  # only 104 entities

gtots <- global_totals %>%
  filter(year %in% years) %>%
  group_by(fuel) %>%
  summarise(value=sum(value, na.rm=TRUE)) %>%
  pivot_wider(names_from = fuel) %>%
  mutate(xcement=total - cement,
         xcc=total - cement - coal)
# matches global6518 1409854
# xcement is 1372735 vs global6518_xcement 1380823
# 1372735 / 1380823 # 99.4%, seems ok

# check dca_base
# estimate dca base as total less coal, cement ----
entities <- eachevery %>%
  filter(etype=="MtCO2e", year %in% years) %>%
  group_by(uid, uname, ptype) %>%
  summarise(mtco2e=sum(value, na.rm=TRUE), .groups="drop")

dca_base <- entities %>%
  left_join(coalcement_17502018 %>% select(uid, contains("share")),
            by="uid") %>%
  mutate(dca_base=mtco2e * (1 - coal_share - cement_share))


# determine who is reachable ----
nexus <- xldata %>%
  select(uid, uname, ptype, fd, reachable) %>%
  left_join(dca_base %>% select(-ptype), by = c("uid", "uname")) %>%  # use ptype from the Excel file
  mutate(gtot_xcement=gtots$xcement,
         pct_emissions=dca_base / gtot_xcement, # really a proportion, not a percent  # used to be global6518_xcement
         nexus=ifelse(str_detect(reachable, coll("yes")), TRUE, FALSE)) 
summary(nexus)  # this has all 108 entities

nexus %>% filter(is.na(dca_base))


# analysis ----
glimpse(entities)

# get all 108 entities
dca_total_max <- 4000 # $ billions
dca_total_annual <- 50 # $ billions
threshold_proportion <- 0.0005

# use coal_drops to be safe but issue should never arise per code below
coal_drops <- c(7, 11, 16, 19, 21, 24, 25, 36, 44, 52, 56, 57, 74, 83, 99, 100, 104)

# taxpayers  %>%
#   select(uid, uname, ptype, reachable, dca_base, pct_emissions) %>%
#   filter(uid %in% coal_drops)

# taxpayers  %>%
#   filter(taxable) %>%
#   select(uid, uname, ptype, reachable, dca_base, pct_emissions) %>%
#   filter(uid %in% coal_drops)

tax_base <- nexus %>%
  select(uid, uname, ptype, fd, reachable, nexus, pct_emissions, dca_base) %>%
  mutate(payor=nexus & 
           (!uid %in% coal_drops) &
           (!is.na(pct_emissions)) &
           (pct_emissions >= threshold_proportion),
         payor_dca_base=sum(ifelse(payor, dca_base, 0)),
         pct_payors=ifelse(payor, dca_base / payor_dca_base, 0),
         dca_max = pct_emissions * dca_total_max,
         dca_annual = pct_payors * dca_total_annual) %>%
  mutate(group=case_when(ptype=="Nation-State" ~ "nation",
                         ptype=="SOE" ~ "SOE",
                         ptype=="IOC" & fd=="foreign" ~ "fIOC",
                         ptype=="IOC" & fd=="domestic" ~ "dIOC",
                         TRUE ~ "ERROR"),
         label=factor(group,
                      levels=c("nation", "SOE", "fIOC", "dIOC"),
                      labels=c("Nation-State Entity",
                               "State-Owned Enterprise",
                               "Foreign Investor-Owned Company",
                               "Domestic Investor-Owned Company"))) %>% 
  arrange(payor, -pct_emissions)
glimpse(tax_base)
summary(tax_base)
tax_base %>% filter(is.na(payor))

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
  select(uid, uname, group, dca_annual)


# do any of the payors or nonpayors look wrong?
# payors
tax_base %>%
  filter(payor) %>%
  select(uid, uname, group, reachable, dca_annual) %>%
  arrange(group, desc(dca_annual))
# looks good - maybe Libya National Oil and Sinopec are questionable; maybe Petronas, PetroChina questionable

# nonpayors
tax_base %>%
  filter(!payor) %>%
  select(uid, uname, group, reachable, pct_emissions) %>%
  arrange(group, desc(pct_emissions))

# add profits to the mix ----

final <- tax_base %>%
  left_join(finwide %>%
              select(uid, profits_mean, profits_min, profits_max), 
            by = "uid") %>%
  mutate(across(c(profits_mean, profits_min, profits_max),
                ~ .x / 1000),  # put into $ billions
         pct_mean=dca_annual / profits_mean,
         pct_min = dca_annual / profits_min,
         pct_max = dca_annual / profits_max)
summary(final)



# CREATE WORKBOOK EXCLUDING MAXIMUM DAMAGE ASSESSMENT ----
# workbook prep ----
# data we will use
df <- final %>%
  mutate(blank1="", blank2="", blank3="") # so we can have blank columns

use_names <- quote(c(uname, pct_emissions, blank1, pct_payors, blank2, dca_annual,
                     blank3, profits_mean, pct_mean))


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
outfile <- "PPCompanyProfitsTable_v1.xlsx"  # name of the output file
shellfile <- "PPCompanyShell_profits.xlsx"
snames <- getSheetNames(here::here("data_raw", shellfile))
wb <- loadWorkbook(file = here::here("data_raw", shellfile))
# get the sheets we want
sname <- "ProfitsTable"
cloneWorksheet(wb, sheetName=sname, clonedSheet = "TableShell")
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
            dca_annual=sum(dca_annual),
            blank3=first(blank3))

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
pct_cols <- c(2, 4, 9)
dollar_cols <- c(6, 8)
last_col <- 9

# default styles for the workbook
first_data_row <- 10
addStyle(wb, sheet = sname, style = pct2, rows=first_data_row:200, cols = pct_cols, gridExpand = TRUE)
addStyle(wb, sheet = sname, style = curr3, rows=first_data_row, cols = dollar_cols)
addStyle(wb, sheet = sname, style = num3, rows=(first_data_row + 1):200, cols = dollar_cols, gridExpand = TRUE)

# styles for each block
addStyle(wb, sheet = sname, style = totstyle, rows = soe_headrow, cols = 1:last_col, stack=TRUE)
addStyle(wb, sheet = sname, style = bold_style, rows = soe_headrow, cols = 1, stack=TRUE)
addStyle(wb, sheet = sname, style = curr3, rows=soe_totalrow, cols = dollar_cols)
addStyle(wb, sheet = sname, style = totstyle, rows = soe_totalrow, cols = 1:last_col, stack=TRUE)

addStyle(wb, sheet = sname, style = totstyle, rows = fioc_headrow, cols = 1:last_col, stack=TRUE)
addStyle(wb, sheet = sname, style = bold_style, rows = fioc_headrow, cols = 1, stack=TRUE)
addStyle(wb, sheet = sname, style = curr3, rows=fioc_totalrow, cols = dollar_cols)
addStyle(wb, sheet = sname, style = totstyle, rows = fioc_totalrow, cols = 1:last_col, stack=TRUE)

addStyle(wb, sheet = sname, style = curr3, rows=soefioc_totalrow, cols = dollar_cols)
addStyle(wb, sheet = sname, style = totstyle, rows = soefioc_totalrow, cols = 1:last_col, stack=TRUE)

addStyle(wb, sheet = sname, style = totstyle, rows = dioc_headrow, cols = 1:last_col, stack=TRUE)
addStyle(wb, sheet = sname, style = bold_style, rows = dioc_headrow, cols = 1, stack=TRUE)
addStyle(wb, sheet = sname, style = curr3, rows=dioc_totalrow, cols = dollar_cols)
addStyle(wb, sheet = sname, style = totstyle, rows = dioc_totalrow, cols = 1:last_col, stack=TRUE)

addStyle(wb, sheet = sname, style = curr3, rows=grand_totalrow, cols = dollar_cols)
addStyle(wb, sheet = sname, style = totstyle, rows = grand_totalrow, cols = 1:last_col, stack=TRUE)
addStyle(wb, sheet = sname, style = bold_style, rows = grand_totalrow, cols = 1, stack=TRUE)


# save workbook ----

saveWorkbook(wb, file = here::here("results", outfile), overwrite = TRUE)
openXL(here::here("results", outfile))


