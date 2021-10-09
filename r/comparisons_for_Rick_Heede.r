
# libraries ----

library(tidyverse)
options(tibble.print_max = 65, tibble.print_min = 65) # if more than 60 rows, print 60 - enough for states
# ggplot2 tibble tidyr readr purrr dplyr stringr forcats
library(scales)
library(stringi)

library(readxl)
library(btools)
library(kableExtra)

# require(devtools)
# install_version("openxlsx", version = "4.2.3", repos = "http://cran.us.r-project.org")  # 4.2.4 broke my code
library(openxlsx)
library(gt)

# functions ----


# get new and old data ----
#.. safely get selected objects from old data ----
path <- here::here("data", "maindata.rdata")
load(path, olddata <- new.env())
names(olddata)
udca_base <- olddata$udca_base

#.. new data ----
load(file=here::here("data", "constants.rdata"), verbose=TRUE)  # overwrite maindata constants
load(file=here::here("data", "pollpay_data.Rdata"), verbose=TRUE)

# make new-old crosswalk ----
dnew <- unames %>% select(uid, uname, pname_oil)
#  mutate(sound=refinedSoundex(uname, clean=FALSE, maxCodeLen=12))
dold <- olddata$uninames %>% select(uid_old=uid, uname_old=uniname, pname_oil=pname1_sumfiles)
xwalk <- full_join(dnew, dold, by = "pname_oil")
rm(dnew, dold)


# functions to get globals and dca_base for a given set of years ----
gtotals <- function(years){
  global_totals %>%
    filter(year %in% years) %>%
    group_by(fuel) %>%
    summarise(value=sum(value, na.rm=TRUE)) %>%
    pivot_wider(names_from = fuel) %>%
    mutate(
      other=total_xmeth - cement - coal,
      xcement=total - cement,
      fyear=years[1], lyear=years[length(years)],
      period=paste0(fyear, "-", lyear)) %>%
    select(period, other, cement, coal, total_xmeth, methane, total, xcement)
}


etotals <- function(years){
  entities <- eachevery %>%
    filter(etype=="MtCO2e", year %in% years) %>%
    group_by(uid, uname, ptype) %>%
    summarise(totemit=sum(value, na.rm=TRUE), .groups="drop")
  
  ecoalcem <- entities %>%
    left_join(coalcement_17502018 %>% select(uid, contains("share")),
              by="uid") %>%
    mutate(cement=totemit * cement_share,
           coal=totemit * coal_share,
           xcement=totemit - cement,
           xcemcoal = xcement - coal,
           fyear=years[1], lyear=years[length(years)],
           period=paste0(fyear, "-", lyear)) %>%
    select(uid, uname, ptype, period, totemit, cement, xcement, coal, xcemcoal)
  
  ecoalcem
}

get_base <- function(years){
  gtots <- gtotals(years)
  etots <- etotals(years)
  dca_base <- etots %>%
    mutate(gcement=gtots$cement,
           gxcement=gtots$xcement,
           gtotal=gtots$total,
           gcoal=gtots$coal)
}


# get reachable data ----
reachable1 <- read_excel(here::here("data_raw", "top108(18).xlsx"),
                        sheet="receipts_table",
                        range="A3:W111",
                        col_names = TRUE, col_types = "text") %>%
  mutate(uid=as.integer(uid))

reachable2 <- reachable1 %>%
  select(uid_old=uid, uname_old=uniname, fd, ptype, reachable) %>%
  full_join(xwalk, by=c("uid_old", "uname_old"))

reachable <- reachable2 %>%
  select(uid, uname, ptype, fd, reachable) %>%
  mutate(reach=case_when(str_detect(reachable, "yes") ~ TRUE,
                         TRUE ~ FALSE))
count(reachable, reach)
count(reachable, ptype, reach)


# get base data ----
compare <- bind_rows(get_base(1965:2018),
                     get_base(2000:2018),
                     get_base(2000:2017))

compare2 <- compare %>%
  left_join(reachable %>% select(uid, fd, reach, reachable),
            by="uid") %>%
  mutate(emit_pct=xcement / gxcement,
         noncc_pct=xcemcoal / gxcement,
         payor=(noncc_pct >= 0.0005) & reach) %>%
  unite(col="group", ptype, fd)

count(compare2, period, payor)
count(compare2 %>% filter(payor), period, group)

compare2 %>%
  filter(period=="1965-2018", payor) %>%
  arrange(group, desc(noncc_pct)) %>%
  select(uid, uname, group, xcemcoal, noncc_pct)


compare3 <- compare2 %>%
  group_by(period) %>%
  mutate(paybase=payor * xcemcoal,
         paypct=paybase / sum(paybase))
summary(compare3)

compare3 %>% filter(str_detect(uname, "Chesa"))


compwide <- compare3 %>%
  mutate(assess=paypct * 50) %>%
  filter(payor) %>%
  select(uid, uname, group, period, assess) %>%
  pivot_wider(names_from = period, values_from = assess, values_fill = 0) %>%
  mutate(d0018_6518=`2000-2018` - `1965-2018`,
         d0018_0017=`2000-2018` - `2000-2017`) %>%
  arrange(desc(group), desc(`1965-2018`))

tab <- compwide %>%
  select(-uid) %>%
  gt(groupname_col = "group") %>%
  summary_rows(
    groups = TRUE,
    columns = -uname,
    fns = list(total = "sum"),
    formatter = fmt_currency, decimals = 3
  ) %>%
  grand_summary_rows(
    columns = -uname,
    fns = list(total = "sum"),
    formatter = fmt_currency, decimals = 3
  ) %>%
  tab_header(
    title = "Climate damage tax using different emissions periods",
    subtitle = "$50 billion aggregate tax"
  ) %>% 
  cols_label(
    uname = "",
    d0018_6518 = md("2000-2018<br>minus<br>1965-2018"),
    d0018_0017 = md("2000-2018<br>minus<br>2000-2017")
  ) %>%
  fmt_currency(
    columns = -c(uname),
    decimals = 3
  )
tab
gtsave(tab, here::here("results", "compare.png"))




# NONONO simple workbook ----
outfile <- "period_comparisons.xlsx"  

wb <- createWorkbook()
## Add worksheets
snames <- c("1965_2018", "2000_2018", "2000_2017")
speriods <- c("1965-2018", "2000-2018", "2000-2017")
purrr::map(snames, function(s) addWorksheet(wb, s))

for(i in 1:3){
  data <- compare3 %>%
    filter(period==speriods[i])
  writeData(wb, sheet = snames[i], x=data)
}

saveWorkbook(wb, here::here("results", outfile), overwrite = TRUE)


# workbook ----
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




