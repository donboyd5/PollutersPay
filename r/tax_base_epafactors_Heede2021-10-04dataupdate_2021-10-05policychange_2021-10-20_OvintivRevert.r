
# for an existing project, create a github repo and push
# usethis::use_github()

# check gitignore (from bash)
# https://git-scm.com/docs/git-check-ignore
# git check-ignore [<options>] <pathname>.???
# git check-ignore [<options>] --stdin
# git check-ignore --no-index C:/Users/donbo/Documents/R_projects/PollutersPay/
# git check-ignore --no-index --verbose --non-matching C:\Users\donbo\Documents\R_projects\PollutersPay\  



# libraries ----

library(tidyverse)
options(tibble.print_max = 75, tibble.print_min = 75) # if more than 60 rows, print 60 - enough for states
# ggplot2 tibble tidyr readr purrr dplyr stringr forcats
library(scales)

library(readxl)
library(btools)
library(kableExtra)

# devtools::install_version("openxlsx", version = "4.2.3", repos = "http://cran.us.r-project.org")  # 4.2.4 broke my code
# does dev version fix it?
# devtools::install_github("ycphs/openxlsx")
library(openxlsx)
library(gt)

# constants, bankruptcies, low gdp, other ----
# uids (new data) of coal entities
coal_entities <- c(7, 11, 16, 19, 21, 24, 25, 36, 44, 52, 56, 57, 74, 83, 99, 100, 104)

#.. bankruptcies
# coal
# From Harry Stein (with my uid before)
# U.S. coal company bankruptcies from the list shared (1/1/15 to 9/1/21)
# 63 Peabody, filed in 2016
# 7  Arch Coal Company, filed on January 11, 2016
# <NA> Alpha Natural Resources, filed on August 3, 2015 (merged with Contura in 2018)* [22 Contura]
# However, Contura does not meet the controlled group test.
# 17 Cloud Peak, filed on May 10, 2019
# 52 Murray Energy, filed on October, 29, 2019
# 103 Westmoreland Coal Co., filed on October 9, 2018
# 7, 17, 52, 63, 103

# *Alpha Natural Resources filed for bankruptcy, and some of its creditors
# formed a separate company called Contura, which later merged with Alpha. This
# may mean that Contura does not qualify for the bankruptcy exemption in our
# bill, since that only applies to a controlled group of companies if each
# member of the group has filed for bankruptcy (e.g. a large profitable oil
# company is not exempt solely because they happened to buy a bankrupt company).
# This may also be relevant if there are other cases where a company filed for
# bankruptcy and was subsequently acquired.

# U.S. coal companies on list that do not seem to have bankruptcy filings (1/1/15 to 9/1/21)
# CONSOL Energy
# North American Coal Corporation
# Alliance Resource Partners
# Vistra (emerged from the 2014 bankruptcy of TXU)
# Kiewit Mining Group

coal_bankrupt <- c(7, 17, 52, 63, 103)

# other
noncoal_bankrupt <- 14  # Chesapeake

#.. low gdp <= $10k per person in 2020 per World Bank, according to Raj
# 27 Ecopetrol, Colombia
# 48 Libya National Oil Corp., Libya
# 68 Petroleo Brasileiro (Petrobras), Brazil
# 69 Petroleos de Venezuela, Venezuela
# 70 Petroleos Mexicanos (Pemex), Mexico
# 90 Sonangol, Angola
low_gdp_ong <- c(27, 48, 68, 69, 70, 90)
low_gdp_coal <- c(19, 88)
low_gdp <- c(low_gdp_ong, low_gdp_coal)

#.. years
# years <- 1965:2018
years <- 2000:2018


# get new data with EPA factors EVERYTHING IS JUST 2000-2018 ----
# load(file=here::here("data", "constants.rdata"), verbose=TRUE)  # overwrite maindata constants
# load(file=here::here("data", "pollpay_data.Rdata"), verbose=TRUE)

load(file=here::here("data", "epa_base.Rdata"), verbose=TRUE)
sipro <- readRDS(paste0(r"(C:\Users\donbo\Documents\R_projects\PollutersPay\data\)", "siprodata.rds"))
xwalk <- readRDS(here::here("data", "xwalk.rds"))

xwalk %>% filter(uid %in% coal_bankrupt) %>% select(table_name)

gff  # global fossil fuels


# get the base ----
length(unique(epa_ongc$uname))  # 91

ongc_wide <- epa_ongc %>%
  select(-c(quantity, units)) %>%
  pivot_wider(names_from = fuel, values_from = mtco2, values_fill = 0) %>%
  mutate(ongc=oil + gas + coal)


# get nexus status ----
fn <- "top108(21).xlsx"
xldata1 <- read_excel(here::here("data_raw", fn), sheet="receipts_table", range="A3:AV111", col_types = "text")
glimpse(xldata1)

# put proper uid on the file
xldata <- xldata1 %>%
  select(uid_old=uid, uniname, ptype, symbol, fd, owner_notes, country, market_cap, profile, reachable) %>%  # use ptype from here
  mutate(uid_old=as.integer(uid_old),
         market_cap=as.numeric(market_cap)) %>%
  left_join(xwalk %>% select(uid, uname, table_name, uid_old, uname_old), by = "uid_old") %>%
  select(-c(uid_old, uniname, uname_old))  # after verifying that all is good
glimpse(xldata)

nexus <- xldata %>%
  select(uid, uname, table_name, ptype, fd, profile, reachable) %>%
  mutate(nexus=ifelse(str_detect(reachable, coll("yes")), TRUE, FALSE))
summary(nexus)  # this has all 108 entities

nexus %>%
  filter(ptype=="SOE", nexus) %>%
  arrange(table_name)


# merge nexus and determine elements that go into potential payor status ----

# CAUTION: TO REPRODUCE "PollutersPay_2000-2018_th1000_ex1000_keeplowgdp_v2.xlsx" -----
#..  We need to classify Ovintiv as foreign (because predecessor EnCana was Canadian but Ovintiv is USA) ----
# ovintiv_revert <- FALSE
ovintiv_revert <- TRUE  # use this to reproduce the table with Ovintiv as foreign
# END CAUTION ----

base <- ongc_wide %>%
  left_join(nexus %>% 
              select(uid, uname, ptype, fd, nexus), by = c("uid", "uname")) %>%
  # REVERT Ovintiv if needed ----
  mutate(fd=ifelse(ovintiv_revert & str_detect(table_name, "Ovintiv"), "foreign", fd)) %>%
  # END REVERT
  mutate(bankrupt=uid %in% c(coal_bankrupt, noncoal_bankrupt),
         zeroemit=ongc==0,
         lowgdp=uid %in% low_gdp) %>%
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
  arrange(table_name)

summary(base)
base %>%
  filter(str_detect(table_name, "India"))

base %>%
  filter(nexus & !bankrupt & oil==0 & gas==0 & coal>0 & fd=="domestic" & ptype=="IOC") %>%
  select(uid, table_name, oil, gas, coal, ongc)


base %>%
  filter(nexus & !bankrupt & oil==0 & gas==0 & coal>0 & fd=="foreign" & ptype=="IOC") %>%
  select(uid, table_name, oil, gas, coal, ongc)

# get counts of entities lost and kept
base %>%
  mutate(toolow=ongc < 500) %>%
  summarise(n=n(),
            bankrupt_loss=sum(bankrupt),
            non_bankrupt=sum(!bankrupt),
            nexus_loss=sum(!nexus & !bankrupt),
            potential=sum(nexus & !bankrupt),
            threshold_loss=sum(nexus & !bankrupt & toolow),
            potential_big=sum(nexus & !bankrupt & !toolow),
            gdp_loss=sum(lowgdp & nexus & !bankrupt & !toolow),
            gdplimited=sum(!lowgdp & nexus & !bankrupt & !toolow))


# taxbase function ----
get_taxbase <- function(base, droplowgdp, threshold, exclusion, totannual=50){
  base %>%
    mutate(threshold=!!threshold,
           exclusion=!!exclusion,
           droplowgdp=!!droplowgdp,
           potential_payor=nexus & !bankrupt & !(lowgdp & droplowgdp),
           payor=potential_payor & (ongc >= threshold),
           taxbase=pmax(ongc - exclusion, 0, na.rm=TRUE) * payor,
           pct_taxbase=taxbase / sum(taxbase),
           tax_annual=pct_taxbase * !!totannual)
}


# workbook prep for tax ----
source(here::here("r", "write_wb_v4.r"))

basename <- "PollutersPay_2000-2018"
suffix <- "_v2.xlsx"
outdir <- here::here("results")

getfpath <- function(basename, thereshold, exclusion, droplowgdp, suffix, outdir){
  drop <- ifelse(droplowgdp, "_droplowgdp", "_keeplowgdp")
  paste0(outdir, "/", basename, "_th", threshold, "_ex", exclusion, drop, suffix)
}


# create workbooks for tax CAUTION: look at ovintiv_revert ----
# no threshold exclusion, keep low gdp
threshold <- 0
exclusion <- 0
droplowgdp <- FALSE
taxbase <- get_taxbase(base, droplowgdp, threshold, exclusion, totannual=50) %>% filter(payor)
taxbase %>% 
  filter(coal > 0, fd=="domestic", !uid %in% c(15, 35)) %>%
  select(uid, table_name, oil, gas, coal, ongc)
outpath <- getfpath(basename, thereshold, exclusion, droplowgdp, suffix, outdir)
outpath
makewb(taxbase, outpath)
openXL(outpath)

# no threshold exclusion, drop low gdp
threshold <- 0
exclusion <- 0
droplowgdp <- TRUE
taxbase <- get_taxbase(base, droplowgdp, threshold, exclusion, totannual=50) %>% filter(payor)
outpath <- getfpath(basename, thereshold, exclusion, droplowgdp, suffix, outdir)
outpath
makewb(taxbase, outpath)
openXL(outpath)


# 500 MtCO2 threshold exclusion, keep low gdp
threshold <- 500
exclusion <- 500
droplowgdp <- FALSE
taxbase <- get_taxbase(base, droplowgdp, threshold, exclusion, totannual=50) %>% filter(payor)
outpath <- getfpath(basename, thereshold, exclusion, droplowgdp, suffix, outdir)
outpath
makewb(taxbase, outpath)
openXL(outpath)

# 500 MtCO2 threshold exclusion, drop low gdp
threshold <- 500
exclusion <- 500
droplowgdp <- TRUE
taxbase <- get_taxbase(base, droplowgdp, threshold, exclusion, totannual=50) %>% filter(payor)
outpath <- getfpath(basename, thereshold, exclusion, droplowgdp, suffix, outdir)
outpath
makewb(taxbase, outpath)
openXL(outpath)

# 10/7/2021 1000 MtCO2 threshold -- 10/20/2021 this is the latest proposal ----
threshold <- 1000
exclusion <- 1000
droplowgdp <- FALSE
taxbase <- get_taxbase(base, droplowgdp, threshold, exclusion, totannual=50) %>% filter(payor)
outpath <- getfpath(basename, thereshold, exclusion, droplowgdp, suffix, outdir)
outpath
makewb(taxbase, outpath)
openXL(outpath)


# coal companies ----
taxbase %>%
  filter(uid %in% coal_entities) %>%
  select(uid, table_name, oil, gas, coal, ongc, ptype, fd, payor:tax_annual)



# tax as % of market cap ----
threshold <- 500
exclusion <- 500
droplowgdp <- FALSE
taxbase <- get_taxbase(base, droplowgdp, threshold, exclusion, totannual=50) %>% filter(payor)

glimpse(sipro)
mcap <- taxbase %>%
  left_join(xldata %>% select(uid, ticker=symbol, owner_notes, market_cap), by = "uid") %>%
  left_join(sipro %>% 
              select(ticker, company, mktcap_q1) %>%
              mutate(mktcap_q1=mktcap_q1 / 1000),  # convert from millions to billions
            by = "ticker") %>%
  mutate(tax_10year=10 * tax_annual,
         mktcap=mktcap_q1,
         mktcap=ifelse(group=="SOE", NA_real_, mktcap), # drop data for SOEs as it may be suspect
         # mktcap=ifelse(is.na(mktcap_q1) & !is.na(market_cap), market_cap, mktcap_q1),
         fnote=ifelse(is.na(mktcap_q1) & !is.na(market_cap), "mkt cap from web", ""),
         pct=tax_10year / mktcap * 100)

mcap %>% 
  select(group, uid, table_name, ticker, owner_notes) %>%
  filter(!is.na(owner_notes)) %>%
  arrange(group, table_name)

mcap %>% 
  select(uid, table_name, group, ticker, tax_annual, company, mktcap) %>%
  arrange(group, desc(tax_annual))

mcap %>% 
  select(uid, table_name, group, ticker, coal, ongc, tax_10year, mktcap, pct, fnote) %>%
  arrange(group, desc(pct)) %>%
  filter(group=="dIOC") %>%
  # filter(group=="fIOC") %>%
  # filter(group=="SOE") %>%
  mutate(ticker=ifelse(is.na(ticker), "", ticker)) %>%
  select(group, table_name, ticker, coal, ongc, tax_10year, mktcap, pct, fnote) %>%
  kbl(format="rst", digits=c(0, 0, 0, 0, 0, 2, 2, 2, 0), format.args = list(big.mark=","))
  

# %>%
#   gt() %>%
#   tab_header(
#     title = md(paste0("**", "Potential tax as % of market capitalization", "**"))
#     ) %>%
#   cols_label(coal=md("**Coal MtCO2**"),
#              ongc=md("**Total MtCO2**")) %>%
#   cols_width(
#     name ~ pct(20),
#     value ~ pct(80)
#   ) %>%
#   tab_style(
#     style = list(
#       cell_text(weight = "bold")
#     ),
#     locations = cells_body(
#       columns = name)
#   ) %>%
#   cols_align(align = "left",
#              columns = everything()
#   ) %>%
#   fmt_markdown(columns = value) %>%
#   fmt_missing(columns = value,
#               missing_text="")

sipro %>%
  filter(str_detect(company, "Shell")) %>%
  write_csv(here::here("ignore", "shell.csv"))

sipro %>%
  filter(str_detect(company, "Kiewit")) %>%
  write_csv(here::here("ignore", "kiewit.csv"))




