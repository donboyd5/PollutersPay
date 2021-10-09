
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

# constants ----
# uids (new data) of coal entities
coal_entities <- c(7, 11, 16, 19, 21, 24, 25, 36, 44, 52, 56, 57, 74, 83, 99, 100, 104)
bankrupt <- 14  # Chesapeake

# years <- 1965:2018
years <- 2000:2018


# get new data with EPA factors EVERYTHING IS JUST 2000-2018 ----
# load(file=here::here("data", "constants.rdata"), verbose=TRUE)  # overwrite maindata constants
# load(file=here::here("data", "pollpay_data.Rdata"), verbose=TRUE)

load(file=here::here("data", "epa_base.Rdata"), verbose=TRUE)
sipro <- readRDS(paste0(r"(C:\Users\donbo\Documents\R_projects\PollutersPay\data\)", "siprodata.rds"))
xwalk <- readRDS(here::here("data", "xwalk.rds"))
xwalk %>% filter(uid %in% coal_entities) %>% select(uname)

gff  # global fossil fuels


# get the base ----
length(unique(epa_ongc$uname))  # 91
dca_base <- epa_ongc %>%
  filter(fuel %in% c("oil", "gas")) %>%
  group_by(uid, uname) %>%
  summarise(dca_base=sum(mtco2, na.rm=TRUE), .groups="drop")
summary(dca_base)  # 71 companies
sum(dca_base$dca_base)

ongc_wide <- epa_ongc %>%
  select(-c(quantity, units)) %>%
  pivot_wider(names_from = fuel, values_from = mtco2, values_fill = 0) %>%
  mutate(dca_base=oil + gas)


# get nexus status ----
fn <- "top108(19).xlsx"
xldata1 <- read_excel(here::here("data_raw", fn), sheet="receipts_table", range="A3:AU111", col_types = "text")
glimpse(xldata1)

# put proper uid on the file
xldata <- xldata1 %>%
  select(uid_old=uid, uniname, table_name, ptype, symbol, fd, country, reachable) %>%  # use ptype from here
  mutate(uid_old=as.integer(uid_old)) %>%
  left_join(xwalk %>% select(uid, uname, uid_old, uname_old), by = "uid_old") %>%
  select(-c(uid_old, uniname, uname_old))  # after verifying that all is good
glimpse(xldata)

nexus <- xldata %>%
  select(uid, uname, ptype, fd, reachable) %>%
  mutate(nexus=ifelse(str_detect(reachable, coll("yes")), TRUE, FALSE))
summary(nexus)  # this has all 108 entities
sum(nexus$nexus)


# merge nexus and determine elements that go into potential payor status ----
dca_total_max <- 4000 # $ billions
dca_total_annual <- 50 # $ billions
topn <- 35

base <- ongc_wide %>%
  left_join(nexus, by = c("uid", "uname")) %>%
  mutate(coalco=uid %in% coal_entities,
         bankrupt= uid %in% !!bankrupt,
         zeroemit= dca_base==0,
         potential_payor=nexus & !coalco & !bankrupt & !zeroemit) %>%
  group_by(potential_payor) %>%
  mutate(rank=rank(desc(dca_base))) %>%
  ungroup %>%
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
  arrange(uname)

summary(base)


base %>% 
  filter(coalco | zeroemit | coal >= 100) %>% 
  select(uid, uname, oilgas=dca_base, coal)

base %>% 
  filter(coalco | zeroemit) %>% 
  select(uid, uname, oilgas=dca_base, coal)

base %>% 
  filter(!(coalco | zeroemit)) %>% 
  filter(bankrupt) %>%
  select(uid, uname, oilgas=dca_base, coal)

base %>% 
  filter(!(coalco | zeroemit | bankrupt)) %>% 
  filter(!nexus) %>%
  select(uid, uname, oilgas=dca_base, coal)

base %>% 
  filter(!(coalco | zeroemit | bankrupt)) %>%
  filter(nexus) %>%
  select(uid, uname, oilgas=dca_base, coal)




getbase <- function(base, topn, totannual){
  base %>%
    mutate(payor=potential_payor & (rank <= !!topn),
           taxbase=dca_base * payor,
           pct_taxbase=taxbase / sum(taxbase),
           dca_annual=pct_taxbase * !!totannual)
}


# merge nexus and determine payor status ----
dca_total_max <- 4000 # $ billions
dca_total_annual <- 50 # $ billions
topn <- 30
# write workbook for top-n-based taxes ----
source(here::here("r", "write_wb_v3.r"))

basename <- "PollutersPay_2000-2018_top"
suffix <- "_v1.xlsx"
outdir <- here::here("results")


# top 30
topn <- 30
taxbase <- getbase(base, topn, totannual=50) %>% filter(payor)
outpath <- paste0(outdir, "/", basename, topn, suffix)
makewb(taxbase, outpath)
openXL(outpath)

# top 35
topn <- 35
taxbase <- getbase(base, topn, totannual=50) %>% filter(payor)
outpath <- paste0(outdir, "/", basename, topn, suffix)
makewb(taxbase, outpath)
openXL(outpath)

# top 40
topn <- 40
taxbase <- getbase(base, topn, totannual=50) %>% filter(payor)
outpath <- paste0(outdir, "/", basename, topn, suffix)
makewb(taxbase, outpath)
openXL(outpath)

# top 48
topn <- 48
taxbase <- getbase(base, topn, totannual=50) %>% filter(payor)
outpath <- paste0(outdir, "/", basename, topn, suffix)
makewb(taxbase, outpath)
openXL(outpath)

