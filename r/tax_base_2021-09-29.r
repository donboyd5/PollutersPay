
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


# get new good data ----
load(file=here::here("data", "constants.rdata"), verbose=TRUE)  # overwrite maindata constants
load(file=here::here("data", "pollpay_data.Rdata"), verbose=TRUE)
xwalk <- readRDS(here::here("data", "xwalk.rds"))

# get relevant global totals: oil, NG, and coal -related emissions, NOT cement or methane ----
count(global_totals, fuel)
global_totals %>%
  filter(year==2018) %>%
  select(-source)

gtots <- global_totals %>%
  filter(year %in% years) %>%
  group_by(fuel) %>%
  summarise(fyear=years[1],
            lyear=years[length(years)],
            value=sum(value, na.rm=TRUE)) %>%
  pivot_wider(names_from = fuel) %>%
  mutate(oilngflare=total - methane - cement - coal,
         total_xmeth_xcem=total - methane - cement) %>%  # global excludes methane and cement
  select(fyear, lyear, oilngflare, coal, total_xmeth_xcem, cement, total_xmeth, methane, total)
gtots


# now get the firm-level data ----
count(eachevery, etype, ename)


entities <- eachevery %>%
  filter(year %in% years, etype != "MtCH4") %>%
  mutate(etype=str_to_lower(etype)) %>%
  group_by(uid, uname, etype) %>%
  summarise(fyear=years[1],
            lyear=years[length(years)],
            value=sum(value, na.rm=TRUE), .groups="drop") %>%
  pivot_wider(names_from = etype) %>%
  mutate(methane=mtco2e - mtco2)
sum(entities$methane)
sum(entities$methane) / sum(entities$mtco2e)


# get the base: remove methane, cement, coal ----
dca_base <- entities %>%
  left_join(coalcement_17502018 %>% select(uid, contains("share")),
            by="uid") %>%
  mutate(
    # shares based on 1751-2018
    coal=mtco2e * coal_share,
    cement=mtco2e * cement_share,
    tot_xcc=mtco2e - coal - cement,  # removes the methane portion that is in coal so we have to reduce methane
    methane_share = methane / mtco2e,
    meth_noncoal=tot_xcc * methane_share,
    dca_base=mtco2e - coal - cement - meth_noncoal)  # this should guarantee no negative numbers
summary(dca_base)
sum(dca_base$dca_base)


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


# merge nexus and determine payor status ----
# gtot_xcement=gtots$xcement, pct_emissions=dca_base / gtot_xcement, # really a proportion, not a percent  # used to be global6518_xcement
# total_xmeth_xcem
dca_total_max <- 4000 # $ billions
dca_total_annual <- 50 # $ billions
# threshold_original <- 0.0005
# threshold_alt <- 0.0015


# write workbooks ----
source(here::here("r", "write_wb.r"))

sufx <- "_v3.xlsx"


#.. 0.05%, no exclusion ----
fbase <- "PollutersPay_2000-2018_0005_noexclude"
fname <- paste0(fbase, sufx)
outpath <- here::here("results", fname)
makewb(tottax=dca_total_annual, dca_base, nexus, years=2000:2018, threshold=0.0005, exclusion="none",
       coal_entities, bankrupt, gtots, outpath)
openXL(outpath)


#.. 0.15%, no exclusion ----
fbase <- "PollutersPay_2000-2018_0015_noexclude"
fname <- paste0(fbase, sufx)
outpath <- here::here("results", fname)
makewb(tottax=dca_total_annual, dca_base, nexus, years=2000:2018, threshold=0.0015, exclusion="none",
       coal_entities, bankrupt, gtots, outpath)
openXL(outpath)


#.. 0.15%, exclude the threshold amount ----
fbase <- "PollutersPay_2000-2018_0015_exclusion"
fname <- paste0(fbase, sufx)
outpath <- here::here("results", fname)
makewb(tottax=dca_total_annual, dca_base, nexus, years=2000:2018, threshold=0.0015, exclusion="threshold",
       coal_entities, bankrupt, gtots, outpath)
openXL(outpath)


# scratch ----
global_totals %>%
  group_by(fuel) %>%
  summarise(value=sum(value, na.rm=TRUE)) 
gtots


# OLD ----
# do the calcs using a proportion ----
threshold_proportion <- threshold_alt
# threshold_proportion <- threshold_original
exclusion <- "none"
exclusion <- "threshold"


taxbase <- dca_base %>%
  left_join(nexus, by = c("uid", "uname")) %>%
  mutate(coalco=uid %in% coal_entities,
         pct_emissions = dca_base / gtots$total_xmeth_xcem,
         threshold=threshold_proportion,
         sizeok = pct_emissions >= threshold,
         payor= nexus & !coalco & sizeok,
         taxbase_nox=dca_base * payor,
         taxbase_ex=(dca_base - threshold * gtots$total_xmeth_xcem) * payor,
         exclude=exclusion,
         taxbase=ifelse(exclude=="none", taxbase_nox, taxbase_ex),
         pct_taxbase=taxbase / sum(taxbase),
         dca_annual=pct_taxbase * dca_total_annual,
         group=case_when(ptype=="Nation-State" ~ "nation",
                         ptype=="SOE" ~ "SOE",
                         ptype=="IOC" & fd=="foreign" ~ "fIOC",
                         ptype=="IOC" & fd=="domestic" ~ "dIOC",
                         TRUE ~ "ERROR"),
         label=factor(group,
                      levels=c("nation", "SOE", "fIOC", "dIOC"),
                      labels=c("Nation-State Entity",
                               "State-Owned Enterprise",
                               "Foreign Investor-Owned Company",
                               "Domestic Investor-Owned Company")))

# glimpse(taxbase)
# summary(taxbase)
sum(taxbase$dca_annual)


# reporting
#.. aggregate reporting
taxbase %>%
  group_by(group, label) %>%
  summarise(fyear=first(fyear),
            lyear=first(lyear),
            threshold=first(threshold),
            exclude=first(exclude),
            n=n(),
            payor=sum(payor),
            dca_annual=sum(dca_annual),
            .groups="drop") %>%
  mutate(tax_share=dca_annual / sum(dca_annual)) %>%
  select(fyear, lyear, threshold, exclude, group, label, everything())

#.. entity reporting
taxbase %>%
  filter(dca_annual > 0) %>%
  select(fyear, lyear, threshold, exclude, group, label, uid, uname, pct_emissions, pct_taxbase, dca_annual) %>%
  arrange(group, desc(pct_emissions))
