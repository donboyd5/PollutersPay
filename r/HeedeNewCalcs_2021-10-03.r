
# Rick's email 2021-10-02 ----

# Hi All - I completed all my new calculations that quantify emissions on the
# basis of production of crude oil & NGL, natural gas, and coal from database of
# production 2000-2018.  Pls consider this preliminary, based on our new
# inventory protocol. I am reasonably confident in the results, but it would be
# preferable to have more review time, verifying calculations, and having the
# model and results peer reviewed.

# Don: let me know if you spot anything peculiar.
 
# Then calculated emissions using EPA Emission Factors that we discussed
# earlier. Then calculated the percentage contribution of each assessable person
# of global fossil fuel combustion (Global Carbon Project / CDIAC data
# 2000-2018).
 
# Three separate lists are shown in the attached file: Oil &B Gas only, Coal
# only, and combined oil, gas, and coal - all ranked by attributed CO2
# emissions, and all shown in percent of global fossil fuel emissions.


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


# get Rick's new data ----
fn <- "NewCalcs EPA EFs Oct21.xlsx"

# read each major block

#.. global fossil fuel emissions ----
# R112:T115
gff1 <- read_excel(here::here("data_raw", fn), sheet="Oil Gas Coal CO2 2000-2018",
                  range="R112:T115", col_types = "text", col_names=c("gvalue", "junk", "fuel"))
gff1
gff <- gff1 %>%
  mutate(fuel=ifelse(fuel=="Sum FF", "total", str_to_lower(fuel)),
         gvalue=as.numeric(gvalue),
         fyear=2000, lyear=2018) %>%
  select(fuel, gvalue, fyear, lyear)
gff

#.. oilgas ----
oilgas1 <- read_excel(here::here("data_raw", fn), sheet="Oil Gas Coal CO2 2000-2018",
                     range="B13:F87", col_types = "text", 
                     col_names=c("value", "j1", "pct", "j2", "pname"))
oilgas1
oilgas <- oilgas1 %>%
  mutate(value=as.numeric(value),
    fyear=2000, lyear=2018,
    fuel="oilgas") %>%
  filter(!is.na(value)) %>%
  select(pname, fuel, value, fyear, lyear)
oilgas

#.. coal ----
coal1 <- read_excel(here::here("data_raw", fn), sheet="Oil Gas Coal CO2 2000-2018",
                      range="I13:M38", col_types = "text", 
                      col_names=c("value", "j1", "pct", "j2", "pname"))
coal1
coal <- coal1 %>%
  mutate(value=as.numeric(value),
         fyear=2000, lyear=2018,
         fuel="coal") %>%
  filter(!is.na(value)) %>%
  select(pname, fuel, value, fyear, lyear)
coal

#.. combined ----
combo1 <- read_excel(here::here("data_raw", fn), sheet="Oil Gas Coal CO2 2000-2018",
                    range="P13:X105", col_types = "text", 
                    col_names=c("oilgas", "coal", "sum1", "j1", "sum2", "j2", "pct", "j3", "pname"))
combo1

combo2 <- combo1 %>%
  mutate(across(c(oilgas, coal, sum1, sum2), as.numeric),
         fyear=2000, lyear=2018) %>%
  filter(!is.na(pname)) %>%
  select(pname, oilgas, coal, sum1, sum2, fyear, lyear)
combo2

# check sums
check <- combo2 %>%
  mutate(bad1 = sum1 != sum2,
         bad2 = (naz(oilgas) + naz(coal)) != sum1) %>%
  filter(bad1 | bad2)
check  # good, no problem

combo3 <- combo2 %>%
  select(pname, oilgas, coal, ogc=sum1, fyear, lyear)

#.. compare across the new files ----
mrg1 <- combo3 %>%
  full_join(oilgas %>% select(pname, oilgas1=value), by="pname") %>%
  full_join(coal %>% select(pname, coal1=value), by="pname") %>%
  arrange(pname) %>%
  # for companies that have coal as well as oilgas, Rick appears to have created a
  # new name for the coal portion - here I create a mergeable name
  mutate(pname2=pname,
         pname2=ifelse(pname=="BP Coal, UK", "BP, UK", pname2),
         pname2=ifelse(pname=="Chevron Mining / Pittsburgh & Midway", "Chevron, USA", pname2),
         pname2=ifelse(pname=="Exxon Mobil", "ExxonMobil, USA", pname2))

# now we can collapse
mrg2 <- mrg1 %>%
  select(-pname) %>%
  rename(pname=pname2) %>%
  group_by(pname) %>%
  summarise(across(everything(), ~sum(.x, na.rm=TRUE)))


# looks good, so check values
mrg2 %>%
  mutate(bad=(oilgas1 != oilgas) |
           (coal1 != coal)) %>%
  filter(bad)
# djb: In combo3, Coal India is counted as oilgas when it should be counted as coal
# pname             oilgas  coal    ogc fyear lyear oilgas1  coal1 bad  
# <chr>              <dbl> <dbl>  <dbl> <dbl> <dbl>   <dbl>  <dbl> <lgl>
#   1 Coal India, India 14949.     0 14949.  2000  2018       0 14949. TRUE 

mrg2 %>%
  summarise(across(-pname, ~sum(.x, na.rm=TRUE)))


sum(mrg1$oilgas, na.rm=TRUE) - sum(mrg1$oilgas1, na.rm=TRUE)
sum(mrg1$coal, na.rm=TRUE) - sum(mrg1$coal1, na.rm=TRUE)


# create a version that is corrected as far as I can tell ----

good <- mrg2 %>%
  select(pname, oilgas=oilgas1, coal=coal1, ogc, fyear, lyear)

good %>%
  mutate(diff=ogc - oilgas - coal) %>%
  summarise(across(-pname, ~sum(.x, na.rm=TRUE)))

good %>% filter(str_detect(pname, "India"))

# merge with unames ----
xwalk <- readRDS(here::here("data", "xwalk.rds"))

good2 <- good %>%
  mutate(uname=case_when(str_detect(pname, "ADNOC") ~ "Abu Dhabi, United Arab Emirates",
                         str_detect(pname, "Chesapeake") ~ "Chesapeake Energy, USA",
                         str_detect(pname, "CNOOC") ~ "CNOOC (China National Offshore Oil Co.), PR China (acq Nexen Jan2013)",
                         str_detect(pname, "Contura") ~ "Contura (AlphaNR, Massey), USA",
                         str_detect(pname, "Gazprom") ~ "Gazprom, Russia",
                         str_detect(pname, "Lukoil") ~ "Lukoil, Russia",
                         str_detect(pname, "National Iranian") ~ "National Iranian Oil Co., Iran",
                         str_detect(pname, "North American Coal") ~ "North American Coal, USA",
                         str_detect(pname, "Obsidian") ~ "Obsidian, Canada",
                         str_detect(pname, "Oil and Natural Gas Corporation, India") ~ "Oil and Gas Corp., India",
                         str_detect(pname, "PEMEX") ~ "Petroleos Mexicanos (Pemex), Mexico",
                         str_detect(pname, "Repsol") ~ "Repsol, Spain (acq Talisman May2015)",
                         str_detect(pname, "Rio Tinto") ~ "Rio Tinto, UK",
                         str_detect(pname, "Royal Dutch Shell") ~ "Royal Dutch Shell, Netherlands (acq BG Feb16)",
                         str_detect(pname, "TotalEnergies") ~ "Total, France",
                         str_detect(pname, "Vistra") ~ "Vistra Luminant, USA",
                         str_detect(pname, "Westmoreland") ~ "Westmoreland Mining, USA",
                         str_detect(pname, "Yukos") ~ "Yukos, Russia",
                         # str_detect(pname, "") ~ "",
                         # str_detect(pname, "") ~ "",
                         TRUE ~ pname))

good3 <- xwalk %>%
  select(uid, uname) %>%
  full_join(good2 %>% select(pname, uname), 
            by = "uname")

good3 %>% filter(is.na(uid))


# mystery names ----
# BG Group and Royal Dutch Shell
# BG Group plc was a British multinational oil and gas company headquartered in
# Reading, United Kingdom.[3][4] On 8 April 2015, Royal Dutch Shell announced
# that it had reached an agreement to acquire BG Group for $70 billion, subject
# to regulatory and shareholder agreement. The sale was completed on 15 February
# 2016. Prior to the takeover, BG Group was listed on the London Stock Exchange
# (BG.L) and was a constituent of the FTSE 100 Index. In the 2015 Forbes Global
# 2000, BG Group was ranked as the 583rd largest public company in the world.[5]

# Talisman and Repsol, per Wikipedia
# Talisman Energy Inc. was a Canadian multinational oil and gas exploration and
# production company headquartered in Calgary, Alberta. It was one of Canada's
# largest independent oil and gas companies. Originally formed from the Canadian
# assets of BP Canada Ltd. it grew and operated globally, with operations in
# Canada (B.C., Alberta, Ontario, Saskatchewan, Quebec) and the United States of
# America (Pennsylvania, New York, Texas ) in North America; Colombia, South
# America; Algeria in North Africa; United Kingdom and Norway in Europe;
# Indonesia, Malaysia, Vietnam, Papua New Guinea, East Timor and Australia in
# the Far East; and Kurdistan in the Middle East. Talisman Energy has also built
# the offshore Beatrice Wind Farm[2] in the North Sea off the coast of Scotland.
#
# The company was acquired by Repsol in 2015[3] and in January 2016 was renamed
# to Repsol Oil & Gas Canada Inc.[4]


epaongc_all <- good %>%
  mutate(uname=case_when(str_detect(pname, "ADNOC") ~ "Abu Dhabi, United Arab Emirates",
                         str_detect(pname, "Chesapeake") ~ "Chesapeake Energy, USA",
                         str_detect(pname, "CNOOC") ~ "CNOOC (China National Offshore Oil Co.), PR China (acq Nexen Jan2013)",
                         str_detect(pname, "Contura") ~ "Contura (AlphaNR, Massey), USA",
                         str_detect(pname, "Gazprom") ~ "Gazprom, Russia",
                         str_detect(pname, "Lukoil") ~ "Lukoil, Russia",
                         str_detect(pname, "National Iranian") ~ "National Iranian Oil Co., Iran",
                         str_detect(pname, "North American Coal") ~ "North American Coal, USA",
                         str_detect(pname, "Obsidian") ~ "Obsidian, Canada",
                         str_detect(pname, "Oil and Natural Gas Corporation, India") ~ "Oil and Gas Corp., India",
                         str_detect(pname, "PEMEX") ~ "Petroleos Mexicanos (Pemex), Mexico",
                         str_detect(pname, "Repsol") ~ "Repsol, Spain (acq Talisman May2015)",
                         str_detect(pname, "Talisman") ~ "Repsol, Spain (acq Talisman May2015)",
                         str_detect(pname, "Rio Tinto") ~ "Rio Tinto, UK",
                         str_detect(pname, "Royal Dutch Shell") ~ "Royal Dutch Shell, Netherlands (acq BG Feb16)",
                         str_detect(pname, "BG Group") ~ "Royal Dutch Shell, Netherlands (acq BG Feb16)", # merger
                         str_detect(pname, "TotalEnergies") ~ "Total, France",
                         str_detect(pname, "Vistra") ~ "Vistra Luminant, USA",
                         str_detect(pname, "Westmoreland") ~ "Westmoreland Mining, USA",
                         str_detect(pname, "Yukos") ~ "Yukos, Russia",
                         TRUE ~ pname)) %>%
  left_join(xwalk %>% select(uid, uname), by="uname") %>%
  arrange(uname, uid, pname)

epa_ongc <- epaongc_all %>%
  select(-pname) %>%
  group_by(uid, uname, fyear, lyear) %>%
  summarise(across(c(oilgas, coal, ogc), sum), .groups="drop")
# saveRDS(epa_ongc, here::here("data", "epa_ongc.rds"))

save(gff, epa_ongc, file=here::here("data", "epa_base.rdata"))

