
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

# Rick's email 2021-10-04 ----
# See attached (slight revision, having accounted for Shell acquiring BG, and
# Repsol acquiring Talisman). Don already has these numbers, working from my
# list from Saturday. But I've wanted to verify with revision on my end. XLS and
# PDF attached FYI.


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
# coal_entities <- c(7, 11, 16, 19, 21, 24, 25, 36, 44, 52, 56, 57, 74, 83, 99, 100, 104)
# bankrupt <- 14  # Chesapeake

# years <- 1965:2018
years <- 2000:2018


# get Rick's new data ----
# fn <- "NewCalcs EPA EFs Oct21.xlsx"
# fn <- "NewCalcs EPA EFs 4Oct21.xlsx"
fn <- "NewCalcs EPA EFs 4Oct21(updated).xlsx"

# read each major block

#.. global fossil fuel emissions ----
gff1 <- read_excel(here::here("data_raw", fn), sheet="Oil Gas Coal CO2 2000-2018",
                  range="AD90:AF93", col_types = "text", col_names=c("gvalue", "junk", "fuel"))
gff1
gff <- gff1 %>%
  mutate(fuel=ifelse(fuel=="Sum FF", "total", str_to_lower(fuel)),
         gvalue=as.numeric(gvalue),
         fyear=2000, lyear=2018) %>%
  select(fuel, gvalue, fyear, lyear)
gff


#.. oil EPA factors, Table 4 ----
oil1 <- read_excel(here::here("data_raw", fn), sheet="Oil Gas Coal CO2 2000-2018",
                      range="AD14:AH82", col_types = "text", 
                      col_names=c("quantity", "j1", "mtco2", "j2", "pname"))
oil1
oil <- oil1 %>%
  mutate(quantity=as.numeric(quantity),
         mtco2=as.numeric(mtco2),
         fyear=2000, lyear=2018,
         fuel="oil",
         units="mb") %>%
  filter(!is.na(mtco2)) %>%
  select(pname, fuel, quantity, units, mtco2, fyear, lyear)
oil


#.. gas EPA factors, Table 5 ----
gas1 <- read_excel(here::here("data_raw", fn), sheet="Oil Gas Coal CO2 2000-2018",
                   range="AL14:AP84", col_types = "text", 
                   col_names=c("quantity", "j1", "mtco2", "j2", "pname"))
gas1
gas <- gas1 %>%
  mutate(quantity=as.numeric(quantity),
         mtco2=as.numeric(mtco2),
         fyear=2000, lyear=2018,
         fuel="gas",
         units="bcf") %>%
  filter(!is.na(mtco2)) %>%
  select(pname, fuel, quantity, units, mtco2, fyear, lyear)
gas


#.. coal EPA factors, Table 6 ----
coal1 <- read_excel(here::here("data_raw", fn), sheet="Oil Gas Coal CO2 2000-2018",
                    range="AT13:AZ39", col_types = "text", 
                    col_names=c("quantity", "j1", "mtco2", "j2", "pct", "j3",  "pname"))
coal1

coal <- coal1 %>%
  mutate(quantity=as.numeric(quantity),
         mtco2=as.numeric(mtco2),
         fyear=2000, lyear=2018,
         fuel="coal",
         units="mt") %>%
  filter(!is.na(mtco2)) %>%
  select(pname, fuel, quantity, units, mtco2, fyear, lyear)
coal

#.. oil natural gas, coal, EPA factors ----
ongc1 <- bind_rows(oil, gas, coal)
glimpse(ongc1)
ongc1

ongc1 %>%
  group_by(fuel) %>%
  summarise(mtco2=sum(mtco2))

# clean and conform company names ----
# merge with unames ----
xwalk <- readRDS(here::here("data", "xwalk.rds"))

ongc2 <- ongc1 %>%
  mutate(uname=case_when(str_detect(pname, "ADNOC") ~ "Abu Dhabi, United Arab Emirates",
                         str_detect(pname, "BP Coal, UK") ~ "BP, UK",
                         str_detect(pname, "Chesapeake") ~ "Chesapeake Energy, USA",
                         str_detect(pname, "Chevron Mining / Pittsburgh & Midway") ~ "Chevron, USA",
                         str_detect(pname, "CNOOC") ~ "CNOOC (China National Offshore Oil Co.), PR China (acq Nexen Jan2013)",
                         str_detect(pname, "Contura") ~ "Contura (AlphaNR, Massey), USA",
                         str_detect(pname, "Exxon Mobil") ~ "ExxonMobil, USA",
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
                         str_detect(pname, "Wintershall") ~ "Wintershall, Germany",
                         str_detect(pname, "Yukos") ~ "Yukos, Russia",
                         # str_detect(pname, "") ~ "",
                         # str_detect(pname, "") ~ "",
                         # str_detect(pname, "") ~ "",
                         TRUE ~ pname))

check <- ongc2 %>%
  left_join(xwalk %>% select(uid, uname), by="uname")

check %>% filter(is.na(uid))

# good now we are ready for final
epa_ongc <- ongc2 %>%
  left_join(xwalk %>% select(uid, uname, table_name), by="uname") %>%
  select(uid, uname, table_name, fyear, lyear, fuel, mtco2, quantity, units)

epa_ongc %>%
  group_by(fuel) %>%
  summarise(n=n(), mtco2=sum(mtco2))

save(gff, epa_ongc, file=here::here("data", "epa_base.rdata"))


# old ---


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


# epaongc_all <- good %>%
#   mutate(uname=case_when(str_detect(pname, "ADNOC") ~ "Abu Dhabi, United Arab Emirates",
#                          str_detect(pname, "Chesapeake") ~ "Chesapeake Energy, USA",
#                          str_detect(pname, "CNOOC") ~ "CNOOC (China National Offshore Oil Co.), PR China (acq Nexen Jan2013)",
#                          str_detect(pname, "Contura") ~ "Contura (AlphaNR, Massey), USA",
#                          str_detect(pname, "Gazprom") ~ "Gazprom, Russia",
#                          str_detect(pname, "Lukoil") ~ "Lukoil, Russia",
#                          str_detect(pname, "National Iranian") ~ "National Iranian Oil Co., Iran",
#                          str_detect(pname, "North American Coal") ~ "North American Coal, USA",
#                          str_detect(pname, "Obsidian") ~ "Obsidian, Canada",
#                          str_detect(pname, "Oil and Natural Gas Corporation, India") ~ "Oil and Gas Corp., India",
#                          str_detect(pname, "PEMEX") ~ "Petroleos Mexicanos (Pemex), Mexico",
#                          str_detect(pname, "Repsol") ~ "Repsol, Spain (acq Talisman May2015)",
#                          str_detect(pname, "Talisman") ~ "Repsol, Spain (acq Talisman May2015)",
#                          str_detect(pname, "Rio Tinto") ~ "Rio Tinto, UK",
#                          str_detect(pname, "Royal Dutch Shell") ~ "Royal Dutch Shell, Netherlands (acq BG Feb16)",
#                          str_detect(pname, "BG Group") ~ "Royal Dutch Shell, Netherlands (acq BG Feb16)", # merger
#                          str_detect(pname, "TotalEnergies") ~ "Total, France",
#                          str_detect(pname, "Vistra") ~ "Vistra Luminant, USA",
#                          str_detect(pname, "Westmoreland") ~ "Westmoreland Mining, USA",
#                          str_detect(pname, "Yukos") ~ "Yukos, Russia",
#                          TRUE ~ pname)) %>%
#   left_join(xwalk %>% select(uid, uname), by="uname") %>%
#   arrange(uname, uid, pname)
# 
# epa_ongc <- epaongc_all %>%
#   select(-pname) %>%
#   group_by(uid, uname, fyear, lyear) %>%
#   summarise(across(c(oilgas, coal, ogc), sum), .groups="drop")
# # saveRDS(epa_ongc, here::here("data", "epa_ongc.rds"))
# 
# save(gff, epa_ongc, file=here::here("data", "epa_base.rdata"))

