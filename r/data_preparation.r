

# About the data ----

#.. notes ----

# SumEach&Every1850-2018 May21.xlsx  # received 2021-08-27
# SumEvery	 all entities 1854-2018, by gas co2 and ch4, and global ff and cement	Rick Heede email'
# SumFiles	 Sum Ranking tab has sums from 1751-2018 	Boyd check, ok'd by Rick Heede

#.. Notes from Rick Heede ----
# Rick Heede May 14 6:15 pm
# I can confrirm that with yet another summary sheet, attached.
# 
# It lists all entities 1854-2018, by gas co2 and ch4, and global ff and cement.
# 
# For your purposes, see column GW for each entity, and cell GW580 for global sum 1965-2018.
# 
# Most usefully for you: see worksheet "All CMEs 1965-2018" clmns V-X.
# 
# The global sum is slight;y revised from Oct20 update (reflecting Glbal Carbon Poroject slight revisions of FF enmissions), now totaling 1,409,854 MtCO2e.
# 
# PS: I have yet to inciorporate Robbie Andrew's "corrections" on cement emissions. (Andrew, Robbie (2019) Global CO2 emissions from cement production, 1928-2018, Earth Syst. Sci. Data, vol. 11:1675-1710.)


  
#.. Overview ----
# SumEach&Every1850-2018 May21.xlsx"  # received 2021-08-27

# SumFiles Boyd May21(Sent in afternoon).xlsx:
#   CDIAC & CME 1810-2018: plot, pay no attention
#   Sum Oil, Gas, Coal, & Cement: type of fuel, annual, 1790-2018 complete plus a bit more
#   Sum Ranking: by entity, 1965-2018??, by emissions group THIS IS IMPORTANT
#   Data for Chart Top 20 source: ignore
#   Chart Top20 MtCO2e Oct20: ignore


# START ----

# libraries ----
library(tidyverse)
options(tibble.print_max = 110, tibble.print_min = 110) # if more than 60 rows, print 60 - enough for states
# ggplot2 tibble tidyr readr purrr dplyr stringr forcats

library(readxl)
library(btools)
library(kableExtra)

# require(devtools)
# devtools::install_version("openxlsx", version = "4.2.3", repos = "http://cran.us.r-project.org")  # 4.2.4 broke my code
library(openxlsx)
library(gt)

library(phonics)
library(fuzzyjoin)

# av <- available.packages(filters=list())
# av[av[, "Package"] == "kable", ]

# file names ----
seachev <- "SumEach&Every1850-2018 May21.xlsx"  # received 2021-08-27
severy <- "SumEvery Boyd May21.xlsx"
sfiles <- "SumFiles Boyd May21(Sent in afternoon).xlsx"


# constants to save ----

global

# Global MtCO2 & MtCH4 emissions, 1965-2018
# File: "SumEvery Boyd May21.xlsx", sheet: "All CMEs 1965-2018", cell: W10
global6518 <-  1409854  # same number as in Rick Heede's May 14 6:15 pm email


# IPCC values (28Dec12):
ongl_factors <- list()
ongl_factors$flaring <-  15.9446317  # 15.94  # flaring (co2) = pls * flaring_factor / 10^3 (kg CO2/tCO2)
ongl_factors$vented <-  3.8325929 # 3.833 # vented (co2) = pls * vented_factor / 10^3 (kg CO2/tCO2)
ongl_factors$methanech4 <-  1.9235120  # EPA -- 1.924 # methanech4 = pls * methanech4_factor / 10^3 (kg CH4/tCO2)
ongl_factors$methane_conversion <- 28  # methane (co2) = methanch4 * methane_conversion
ongl_factors$methaneco2 <- ongl_factors$methanech4 * ongl_factors$methane_conversion # approx 53.86 # kg CO2e/tCO2
ongl_factors

ng_factors <- list()
ng_factors$flaring <-   1.735564  #   # flaring (co2) = pls * flaring_factor / 10^3 (kg CO2/tCO2)
ng_factors$vented <-  28.533725 # # vented (co2) = pls * vented_factor / 10^3 (kg CO2/tCO2)
ng_factors$methanech4 <-  9.878313  # EPA -- 1.924 # methanech4 = pls * methanech4_factor / 10^3 (kg CH4/tCO2)
ng_factors$methane_conversion <- 28  # methane (co2) = methanch4 * methane_conversion
ng_factors$methaneco2 <- ng_factors$methanech4 * ng_factors$methane_conversion # approx 53.86 # kg CO2e/tCO2
ng_factors$ownuse <-  57.264386
ng_factors


coal_factors <- list()
coal_factors$methanech4 <- 4.0346184
coal_factors$methane_conversion <- 28  # methane (co2) = methanech4 * methane_conversion
coal_factors$methaneco2 <- coal_factors$methanech4 * coal_factors$methane_conversion  #	112.9693154
coal_factors

# pls: production less sequestration
save(global6518, ongl_factors, ng_factors, coal_factors, file=here::here("data", "constants.rdata"))

load(file=here::here("data", "constants.rdata"), verbose=TRUE)


# SumFiles 1751-2018 ----
#.. Notes: ----
#..   "Sum Ranking" has each entity, by fuel and emissions type, sum over all years, NOT broken down by year ----
#..   "Sum Oil, Gas, Coal, & Cement" has each fuel, by year, sum over all majors, NOT broken down by entity or emissions type ----
path <- here::here("data_raw", sfiles)

#.. ranking of all 108----
#....oilngl ----
oilngl_vnames <- paste0("oilngl_", c("production", "flaring", "vented", "methane", "total"))
oilngl <- read_excel(path, sheet="Sum Ranking", range="B18:M127",  # include the totals
                 col_names = FALSE, col_types = "text") %>%
  select(1, 3, 5, 7:8, 11, 12) %>%
  setNames(c("pid", "pname", oilngl_vnames)) %>%
  filter(!is.na(pname)) %>%
  mutate(pid=ifelse(is.na(pid), "999", pid))
glimpse(oilngl)


#.... ng natural gas ----
ng_vnames <- paste0("ng_", c("production", "flaring", "vented", "methane", "own", "total"))
ng <- read_excel(path, sheet="Sum Ranking", range="P18:AA127",  # include the totals
                    col_names = FALSE, col_types = "text") %>%
  select(1, 3, 5, 7:8, 10:12) %>%
  setNames(c("pid", "pname", ng_vnames)) %>%
  filter(!is.na(pname)) %>%
  mutate(pid=ifelse(is.na(pid), "999", pid))
glimpse(ng)

# .... coal and cement ----
cc_vnames <- c("coal_production", "coal_methane", "coal_total", "cement_production")
cc <- read_excel(path, sheet="Sum Ranking", range="AD18:AO127",  # include the totals
                 col_names = FALSE, col_types = "text") %>%
  select(1, 3, 5, 8:9, 12) %>%
  setNames(c("pid", "pname", cc_vnames)) %>%
  filter(!is.na(pname)) %>%
  mutate(pid=ifelse(is.na(pid), "999", pid),
         cement_total=cement_production)
glimpse(cc)


# .... totals ----
tot_vnames <- paste0("total_", c("production", "flaring", "vented", "own", "methane", "total"))
tot <- read_excel(path, sheet="Sum Ranking", range="AS18:BG127",  # include the totals
                 col_names = FALSE, col_types = "text") %>%
  select(1, 3, 5, 9:11, 13, 15) %>%
  setNames(c("pid", "pname", tot_vnames)) %>%
  filter(!is.na(pname)) %>%
  mutate(pid=ifelse(is.na(pid), "999", pid))
glimpse(tot)

#.. combine dfs ----
df <- oilngl %>%
  rename(pname_oil=pname) %>%
  left_join(ng %>% rename(pname_ng=pname), by="pid") %>%
  left_join(cc %>% rename(pname_cc=pname), by="pid") %>%
  left_join(tot %>% rename(pname_tot=pname), by="pid")

#.. save pnames from SumFiles ----
pnames_sumfiles <- df %>%
  select(pid, starts_with("pname")) %>%
  mutate(ngerr=pname_ng != pname_oil,
         ccerr=pname_cc != pname_oil,
         toterr=pname_tot != pname_oil,
         pid=as.integer(pid))
summary(pnames_sumfiles)
saveRDS(pnames_sumfiles, here::here("data_interim", "pnames_sumfiles.rds"))


#.. check sumfiles data ----
df2 <- df %>%
  rename(pname=pname_oil) %>%
  select(-starts_with("pname_")) %>%
  pivot_longer(cols=-c(pid, pname)) %>%
  mutate(pid=as.integer(pid),
         value=as.numeric(value)) %>%
  filter(!is.na(value))

count(df2, name)

df3 <- df2 %>%
  separate(name, into=c("fuel", "emit"))

# check grand sums
count(df3, emit) # note that there is no total for fuelcem or maybe no detail
# 1 flaring      217
# 2 methane      292
# 3 own          189
# 4 production   302
# 5 total        302
# 6 vented       217

count(df3, fuel)
df3 %>%
  # filter(emit=="own") %>%
  mutate(group=ifelse(fuel=="total", "total", "detail")) %>%
  group_by(pid, pname, emit, group) %>%
  summarise(sum=sum(value), .groups="drop") %>%
  pivot_wider(names_from = group, values_from = sum) %>%
  mutate(diff=detail - total) %>%
  filter(diff != 0) %>%
  arrange(desc(abs(diff)))

# emit good: own, flaring, methane, production, vented, total (within precision)

count(df3, emit, fuel)

df3 %>% 
  filter(pid==999, emit=="total", fuel=="tot")

df3 %>% filter(pid==999)


#.. save sumfiles data ----
sumfiles <- df3 %>%
  mutate(fgroup=ifelse(fuel=="total", "total", "detail"),
         egroup=ifelse(emit=="total", "total", "detail"),
         source="1751-2018 sums by entity [SumFiles Boyd May21(Sent in afternoon).xlsx]") %>%
  select(pid, pname, fgroup, fuel, egroup, emit, value, source)

sumfiles %>% filter(fgroup=="total", egroup=="total")
saveRDS(sumfiles, here::here("data_interim", "sumfiles_1751_2018.rds"))


#.. get global totals cement and other totals from SumFiles ----



# global_cement <- cement1 %>%
#   pivot_longer(cols=everything(), names_to = "year") %>%
#   mutate(year=as.integer(str_sub(year, 2, 5)),
#          value=as.numeric(value),
#          source="1751-2018 sums by entity [SumFiles Boyd May21(Sent in afternoon).xlsx]")
# ht(global_cement)


# saveRDS(global_totals, here::here("data_interim", "global_totals.rds"))


# do some checks
# global_cement %>%
#   filter(year %in% 1751:2018) %>%  # the SumRanking years
#   summarise(value=sum(value))  # 39874 
# Sum Ranking has 39,922, so very close
# 39874 / 39922 = 0.9987977

# coalcement <- usumrank_long %>%
#   mutate(vname = case_when(
#     category == "all" & etype == "total" ~ "total",
#     category == "all" & etype == "fmethane" ~ "fmethane",
#     category == "cement" & etype == "total" ~ "cement",
#     category == "coal" & etype == "fmethane" ~ "coal_methane",
#     category == "coal" & etype == "total" ~ "coal",
#     TRUE ~ "drop"
#   )) %>%
#   filter(vname != "drop") %>%
#   select(uid, uniname, ptype, vname, value) %>%
#   pivot_wider(names_from = vname) %>%
#   mutate(across(-c(uniname, ptype), .fns=replace_na, 0),
#          coal_share=coal / total,
#          cement_share=cement / total)
# 
# 
# global_shares <- xcoalcement %>%
#   summarise(dca_base=sum(dca_base), coal=sum(coal), cement=sum(cement), total=sum(total)) %>%
#   mutate(coal_share=coal / total, cement_share=cement / total)
# global_shares
# 
# global6518_xcement <- global6518 * (1 - global_shares$cement_share)




# EachEvery 1850-2018 ----
#.. Notes: ----
#.... "SumEach&Every1850-2018 May21.xlsx", 1850-2018, annual, by entity, NOT by fuel or emission type ----
path <- here::here("data_raw", seachev)
excel_sheets(path)
# "Control Sheet" # not needed
# "Sum each CME"

#.. Notes: ----
#..   "Sum each CME" ----
df <- read_excel(path, sheet = "Sum each CME", range="B11:FV566", col_types = "text")
glimpse(df)

#.. initial cleaning before pivot ----
rnames <- c(c("id", "pename", "petype", "note1", "anno", "note2", "note3", "note4"), names(df)[9:length(names(df))])
# pename holds both polluter name and emission name
# petype holds both polluter type and emission type
rnames
df2 <- df %>%
  setNames(rnames)
count(df2, pename)
count(df2, petype)
count(df2, note1)
count(df2, note2)
count(df2, note3)
count(df2, note4)
options(tibble.print_max = 200, tibble.print_min = 200) # if more than 60 rows, print 60 - enough for states
count(df2, anno)
options(tibble.print_max = 65, tibble.print_min = 65) # if more than 60 rows, print 60 - enough for states
# ok to drop note1, note3, note4

tmp <- df2 %>%
  select(id, petype) %>%
  mutate(petype=if_else(id=="76" & !is.na(id), "Nation-State", petype))

# start
df3 <- df2 %>%
  mutate(id=as.integer(id),
         petype=if_else(id==76 & !is.na(id), "Nation-State", petype),  # fix PTTEP, Thailand
         pname=case_when(petype %in% c("IOC", "Nation-State", "SOE") ~ pename,
                         TRUE ~ NA_character_),
         ename=case_when(petype %in% c("MtCH4", "MtCO2", "MtCO2e") ~ pename,
                         TRUE ~ NA_character_),
         ptype=case_when(petype %in% c("IOC", "Nation-State", "SOE") ~ petype,
                         TRUE ~ NA_character_),
         etype=case_when(petype %in% c("MtCH4", "MtCO2", "MtCO2e") ~ petype,
                         TRUE ~ NA_character_)) %>%
  filter(!is.na(pename)) %>%  # these are the only ones we need -- must do before fill
  fill(c(id, pname, ptype)) %>%
  rename(fuelavail=note2, yearsavail=anno) %>%
  mutate(source="SumEach&Every1850-2018 May21.xlsx") %>%
  select(-c(pename, petype, note1, note3, note4)) %>%
  select(id, pname, ptype, ename, etype, fuelavail, yearsavail, everything())
# at this point we should have 4 records per entity = 4 x 108 = 432
count(df3, id, pname) %>% filter(n != 4)
glimpse(df3)
count(df3, ptype)
count(df3, etype)
count(df3, ename)
count(df3, fuelavail)
count(df3, fuelavail, yearsavail)

# now split into three files
ptypes <- df3 %>%
  select(id_ee=id, pname_ee=pname, ptype, source) %>%
  distinct()
saveRDS(ptypes, file=here::here("data_interim", "ptypes.rds"))

fuels <- df3 %>%
  filter(!is.na(fuelavail)) %>%
  select(id_ee=id, pname_ee=pname, fuelavail, yearsavail, source)
saveRDS(fuels, file=here::here("data_interim", "fuels.rds"))

# continue on with data cleaning
df4 <- df3 %>%
  filter(!is.na(etype)) %>% # we don't need the header rec for each pname
  select(-fuel, -yearsavail)
# we should have 3 x 108 records- 324
count(df4, id, pname) %>% filter(n!=3)
names(df4)

# pivot and complete cleaning
df5 <- df4 %>%
  pivot_longer(-c(id, pname, ptype, ename, etype, source), names_to = "year") %>%
  mutate(year=as.integer(year))

check <- df5 %>% filter(!is.na(value))  # examine the values -- good, none appear to be a non-number

df6 <- df5 %>%
  mutate(value=as.numeric(value)) %>%
  filter(!is.na(value)) %>%
  arrange(id, etype)

# get some sums to compare against Rick Heede's spreadsheet
df6 %>%
  filter(etype=="MtCO2e") %>%
  group_by(year) %>%
  summarise(n=n(), sum=sum(value)) %>%
  ht

df6 %>%
  filter(etype=="MtCO2e") %>%
  summarise(sum=sum(value))  # good, matches: 1,258,891
eachevery <- df6
saveRDS(eachevery, file=here::here("data_interim", "eachevery.rds"))


# quick  check to see how  different years affect the largest
eachevery %>%
  filter(etype=="MtCO2e", year >= 1965, id!=36) %>%
  filter(!str_detect(pname, "coal"), !ptype=="Nation-State") %>%
  group_by(id, pname) %>%
  summarise(val1965plus=sum(value[year >= 1965]), val2000plus=sum(value[year >= 2000]), .groups = "drop") %>%
  mutate(share1965plus=val1965plus / sum(val1965plus),
         share2000plus=val2000plus / sum(val2000plus),
         change=share2000plus - share1965plus) %>%
  arrange(-abs(change))

eachevery %>%
  filter(etype=="MtCO2e", year >= 2000) %>%
  group_by(id, pname) %>%
  summarise(value=sum(value), .groups = "drop") %>%
  arrange(-value)


# Create uniform names to use in reports etc. ----
# fuzzy merge of the files on pname
# 1. create files for matching
# 2. preprocess pname
# 3. do fuzzy full join, find closest
# 4. post processing

pnames_sumfiles <- readRDS(here::here("data_interim", "pnames_sumfiles.rds")) %>% # 109 obs
  rename(id_sf=pid)

pnames_eachevery <- readRDS(here::here("data_interim", "eachevery.rds")) %>%
  select(id_ee=id, pname_ee=pname) %>%
  distinct()  # 108 obs

pnames_sumfiles %>%
  filter(ngerr | ccerr | toterr)

# pid pname_oil                                       pname_ng                  pname_cc                  pname_tot         ngerr ccerr toterr
# <int> <chr>                                       <chr>                     <chr>                     <chr>             <lgl> <lgl> <lgl> 
#   1    79 HeidelbergCement, Germany (acq Italcementi) HeidelbergCement, Germany HeidelbergCement, Germany HeidelbergCement~ TRUE  TRUE  TRUE 

preprocess <- function(pname){
  # fix the ones we know will not match
  # function to create a matching variable for the few that will not match
  # or where we want to give a new name for reporting purposes
  # the pname created here is the name we will want to use as our best name
  case_when(str_detect(pname, "BHP") ~ "BHP Billiton, Australia",
            
            str_detect(pname, "Chesapeake") ~ "Chesapeake Energy, USA",
            
            str_detect(pname, "China, PR") |
              str_detect(pname, "China, Peoples Rep") ~ "China, PR (coal & cement only)",
            
            str_detect(pname, "CNOOC") ~ "CNOOC (China National Offshore Oil Co.), PR China (acq Nexen Jan2013)",
            
            str_detect(pname, "Oil and Gas Corp., India") |
              str_detect(pname, "Oil & Gas Corp., India") ~ "Oil and Gas Corp., India",
            
            str_detect(pname, "Pemex") ~ "Petroleos Mexicanos (Pemex), Mexico",
            str_detect(pname, "Petrobras") ~ "Petroleo Brasileiro (Petrobras), Brazil",
            
            str_detect(pname, "RAG, Germany") |
              str_detect(pname, "Ruhrkohle") ~ "Ruhrkohle AG (RAG), Germany",
            
            TRUE ~ pname)
}

sound_name <- function(pname){
  refinedSoundex(pname, clean=FALSE, maxCodeLen=12)
  # pname
}

df1 <- pnames_sumfiles %>%
  filter(id_sf != 999) %>%  # drop the all-company-totals record
  select(-contains("err")) %>%
  mutate(pname=preprocess(pname_oil),
         sound=sound_name(pname))

df2 <- pnames_eachevery %>%
  mutate(pname=preprocess(pname_ee),
         sound=sound_name(pname))

# df1 %>% filter(str_detect(pname, "CNOOC"))
# df2 %>% filter(str_detect(pname, "CNOOC"))

df <- stringdist_full_join(df1, df2,
                           by="sound",
                           # max_dist=1,
                           method="jw", # good for similar items that might be misspelled
                           distance_col="dist") %>%
  select(id_sf, id_ee, pname.x, pname.y, sound.x, sound.y, dist, everything()) %>%
  arrange(pname_oil, dist)

best <- df %>%
  arrange(pname_oil, dist) %>%
  group_by(pname_oil) %>%
  filter(row_number()==1) %>%
  ungroup

best %>%
  filter(sound.x != sound.y) %>%
  select(pname.x, pname.y, sound.x, sound.y)

best %>%
  filter(pname.x != pname.y) %>%
  select(pname.x, pname.y, sound.x, sound.y)

# best looks good, so post process
best_post <- best %>%
  mutate(uname=pname.x,
         uname=ifelse(uname=="Russian Federation",
                      "Russian Federation (excl. FSU) (coal)",
                      uname),
         uname=ifelse(str_sub(uname, -2, -1)=="US",
                      paste0(uname, "A"),
                      uname),
         uname=ifelse(!str_detect(uname, coll("(coal)")) &
                        str_detect(pname.y, coll("(coal)")),
                      paste0(uname, " (coal)"),
                      uname)) %>%
  arrange(uname) %>%
  mutate(uid=row_number()) %>%
  select(uid, uname, starts_with("pname"), everything())

check <- best_post %>% select(uname, pname_oil, pname_ee)

# slim the file and save
# this file can be matched against others
unames <- best_post %>%
  select(uid, uname, contains("pname_"), contains("id_"))

saveRDS(unames, here::here("data_interim", "unames.rds"))

#.. update unames to create table_name based on Rick Heede's updates 10/5/2021 ----
# Change the name of Husky to Cenovus (name change Jan 2021) 
# Change the name of EnCana to Ovintiv (Jan 2020) 
# Change the name of Total to TotalEnergies (Jul 21).
unames1 <- readRDS(here::here("data_interim", "unames.rds"))
unames <- unames1 %>%
  mutate(table_name=case_when(str_detect(uname, "Husky") ~ "Cenovus, Canada (name change Jan 2021)",
                              str_detect(uname, "EnCana") ~ "Ovintiv, USA (Jan 2020)",  # no longer foreign now headquartered in Denver
                              str_detect(uname, "Total") ~ "TotalEnergies, France (Jul 2021)",  # 4-digit year
                              TRUE ~ uname))
unames %>%
  filter(uname != table_name)

saveRDS(unames, here::here("data_interim", "unames.rds"))


# GET GLOBAL TOTALS BY YEAR, BROKEN DOWN AS MUCH AS POSSIBLE ----
#..First, SumFiles CAUTION this does not include methane ----
path <- here::here("data_raw", sfiles)

coal1 <- read_excel(path, sheet="Sum Oil, Gas, Coal, & Cement", 
                    range="E47:HZ47",
                    col_names = paste0("y", 1790:2019),  # note that this ends in 2019!
                    col_types = "text")


cement1 <- read_excel(path, sheet="Sum Oil, Gas, Coal, & Cement", 
                      range="EM61:HZ61",
                      col_names = paste0("y", 1928:2019),  # note that this ends in 2019!
                      col_types = "text")

total1 <- read_excel(path, sheet="Sum Oil, Gas, Coal, & Cement", 
                     range="E89:HZ89",
                     col_names = paste0("y", 1790:2019),  # note that this ends in 2019!
                     col_types = "text")

global_sumfiles <- bind_rows(coal1 %>% mutate(fuel="coal"),
                           cement1 %>% mutate(fuel="cement"),
                           total1 %>% mutate(fuel="total_xmeth")) %>%
  pivot_longer(cols=-fuel, names_to = "year") %>%
  mutate(year=as.integer(str_sub(year, 2, 5)),
         value=as.numeric(value),
         source="1751-2018 sums by entity [SumFiles Boyd May21(Sent in afternoon).xlsx]") %>%
  filter(!is.na(year), !is.na(value))
summary(global_sumfiles)

# global_sumfiles %>% filter(fuel=="total_xmeth", year %in% 2000:2018) %>% summarise(value=sum(value))


#..Now methane from EachEvery ----
path <- here::here("data_raw", seachev)
df1 <- read_excel(path, sheet = "Sum each CME", range="J576:FW580",
                 col_names = paste0("y", 1850:2019),
                 col_types = "text")

global_ee <- df1 %>%
  mutate(fuel=case_when(row_number()==1 ~ "total_xmeth",
                        row_number()==5 ~ "total",
                        TRUE ~ "other")) %>%
  filter(str_detect(fuel, "total")) %>%
  pivot_longer(cols = -fuel, names_to = "year") %>%
  mutate(year=as.integer(str_sub(year, 2, 5)),
         value=as.numeric(value)) %>%
  filter(!is.na(year), !is.na(value)) %>%
  pivot_wider(names_from = fuel) %>%
  mutate(methane=total - total_xmeth) %>%
  pivot_longer(cols= - year, names_to = "fuel") %>%
  filter(value != 0) %>%
  mutate(source="SumEach&Every1850-2018 May21.xlsx")

#.. combine to get global totals ----
global_totals1 <- bind_rows(global_sumfiles,
                           global_ee)
# check to see if duplicates agree
dups <- global_totals1 %>%
  group_by(fuel, year) %>%
  mutate(n=n()) %>%
  filter(n > 1) %>%
  mutate(sname=case_when(str_detect(source, "SumFiles") ~ "sf",
                         str_detect(source, "Each") ~ "ee",
                         TRUE ~ "error")) %>%
  ungroup
count(dups, sname)
dups2 <- dups %>%
  select(fuel, year, sname, value) %>%
  pivot_wider(names_from = sname) %>%
  mutate(diff=ee - sf) %>%
  arrange(-abs(diff))
dups2
# all good - let's just keep the eachevery values for dups -- we'll drop 170

global_totals <- global_totals1 %>%
  group_by(fuel, year) %>%
  arrange(desc(source)) %>%
  filter(row_number()==1) %>%
  ungroup %>%
  arrange(fuel, year)
summary(global_totals)
count(global_totals, fuel)
saveRDS(global_totals, here::here("data_interim", "global_totals.rds"))


# check the totals ex methane
global_totals %>%
  filter(fuel=="total_xmeth") %>%
  ggplot(aes(year, value, color=str_sub(source, 1, 10))) +
  geom_line() +
  geom_point()

global_totals %>%
  filter(fuel=="total_xmeth", year <= 1860) %>%
  arrange(year, source)

global_totals %>%
  filter(fuel=="total", year %in% 1965:2018) %>%
  summarise(value=sum(value, na.rm=TRUE))


# FINAL create files for analysis ----
# load(file=here::here("data_interim", "eachevery.rdata"), verbose=TRUE)
#.. get files ----
unames <- readRDS(here::here("data_interim", "unames.rds"))
ptypes1 <- readRDS(here::here("data_interim", "ptypes.rds"))
fuels1 <- readRDS(here::here("data_interim", "fuels.rds"))
sumfiles1 <- readRDS(here::here("data_interim", "sumfiles_1751_2018.rds"))
eachevery1 <- readRDS(here::here("data_interim", "eachevery.rds"))
global_totals <- readRDS(here::here("data_interim", "global_totals.rds"))

#.. reorganize, putting unique id and name on each file ----
ptypes2 <- ptypes1 %>%
  left_join(unames, by = c("id_ee", "pname_ee")) %>%
  select(uid, uname, ptype, source) %>%
  arrange(uid)

fuels2 <- fuels1 %>%
  left_join(unames, by = c("id_ee", "pname_ee")) %>%
  select(uid, uname, fuelavail, yearsavail, source) %>%
  arrange(uid)

sumfiles2 <- sumfiles1 %>%
  filter(pid!=999) %>%
  rename(id_sf=pid) %>%
  left_join(unames, by = c("id_sf")) %>%
  select(uid, uname, fgroup, fuel, egroup, emit, value, source) %>%
  arrange(uid)

# create 1750-2018 coal, cement, total, and non-cement and non-coal totals by company ---
count(sumfiles2, fgroup, fuel)
count(sumfiles2, emit)
coalcement_17502018 <- sumfiles2 %>%
  filter(emit=="total") %>%
  filter(fgroup=="detail" & fuel %in% c("cement", "coal") |
           fgroup=="total" & fuel=="total") %>%
  select(uid, uname, fuel, value, source) %>%
  pivot_wider(names_from = fuel, values_fill = 0) %>%
  mutate(coal_share=coal / total,
         cement_share=cement / total) %>%
  select(uid, uname, cement, coal, total, coal_share, cement_share, source)
summary(coalcement_17502018)

eachevery2 <- eachevery1 %>%
  rename(id_ee=id, pname_ee=pname) %>%
  left_join(unames, by = c("id_ee", "pname_ee")) %>%
  select(uid, uname, ptype, ename, etype, year, value, source) %>%
  arrange(uid)

# TO COME estimate dca base as total less coal, cement ----
# dca_base <- unames


#.. save the files ----
ptypes <- ptypes2
fuels <- fuels2
sumfiles <- sumfiles2
eachevery <- eachevery2
save(unames, ptypes, fuels, sumfiles, eachevery, coalcement_17502018, global_totals, file=here::here("data", "pollpay_data.rdata"))

load(file=here::here("data", "pollpay_data.rdata"), verbose=TRUE)

# create a crosswalk on names ----
load(file=here::here("data", "maindata.rdata"), verbose=TRUE)  # old data

load(file=here::here("data", "constants.rdata"), verbose=TRUE)  # overwrite maindata constants
load(file=here::here("data", "pollpay_data.Rdata"), verbose=TRUE)

# create a crosswalk between the new and old unames
dnew <- unames %>% select(uid, uname, pname_oil, table_name)
#  mutate(sound=refinedSoundex(uname, clean=FALSE, maxCodeLen=12))
dold <- uninames %>% select(uid_old=uid, uname_old=uniname, pname_oil=pname1_sumfiles)
xwalk <- full_join(dnew, dold, by = "pname_oil")
saveRDS(xwalk, here::here("data", "xwalk.rds"))


# END ----


global_totals %>%
  filter(year %in% 1965:2018) %>%
  group_by(fuel) %>%
  summarise(value=sum(value, na.rm=TRUE)) %>%
  pivot_wider(names_from = fuel) %>%
  mutate(cement_share=cement / total,
         coal_share=coal / total,
         xcement=total - cement)








# .................... DIVIDER BETWEEN NEW AND OLD ITEMS ......................----
# OLD  SumEvery 1850-2018 ----
#.. Notes: ----
#..    "Sum each CME" has each entity, CO2 and methane separately, broken down by year, NOT broken down by fuel or emissions type ----
#..    "All CMEs 1965-2018" has each entity's sum, NOT broken down by year, fuel, or emissions type (useful as a check) ----

#.. "Sum each CME" ----
path <- here::here("data_raw", severy)
df <- read_excel(path, sheet="Sum each CME", range="B14:FV573",  # include the totals
              col_names = FALSE, col_types = "text")

years <- read_excel(path, sheet="Sum each CME", range="J11:FV11", col_names = FALSE, col_types = "text") %>%
  as.numeric
yearcols <- 9:ncol(df)
length(years) - length(yearcols)  # should be zero

glimpse(df)
cnames <- c("alphaid", "pename", "petype", "junk1", "note1", "note2", "junk2", "junk3", paste0("y", years))
length(cnames) - ncol(df)

df2 <- df %>%
  setNames(cnames) %>%
  select(-starts_with("junk"))
glimpse(df2[, 1:10])

df3 <- df2 %>%
  mutate(pname=ifelse(!is.na(alphaid), pename, NA_character_),
         ename=ifelse(is.na(alphaid), pename, NA_character_),
         ptype=ifelse())
count(df3, pname)
count(df3, ename)

glimpse(df3)


oilngl_vnames <- paste0("oilngl_", c("production", "flaring", "vented", "methane", "total"))
oilngl <- read_excel(path, sheet="Sum Ranking", range="B18:M127",  # include the totals
                     col_names = FALSE, col_types = "text") %>%
  select(1, 3, 5, 7:8, 11, 12) %>%
  setNames(c("pid", "pname", oilngl_vnames)) %>%
  filter(!is.na(pname)) %>%
  mutate(pid=ifelse(is.na(pid), "999", pid))
glimpse(oilngl)




# create uniform ids for entities ----


