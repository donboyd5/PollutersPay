

# libraries ----
library(tidyverse)
options(tibble.print_max = 65, tibble.print_min = 65) # if more than 60 rows, print 60 - enough for states
# ggplot2 tibble tidyr readr purrr dplyr stringr forcats

library(purrr)
library(readxl)
library(btools)
library(kableExtra)


# locations ----
# https://data.bls.gov/cew/doc/titles/industry/industry_titles.htm
qdir <- r"(C:\Users\donbo\Downloads\2020_annual_by_industry\2020.annual.by_industry\)"
qfn1 <- "2020.annual 211 NAICS 211 Oil and gas extraction.csv"
qfn1a <- "2020.annual 21112 NAICS 21112 Crude petroleum extraction.csv"
qfn1b <- "2020.annual 21113 NAICS 21113 Natural gas extraction.csv"
qfn2 <- "2020.annual 32411 NAICS 32411 Petroleum refineries.csv"

files <- paste0(qdir, c(qfn1, qfn1a, qfn1b, qfn2))

# get data ----
f <- function(file){
  df <- read_csv(file) %>%
    filter(str_sub(area_fips, 3, 6)=="000", own_code=="5")
  df
}

df <- map_dfr(files, f)
count(df, industry_code)

df2 <- df %>%
  select(area_fips, area_title, industry_code, annual_avg_emplvl) %>%
  filter(annual_avg_emplvl > 0) %>%
  mutate(stname=str_replace(area_title, " -- Statewide", ""),
         ind=factor(industry_code, 
                    levels=c(211, 21112, 21113, 32411), 
                    labels=c("extraction", "extraction_oil", "extraction_ng", "refining"))) %>%
  select(area_fips, stname, ind, annual_avg_emplvl) %>%
  pivot_wider(names_from = ind, values_from = annual_avg_emplvl) %>%
  mutate(extraction_check=extraction - (naz(extraction_oil) + naz(extraction_ng)),
         total=naz(extraction) + naz(refining),
         pct=total / total[stname=="U.S. TOTAL"]) %>%
  select(area_fips, stname, extraction_oil, extraction_ng, extraction, refining, total, extraction_check) %>%
  arrange(-total)
df2

df2 %>%
  write_csv(here::here("results", "qcew_oilng_2020.csv"))



# old below here ----

df <- read_csv(paste0(qdir, qfn))
glimpse(df)

df2 <- df %>%
  filter(str_sub(area_fips, 3, 6)=="000", own_code=="5")
ns(df2)

df3 <- df2 %>%
  select(area_fips, area_title, agglvl_title, annual_avg_emplvl, annual_avg_estabs_count, avg_annual_pay) %>%
  filter(annual_avg_emplvl > 0) %>%
  mutate(stname=str_replace(area_title, " -- Statewide", "")) %>%
  select(area_fips, stname, contains("annual")) %>%
  arrange(-annual_avg_emplvl)
df3

df3 %>%
  write_csv(here::here("results", "qcew_naics211_2020.csv"))

df3 %>%
  write_csv(here::here("results", "qcew_naics324_2020.csv"))

