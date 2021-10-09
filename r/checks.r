


# get data ----

load(file=here::here("data", "constants.rdata"), verbose=TRUE)  # overwrite maindata constants
load(file=here::here("data", "pollpay_data.Rdata"), verbose=TRUE)

# checks vs 1965-2018 ----
# Carbon majors sums
count(eachevery, etype, ename)
df1 <- eachevery %>%
  filter(etype=="MtCO2e", year %in% 1965:2018)

df1 %>%
  summarise(value=sum(value, na.rm=TRUE)) # 1149240 - great, matches EachEvery sheet "Sum each CME", GW573

# global sums
global_totals %>%
  filter(year %in% 1965:2018) %>%
  group_by(fuel) %>%
  summarise(value=sum(value, na.rm=TRUE))
# fuel      value
# <chr>     <dbl>
# 1 cement   37118.
# 2 coal    496227.
# 3 total  1266352.  matches the CO2-only total from eachevery

# compare to 
# global total: eachevery Sum each CME, GW580: 1,409,854 CO2 plus methane
#   CO2 only  1,266,352 


# 2000-2018 ----
load(file=here::here("data", "pollpay_data.rdata"), verbose=TRUE)
count(global_totals, fuel)

#.. global totals ----
# Rick Heede shows:
# Global CDIAC CO2 emissions	 599,190 	MtCO2
# Global CDIAC/EDGAR methane	 1,962 	methane MtCH4
# Global CDIAC CO2 plus methane emissions	 654,133 	MtCO2e


g2000_2018 <- global_totals %>%
  filter(year %in% 2000:2018) %>%
  group_by(fuel) %>%
  summarise(mtc02=sum(value, na.rm=TRUE))
g2000_2018  # good, we match

# now get company totals 2000-2018
eachevery
count(eachevery, etype, ename)
count(eachevery, uid, uname)
# uid 11, British Coal Corporation, drops out

entity2000_2018 <- eachevery %>%
  filter(etype=="MtCO2e", year %in% 2000:2018) %>%
  group_by(uid, uname, ptype) %>%
  summarise(value=sum(value, na.rm=TRUE), .groups="drop")

entity2000_2018 %>%
  filter(uid != 16) %>%  # China, PR coal cement
  mutate(pct=value / g2000_2018$mtc02[g2000_2018$fuel=="total"] * 100) %>%
  arrange(desc(value))
# good - this matches Rick Heede's "Calcs PPCF Sep21.xlsx"








