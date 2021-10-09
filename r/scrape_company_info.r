library(tidyverse) # overall grammar
library(tidytext)  # only for reorder_within function
library(scales)    # only for scales function
library(httr)      # http verbs
library(rvest)     # wrapper around xlm2 and httr
library(robotstxt) # only for paths_allowed
library(edgar)


options(tibble.print_max = 65, tibble.print_min = 65) 

# getBusinDescr(cik.no, filing.year, useragent)
# https://www.sec.gov/os/accessing-edgar-data

# current cik list
# https://www.sec.gov/Archives/edgar/cik-lookup-data.txt

# ticker cik https://www.sec.gov/include/ticker.txt
# as json: https://www.sec.gov/files/company_tickers.json


# useragent <-  "Your Name Contact@domain.com"
useragent <- "Don Boyd (don@boydresearch.com)"
output <- getBusinDescr(cik.no = c(1000180, 38079), filing.year = 2019, useragent)


xwalk <- read_table("https://www.sec.gov/include/ticker.txt", col_names=c("symbol", "cik"))

tickers <- c("xom")
cikdf <- tibble(symbol=tickers) %>%
  left_join(xwalk, by="symbol")

ciks <- c(34088)  # xom
output2 <- getBusinDescr(cik.no = cikdf$cik, filing.year = 2019, useragent)

# Business descriptions are stored in 'Business descriptions text' directory.

library(finreportr)

# devtools::install_github("mwaldstein/edgarWebR")
# https://cran.r-project.org/web/packages/edgarWebR/readme/README.html
library(edgarWebR)
# EDGARWEBR_USER_AGENT
useragent <- "Don Boyd (don@boydresearch.com)"
useragent <- "Don Boyd (donboyd5@gmail.com)"
Sys.setenv(EDGARWEBR_USER_AGENT = useragent)

edgar_agent <- Sys.getenv(
  "EDGARWEBR_USER_AGENT",
  unset = "edgarWebR (https://github.com/mwaldstein/edgarWebR)"
)
Sys.getenv("EDGARWEBR_USER_AGENT")
company_filings("AAPL", type = "10-K", count = 10)
df <- company_information("AAPL")

df <- company_information("XOM")
df2 <- company_details("XOM")
df2 <- full_text("XOM")

df3 <- GetIncome(symbol="XOM", year=2020)
df4 <- AnnualReports(symbol="XOM", foreign = FALSE)

parse_text_filing(x, strip = TRUE, include.raw = FALSE, fix.errors = TRUE)

dfi <- try(full_text('intel'))
dfi$content
dfi[28, ]

edgar_agent <- Sys.getenv(
  "EDGARWEBR_USER_AGENT",
  unset = "edgarWebR (https://github.com/mwaldstein/edgarWebR)"
)


tbl.Symbols <- read_html("https://en.wikipedia.org/wiki/List_of_S%26P_500_companies") %>%
  html_nodes(css = "table[id='constituents']") %>%
  html_table() %>%
  data.frame() %>%
  as_tibble() %>% 
  select(Symbol, Company = Security, Sector = GICS.Sector, Industry = GICS.Sub.Industry) %>%
  mutate(Symbol = str_replace(Symbol, "[.]", "-")) %>%
  arrange(Symbol)

tmp <- count(tbl.Symbols, Sector, Industry)

tbl.Symbols %>%
  filter(Sector=="Energy")

















# --
# Importing the rvest library 
# It internally imports xml2 library too 
# --
library(rvest)


# --
# Load the link of Holders tab in a variable, here link
# --
link <- "https://finance.yahoo.com/quote/PS/holders?p=PS"

link <- "https://finance.yahoo.com/quote/RDSA.AS/profile?p=RDSA.AS"


# --
# Read the HTML webpage using the xml2 package function read_html()
# --
driver <- read_html(link)


# --
# Since we know there is a tabular data on the webpage, we pass "table" as the CSS selector
# The variable "allTables" will hold all three tables in it
# --
allTables <- html_nodes(driver, css = "table")

html_table(allTables)[[1]]
html_table(allTables)[[2]]


# --
# Fetch any of the three tables based on their index
# 1. Major Holders
# --
majorHolders <- html_table(allTables)[[1]]
majorHolders

#       X1                                    X2
# 1   5.47%       % of Shares Held by All Insider
# 2 110.24%      % of Shares Held by Institutions
# 3 116.62%       % of Float Held by Institutions
# 4     275 Number of Institutions Holding Shares


# --
# 2. Top Institutional Holders
# --
topInstHolders <- html_table(allTables)[[2]]
topInstHolders

#                             Holder     Shares Date Reported  % Out       Value
# 1      Insight Holdings Group, Llc 18,962,692  Dec 30, 2019 17.99% 326,347,929
# 2                         FMR, LLC 10,093,850  Dec 30, 2019  9.58% 173,715,158
# 3       Vanguard Group, Inc. (The)  7,468,146  Dec 30, 2019  7.09% 128,526,792
# 4  Mackenzie Financial Corporation  4,837,441  Dec 30, 2019  4.59%  83,252,359
# 5               Crewe Advisors LLC  4,761,680  Dec 30, 2019  4.52%  81,948,512
# 6        Ensign Peak Advisors, Inc  4,461,122  Dec 30, 2019  4.23%  76,775,909
# 7         Riverbridge Partners LLC  4,021,869  Mar 30, 2020  3.82%  44,160,121
# 8          First Trust Advisors LP  3,970,327  Dec 30, 2019  3.77%  68,329,327
# 9       Fred Alger Management, LLC  3,875,827  Dec 30, 2019  3.68%  66,702,982
# 10 ArrowMark Colorado Holdings LLC  3,864,321  Dec 30, 2019  3.67%  66,504,964


# --
# 3. Top Mutual Fund Holders
# --
topMutualFundHolders <- html_table(allTables)[[3]]
topMutualFundHolders


