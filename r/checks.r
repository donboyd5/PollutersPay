
tmp <- sipro %>%
  filter(ticker=="GLNCY") %>%
  select(ticker, company, country, business)

tmp %>% write_csv(here::here("ignore", "tmp.csv"))

tmp <- sipro %>%
  filter(str_detect(company, "Glencore")) %>%
  select(ticker, company, country, business)

sipro %>%
  filter(str_detect(company, "Beers")) %>%
  select(ticker, company, country, business)