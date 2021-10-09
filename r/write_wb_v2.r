

# functions ----
fsoe <- function(wb, sname, df, use_names, headrow){
  keyrows$soe_headrow <<- headrow
  
  soe <- df %>%
    filter(payor, group %in% c("SOE", "nation")) %>%
    select(!!use_names) %>%
    arrange(desc(dca_annual))
  
  soe_total <- soe %>%
    summarise(table_name=paste0("  Total for ", nrow(soe), " State-Owned Enterprises"), 
              pct_emissions=sum(pct_emissions), 
              blank1=first(blank1),
              pct_taxbase=sum(pct_taxbase),
              blank2=first(blank2),
              dca_annual=sum(dca_annual))
  
  # write the details for SOE
  writeData(wb, sheet = sname, x="State-Owned Enterprises (SOEs)", startCol=1, startRow=headrow, rowNames = FALSE, colNames = FALSE)
  last_line_written <<- keyrows$soe_headrow
  
  soe_start <- last_line_written + 1
  writeData(wb, sheet = sname, x=soe, startCol=1, startRow=soe_start, rowNames = FALSE, colNames = FALSE)
  last_line_written <<- soe_start + nrow(soe)
  
  # now write the SOE summary
  keyrows$soe_totalrow <<- last_line_written + 1
  writeData(wb, sheet = sname, x=soe_total, 
            startCol=1, startRow=keyrows$soe_totalrow, rowNames = FALSE, colNames = FALSE)
  last_line_written <<- keyrows$soe_totalrow
  return(wb)
}


ffioc <- function(wb, sname, df, use_names, headrow){
  keyrows$fioc_headrow <<- headrow
  
  fioc <- df %>%
    filter(payor, group %in% c("fIOC")) %>%
    select(!!use_names) %>%
    arrange(desc(dca_annual))
  
  fioc_total <- fioc %>%
    summarise(table_name=paste0("  Total for ", nrow(fioc), " Foreign Investor-Owned Companies"), 
              pct_emissions=sum(pct_emissions), 
              blank1=first(blank1),
              pct_taxbase=sum(pct_taxbase),
              blank2=first(blank2),
              dca_annual=sum(dca_annual))
  
  # write the details for fioc
  writeData(wb, sheet = sname, x="Foreign Investor-Owned Companies (IOCs)", startCol=1, startRow=headrow, rowNames = FALSE, colNames = FALSE)
  last_line_written <<- keyrows$fioc_headrow
  
  fioc_start <- last_line_written + 1
  writeData(wb, sheet = sname, x=fioc, startCol=1, startRow=fioc_start, rowNames = FALSE, colNames = FALSE)
  last_line_written <<- fioc_start
  
  # now write the fioc summary
  keyrows$fioc_totalrow <<- last_line_written + nrow(fioc) + 1
  writeData(wb, sheet = sname, x=fioc_total, 
            startCol=1, startRow=keyrows$fioc_totalrow, rowNames = FALSE, colNames = FALSE)
  last_line_written <<- keyrows$fioc_totalrow
  
  return(wb)
}


fdioc <- function(wb, sname, df, use_names, headrow){
  
  keyrows$dioc_headrow <<- headrow
  dioc <- df %>%
    filter(payor, group %in% c("dIOC")) %>%
    select(!!use_names) %>%
    arrange(desc(dca_annual))
  
  dioc_total <- dioc %>%
    summarise(table_name=paste0("  Total for ", nrow(dioc), " Domestic Investor-Owned Companies"), 
              pct_emissions=sum(pct_emissions), 
              blank1=first(blank1),
              pct_taxbase=sum(pct_taxbase),
              blank2=first(blank2),
              dca_annual=sum(dca_annual))
  
  # write the details for dioc
  writeData(wb, sheet = sname, x="Domestic Investor-Owned Companies (IOCs)", startCol=1, startRow=headrow,
            rowNames = FALSE, colNames = FALSE)
  last_line_written <<- keyrows$dioc_headrow
  
  dioc_start <- last_line_written + 1
  writeData(wb, sheet = sname, x=dioc, startCol=1, startRow=dioc_start, rowNames = FALSE, colNames = FALSE)
  last_line_written <<- dioc_start
  
  # now write the fdioc summary
  keyrows$dioc_totalrow <<- last_line_written + nrow(dioc) + 1
  writeData(wb, sheet = sname, x=dioc_total, 
            startCol=1, startRow=keyrows$dioc_totalrow, rowNames = FALSE, colNames = FALSE)
  last_line_written <<- keyrows$dioc_totalrow
  
  return(wb)
}


fgtots <- function(wb, sname, df){
  
  gtdata <- df %>%
    filter(payor, group %in% c("SOE", "nation", "fIOC", "dIOC"))
  
  grand_total <- gtdata %>%
    summarise(table_name=paste0("    Grand Total for ", nrow(gtdata), " Entities"), 
              pct_emissions=sum(pct_emissions),
              blank1=first(blank1),
              pct_taxbase=sum(pct_taxbase),
              blank2=first(blank2),
              dca_annual=sum(dca_annual))
  
  keyrows$grand_totalrow <<- last_line_written + 3
  
  writeData(wb, sheet = sname, x=grand_total, 
            startCol=1, startRow=keyrows$grand_totalrow, rowNames = FALSE, colNames = FALSE)
  last_line_written <<- keyrows$grand_totalrow
  return(wb)
}


fnotes <- function(wb, sname, threshold, exclusion){
  years <- 2000:2018
  note_intro <- "Notes:"
  note1 <- paste0("  * Historical emissions period is ", years[1], "-", years[length(years)])
  note2 <- paste0("  ** Assumes the Climage Damage Compensation Tax is only applied to polluters accounting for at least ", 
                  scales::percent(threshold, accuracy=.001), " of global emissions.")
  if(exclusion=="none"){
    note3 <- "  ** Share of tax is computed without excluding a threshold amount of emissions."
  } else note3 <- "  ** Share of tax is computed after excluding a threshold amount of emissions."
  
  note_start <- last_line_written + 3
  
  writeData(wb, sheet = sname, x=note_intro, 
            startCol=1, startRow=note_start, rowNames = FALSE, colNames = FALSE)
  writeData(wb, sheet = sname, x=note1, 
            startCol=1, startRow=note_start + 1, rowNames = FALSE, colNames = FALSE)
  writeData(wb, sheet = sname, x=note2, 
            startCol=1, startRow=note_start + 2, rowNames = FALSE, colNames = FALSE)
  writeData(wb, sheet = sname, x=note3, 
            startCol=1, startRow=note_start + 3, rowNames = FALSE, colNames = FALSE)
  return(wb)
}


style <- function(wb, sname){
  
  #.. define styles ----
  curr2 <- createStyle(numFmt = "$0.00")
  num2 <- createStyle(numFmt = "0.00")
  curr3 <- createStyle(numFmt = "$0.000")
  num3 <- createStyle(numFmt = "0.000")
  
  pct1 <- createStyle(numFmt = "0.0%")
  pct2 <- createStyle(numFmt = "0.00%")
  
  # for totals rows
  totstyle <- createStyle(fontSize = 11, fgFill = "#f0f0f0", border = "TopBottom", borderColour = "#f0f0f0")
  bold_style <- createStyle(fontSize = 11, textDecoration = "bold")
  
  #.. constants ----
  pct_cols <- c(2, 4)
  dollar_cols <- c(6)
  last_col <- 6
  
  first_data_row <- 10
  last_row <- 200
  
  #.. apply styles for entire workbook ----
  addStyle(wb, sheet = sname, style = pct2, rows=first_data_row:last_row, cols = pct_cols, gridExpand = TRUE)
  addStyle(wb, sheet = sname, style = curr3, rows=first_data_row, cols = dollar_cols)
  addStyle(wb, sheet = sname, style = num3, rows=(first_data_row + 1):last_row, cols = dollar_cols, gridExpand = TRUE)
  
  #.. apply styles for each block ----
  style_block <- function(wb, sname, headrow, totalrow){
    # my styles, last_col, and dollar_cols are defined above
    addStyle(wb, sheet = sname, style = totstyle, rows = headrow, cols = 1:last_col, stack=TRUE)
    addStyle(wb, sheet = sname, style = bold_style, rows = headrow, cols = 1, stack=TRUE)
    addStyle(wb, sheet = sname, style = curr3, rows=totalrow, cols = dollar_cols)
    addStyle(wb, sheet = sname, style = totstyle, rows = totalrow, cols = 1:last_col, stack=TRUE)
    return(wb)
  }
  
  wb <- style_block(wb, sname, keyrows$soe_headrow, keyrows$soe_totalrow)
  wb <- style_block(wb, sname, keyrows$fioc_headrow, keyrows$fioc_totalrow)
  wb <- style_block(wb, sname, keyrows$dioc_headrow, keyrows$dioc_totalrow)
  
  addStyle(wb, sheet = sname, style = curr3, rows=keyrows$grand_totalrow, cols = dollar_cols)
  addStyle(wb, sheet = sname, style = totstyle, rows = keyrows$grand_totalrow, cols = 1:last_col, stack=TRUE)
  addStyle(wb, sheet = sname, style = bold_style, rows = keyrows$grand_totalrow, cols = 1, stack=TRUE)
  
  return(wb)
}


#.. full report ---
makewb <- function(tottax, dca_base, nexus, threshold, exclusion, coal_entities, bankrupt, gtot, outpath){
  
  # prep data ----
  taxbase <- dca_base %>%
    left_join(nexus, by = c("uid", "uname")) %>%
    mutate(coalco=uid %in% coal_entities,
           bankrupt= uid %in% !!bankrupt,
           pct_emissions = dca_base / !!gtot,
           threshold=!!threshold,
           sizeok = pct_emissions >= threshold,
           payor= nexus & !coalco & !bankrupt & sizeok,
           taxbase_nox=dca_base * payor,
           taxbase_ex=(dca_base - threshold * !!gtot) * payor,
           exclude=!!exclusion,
           taxbase=ifelse(exclude=="none", taxbase_nox, taxbase_ex),
           pct_taxbase=taxbase / sum(taxbase),
           dca_annual=pct_taxbase * tottax,
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
  
  
  df <- taxbase %>%
    mutate(blank1="", blank2="") # so we can have blank columns
  
  use_names <- quote(c(uname, pct_emissions, blank1, pct_taxbase, blank2, dca_annual))
  
  # initialize ----
  shellfile <- "PPCompanyShell_v3.xlsx"
  snames <- getSheetNames(here::here("data_raw", shellfile))
  wb <- loadWorkbook(file = here::here("data_raw", shellfile))
  # get the sheets we want
  sname <- "TaxTable"
  cloneWorksheet(wb, sheetName=sname, clonedSheet = "TableShell")
  purrr::walk(snames, function(x) removeWorksheet(wb, x))
  
  last_line_written <<- 7  # global !!!
  keyrows <<- list() # global
  
  wb <- fsoe(wb, sname, df, use_names, headrow=last_line_written + 3)
  wb <- ffioc(wb, sname, df, use_names, headrow=last_line_written + 3)
  wb <- fdioc(wb, sname, df, use_names, headrow=last_line_written + 3)
  
  wb <- fgtots(wb, sname, df)
  wb <- fnotes(wb, sname, threshold, exclusion)
  wb <- style(wb, sname)
  
  #.. end workbook function ----
  saveWorkbook(wb, file = outpath, overwrite = TRUE)
  return(wb)
}







