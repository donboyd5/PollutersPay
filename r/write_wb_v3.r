

# functions ----
fsoe <- function(wb, sname, data, headrow){
  keyrows$soe_headrow <<- headrow
  
  soe <- data %>%
    filter(group %in% c("SOE", "nation")) %>%
    select(-group) %>%
    arrange(desc(dca_annual))
  
  soe_total <- soe %>%
    summarise(table_name=paste0("  Total for ", nrow(soe), " State-Owned Enterprises"), 
              taxbase=sum(taxbase), 
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


ffioc <- function(wb, sname, data, headrow){
  keyrows$fioc_headrow <<- headrow
  
  fioc <- data %>%
    filter(group %in% c("fIOC")) %>%
    select(-group) %>%
    arrange(desc(dca_annual))
  
  fioc_total <- fioc %>%
    summarise(table_name=paste0("  Total for ", nrow(fioc), " Foreign Investor-Owned Companies"), 
              taxbase=sum(taxbase), 
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


fdioc <- function(wb, sname, data, headrow){
  
  keyrows$dioc_headrow <<- headrow
  dioc <- data %>%
    filter(group %in% c("dIOC")) %>%
    select(-group) %>%
    arrange(desc(dca_annual))
  
  dioc_total <- dioc %>%
    summarise(table_name=paste0("  Total for ", nrow(dioc), " Domestic Investor-Owned Companies"), 
              taxbase=sum(taxbase), 
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


fgtots <- function(wb, sname, data){
  
  gtdata <- data %>%
    filter(group %in% c("SOE", "nation", "fIOC", "dIOC"))
  
  grand_total <- gtdata %>%
    summarise(table_name=paste0("    Grand Total for ", nrow(gtdata), " Entities"), 
              taxbase=sum(taxbase),
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


fnotes <- function(wb, sname, data){
  years <- 2000:2018
  note_intro <- "Notes:"
  note1 <- paste0("  * Historical emissions period is ", years[1], "-", years[length(years)])
  note2 <- paste0("  ** Top ", nrow(data), " assessable entities are subject to tax.")

  
  note_start <- last_line_written + 3
  
  writeData(wb, sheet = sname, x=note_intro, 
            startCol=1, startRow=note_start, rowNames = FALSE, colNames = FALSE)
  writeData(wb, sheet = sname, x=note1, 
            startCol=1, startRow=note_start + 1, rowNames = FALSE, colNames = FALSE)
  writeData(wb, sheet = sname, x=note2, 
            startCol=1, startRow=note_start + 2, rowNames = FALSE, colNames = FALSE)
  return(wb)
}


style <- function(wb, sname){
  
  #.. define styles ----
  num0 <- createStyle(numFmt = "#,##0")
  num1 <- createStyle(numFmt = "#,##0.0")
  curr2 <- createStyle(numFmt = "$0.00")
  num2 <- createStyle(numFmt = "#,##0.00")
  curr3 <- createStyle(numFmt = "$0.000")
  num3 <- createStyle(numFmt = "#,##0.000")
  
  pct1 <- createStyle(numFmt = "0.0%")
  pct2 <- createStyle(numFmt = "0.00%")
  
  # for totals rows
  totstyle <- createStyle(fontSize = 12, fgFill = "#f0f0f0", border = "TopBottom", borderColour = "#f0f0f0")
  bold_style <- createStyle(fontSize = 12, textDecoration = "bold")
  
  #.. constants ----
  comma_cols <- c(2)
  pct_cols <- c(4)
  dollar_cols <- c(6)
  last_col <- 6
  
  first_data_row <- 10
  last_row <- 200
  
  #.. apply styles for entire workbook ----
  modifyBaseFont(wb, fontSize = 12, fontColour = "black", fontName = "Calibri")
  addStyle(wb, sheet = sname, style = num0, rows=first_data_row:last_row, cols = comma_cols, gridExpand = TRUE)
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
makewb <- function(taxbase, outpath){

  # prep data ----
  use_names <- quote(c(uname, group, taxbase, blank1, pct_taxbase, blank2, dca_annual))
  
  data <- taxbase %>%
    mutate(blank1="", blank2="") %>%  # so we can have blank columns
    select(!!use_names)
  
  # initialize ----
  shellfile <- "PPCompanyShell_v4.xlsx"
  snames <- getSheetNames(here::here("data_raw", shellfile))
  wb <- loadWorkbook(file = here::here("data_raw", shellfile))
  # get the sheets we want
  sname <- "TaxTable"
  cloneWorksheet(wb, sheetName=sname, clonedSheet = "TableShell")
  purrr::walk(snames, function(x) removeWorksheet(wb, x))
  
  last_line_written <<- 7  # global !!!
  keyrows <<- list() # global
  
  # create report sections ----
  wb <- fsoe(wb, sname, data, headrow=last_line_written + 3)
  wb <- ffioc(wb, sname, data, headrow=last_line_written + 3)
  wb <- fdioc(wb, sname, data, headrow=last_line_written + 3)
  
  wb <- fgtots(wb, sname, data)
  wb <- fnotes(wb, sname, data)
  wb <- style(wb, sname)
  
  #.. end workbook function ----
  saveWorkbook(wb, file = outpath, overwrite = TRUE)
  return(wb)
}







