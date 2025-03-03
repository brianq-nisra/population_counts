# Ctrl + Shift + S to run
# or click Source

library(jsonlite)
library(dplyr)
library(openxlsx)

# CPD file processing ####

# Fix CSV file

suppressWarnings({
  cpd_raw <- readLines("data/CPD_LIGHT_JULY_2024.csv")
})

cpd_fixed <- gsub('Armagh City, Banbridge and Craigavon', '"Armagh City, Banbridge and Craigavon"', cpd_raw, fixed = TRUE)
cpd_fixed <- gsub('Shankill (Armagh, Banbridge and Craigavon)', '"Shankill (Armagh, Banbridge and Craigavon)"', cpd_fixed, fixed = TRUE)
cpd_fixed <- gsub('Dromore (Armagh, Banbridge and Craigavon)', '"Dromore (Armagh, Banbridge and Craigavon)"', cpd_fixed, fixed = TRUE)
cpd_fixed <- gsub('Cathedral (Armagh, Banbridge and Craigavon)', '"Cathedral (Armagh, Banbridge and Craigavon)"', cpd_fixed, fixed = TRUE)

cpd_fixed <- gsub('Newry, Mourne and Down', '"Newry, Mourne and Down"', cpd_fixed, fixed = TRUE)
cpd_fixed <- gsub('Cathedral ("Newry, Mourne and Down")', '"Cathedral (Newry, Mourne and Down)"', cpd_fixed, fixed = TRUE)
cpd_fixed <- gsub('Abbey ("Newry, Mourne and Down")', '"Abbey (Newry, Mourne and Down)"', cpd_fixed, fixed = TRUE)

cpd_fixed <- gsub('Boho,Cleenish and Letterbreen', '"Boho,Cleenish and Letterbreen"', cpd_fixed, fixed = TRUE)

cpd_fixed <- gsub('ANNAGHMORE (ARMAGH CITY, BANBRIDGE AND CRAIGAVON LGD)', '"ANNAGHMORE (ARMAGH CITY, BANBRIDGE AND CRAIGAVON LGD)"', cpd_fixed, fixed = TRUE)
cpd_fixed <- gsub('BAND H (ARMAGH CITY, BANBRIDGE AND CRAIGAVON LGD)', '"BAND H (ARMAGH CITY, BANBRIDGE AND CRAIGAVON LGD)"', cpd_fixed, fixed = TRUE)
cpd_fixed <- gsub('DROMORE (ARMAGH CITY, BANBRIDGE AND CRAIGAVON LGD)', '"DROMORE (ARMAGH CITY, BANBRIDGE AND CRAIGAVON LGD)"', cpd_fixed, fixed = TRUE)
cpd_fixed <- gsub('EGLISH (ARMAGH CITY, BANBRIDGE AND CRAIGAVON LGD)', '"EGLISH (ARMAGH CITY, BANBRIDGE AND CRAIGAVON LGD)"', cpd_fixed, fixed = TRUE)
cpd_fixed <- gsub('KILMORE (ARMAGH CITY, BANBRIDGE AND CRAIGAVON LGD)', '"KILMORE (ARMAGH CITY, BANBRIDGE AND CRAIGAVON LGD)"', cpd_fixed, fixed = TRUE)

cpd_fixed <- gsub('BAND H (NEWRY, MOURNE AND DOWN LGD)', '"BAND H (NEWRY, MOURNE AND DOWN LGD)"', cpd_fixed, fixed = TRUE)
cpd_fixed <- gsub('BELLEEK (NEWRY, MOURNE AND DOWN LGD)', '"BELLEEK (NEWRY, MOURNE AND DOWN LGD)"', cpd_fixed, fixed = TRUE)
cpd_fixed <- gsub('CREGGAN (NEWRY, MOURNE AND DOWN LGD)', '"CREGGAN (NEWRY, MOURNE AND DOWN LGD)"', cpd_fixed, fixed = TRUE)
cpd_fixed <- gsub('GLEN (NEWRY, MOURNE AND DOWN LGD)', '"GLEN (NEWRY, MOURNE AND DOWN LGD)"', cpd_fixed, fixed = TRUE)
cpd_fixed <- gsub('KILLEEN (NEWRY, MOURNE AND DOWN LGD)', '"KILLEEN (NEWRY, MOURNE AND DOWN LGD)"', cpd_fixed, fixed = TRUE)
cpd_fixed <- gsub('KILMORE (NEWRY, MOURNE AND DOWN LGD)', '"KILMORE (NEWRY, MOURNE AND DOWN LGD)"', cpd_fixed, fixed = TRUE)
cpd_fixed <- gsub('MAGHERA (NEWRY, MOURNE AND DOWN LGD)', '"MAGHERA (NEWRY, MOURNE AND DOWN LGD)"', cpd_fixed, fixed = TRUE)


# save new CSV file

writeLines(cpd_fixed, "data/CPD_LIGHT_JULY_2024_fixed.csv")

# Read in new file
cpd_data <- read.csv("data/CPD_LIGHT_JULY_2024_fixed.csv")

# Population figures from NISRA data portal ####

# Change year here
# Check https://data.nisra.gov.uk/table/MYE01T011 to see if new year available (usually updated end of July)
year <- 2022

mye_raw <- fromJSON(
  paste0(
    "https://ws-data.nisra.gov.uk/public/api.jsonrpc?data=%7B%22jsonrpc%22:%222.0%22,%22method%22:%22PxStat.Data.Cube_API.ReadDataset%22,%22params%22:%7B%22class%22:%22query%22,%22id%22:%5B%22TLIST(A1)%22%5D,%22dimension%22:%7B%22TLIST(A1)%22:%7B%22category%22:%7B%22index%22:%5B%22",
    year,
    "%22%5D%7D%7D%7D,%22extension%22:%7B%22pivot%22:null,%22codes%22:false,%22language%22:%7B%22code%22:%22en%22%7D,%22format%22:%7B%22type%22:%22JSON-stat%22,%22version%22:%222.0%22%7D,%22matrix%22:%22MYE01T011%22%7D,%22version%22:%222.0%22%7D%7D"
    )
  )$result


mye_data <- data.frame(DZ2021 = mye_raw$dimension$DZ2021$category$index,
                       Population = mye_raw$value)

# Look up SOA2001 ####

# Check for spelling/groupings for converting old SOA2001 codes to DZ2021 codes
dz_areas <- cpd_data %>% 
  mutate(Area = case_when(SOA2001NAME_OLD %in% c("The Mount_1",
                                                 "The Mount_2", 
                                                 "Ballymacarrett_2", 
                                                 "Ballymacarrett_3") ~ "East Belfast",
                          SOA2001NAME_OLD %in% c("Clandeboye_2",
                                                 "Clandeboye_3",
                                                 "Conlig_3") ~ "North Down",
                          SOA2001NAME_OLD %in% c("Brandywell",
                                                 "Creggan Central_1",
                                                 "Creggan Central_2",
                                                 "Creggan South") ~ "Derry",
                          SOA2001NAME_OLD %in% c("Poleglass_1",
                                                 "Poleglass_2",
                                                 "Falls_2",
                                                 "Falls_3",
                                                 "Twinbrook_1",
                                                 "Twinbrook_2",
                                                 "Whiterock_2",
                                                 "Whiterock_3",
                                                 "Upper Springfield_1",
                                                 "Upper Springfield_3") ~ "West Belfast",
                          SOA2001NAME_OLD %in% c("Kilwaughter_1",
                                                 "Kilwaughter_2",
                                                 "Antiville",
                                                 "Northland",
                                                 "Love Lane") ~ "Carrick and Larne",
                          SOA2001NAME_OLD %in% c("Woodvale_1",
                                                 "Woodvale_2",
                                                 "Woodvale_3",
                                                 "Shankill_1",
                                                 "Shankill_2") ~ "Shankill",
                          SOA2001NAME_OLD %in% c("Ardoyne_1",
                                                 "Ardoyne_2",
                                                 "Ardoyne_3",
                                                 "New Lodge_1",
                                                 "New Lodge_2",
                                                 "New Lodge_3") ~ "North Belfast",
                          SOA2001NAME_OLD %in% c("Drumgask_2",
                                                 "Drumnamoe_1",
                                                 "Drumnamoe_2") ~ "Lurgan")
         ) %>% 
  group_by(Area, DZ2021, DZ2021_name) %>% 
  summarise() %>% 
  filter(!is.na(Area)) %>% 
  left_join(mye_data,
            by = "DZ2021") %>% 
  mutate(link = paste0("https://explore.nisra.gov.uk/local-stats/", DZ2021))

# Sum up for populations
populations <- dz_areas %>% 
  group_by(Area) %>% 
  summarise(Population = sum(Population))

# Excel creation ####

source("style.R")

wb <- createWorkbook()

addWorksheet(wb, "Areas")

writeData(wb, "Areas",
          x = "DZ2021 Codes")

addStyle(wb, "Areas",
         style = titles,
         rows = 1,
         cols = 1)

writeDataTable(wb, "Areas",
               x = dz_areas,
               tableStyle = "none",
               startRow = 2,
               headerStyle = bold)

links <- dz_areas$link

names(links) <- paste("Link to map for", dz_areas$DZ2021_name)
class(links) <- "hyperlink"

writeData(wb, "Areas",
          x = links,
          startRow = 3,
          startCol = 5)

formulae <- paste0("=SUBTOTAL(103, A", 2 + 1:nrow(dz_areas), ")")

writeFormula(wb, "Areas",
             x = formulae,
             startRow = 3,
             startCol = 6)



setColWidths(wb, "Areas", cols = 1:ncol(dz_areas), widths = c(40, 10, 30, 13, 43))
setColWidths(wb, "Areas", cols = ncol(dz_areas) + 1, hidden = TRUE)

addWorksheet(wb, "Population Figures")

writeData(wb, "Population Figures",
          x = paste("Population by Area - based on", year, "Mid-Year Estimates"))

addStyle(wb, "Population Figures",
         style = titles,
         rows = 1,
         cols = 1)

writeDataTable(wb, "Population Figures",
               x = populations,
               tableStyle = "none",
               startRow = 2,
               withFilter = FALSE,
               headerStyle = bold)

formulae <- paste0("=SUMIFS(Areas!D:D, Areas!A:A, A", 2 + 1:nrow(populations),", Areas!F:F, 1)")

writeFormula(wb, "Population Figures",
             x = formulae,
             startRow = 3,
             startCol = 2)

addStyle(wb, "Population Figures",
         style = right_bold,
         rows = 2,
         cols = 2)

addStyle(wb, "Population Figures",
         style = numbers,
         rows = 3:(nrow(populations) + 2),
         cols = 2)

setColWidths(wb, "Population Figures", cols = 1:ncol(populations), widths = c(25, 15))

filename <- paste0("population-by-area-", year, ".xlsx")

saveWorkbook(wb, filename, overwrite = TRUE)
openXL(filename)
