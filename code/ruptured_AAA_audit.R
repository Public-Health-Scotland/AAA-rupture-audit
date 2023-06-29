#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# ruptured_AAA_audit.R
# Karen Hotopp
# June 2023
# 
# Ruptured anuerysm audit
# 
# Written/run on Posit WB
# R version 4.1.2
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Methods - initial case ascertainment
# The national inpatients dataset held by PHS (known as SMR01) will be searched 
# for any males, of age 65 and older, who have presented with a ruptured 
# aneurysm (ICD10 codes I71.3, I71.4, I71.5, I71.6, I71.8, and I71.9). 
# National Records of Scotland (NRS) death records will be searched for males, 
# age 65 and older, with a ruptured aneurysm in any position on the death 
# certificate. CHI number will be used to link these two datasets, keeping only 
# one record per patient. The search will be carried out for hospital admissions 
# and/or deaths from 1st June 2021 to 31st May 2022.


# install.packages("odbc")
# install.packages("dbplyr")
# library(odbc)
library(dplyr)
library(lubridate)
library(stringr)
library(phsmethods)
library(openxlsx)
library(tidylog)


rm(list=ls())
gc()


### 1: Housekeeping ----
## Variables
date_start <- dmy("01062021")
date_end <- dmy("31052022")
icd_codes <- c("I713", "I714", "I715", "I716", "I718", "I719")

## Pathways
wd <- paste0("/PHI_conf/AAA/Topics/Projects/20210929-Ruptured-AAAs-Audit")

simd_path <- paste0("/conf/linkage/output/lookups/Unicode/Deprivation",
                    "/postcode_2023_1_simd2020v2.rds")

## SIMD data
simd <- readRDS(simd_path) |> 
  select(pc8, hb2019name)

## Function
write_report <- function(df1, df2, hb_name) {
  
  ### Setup workbook ----
  ## Reset variable names
  records_extract <- c("Health Board", "UPI", "Surname", "Forename", "Postcode",
                       "Date of Birth", "Age", "Sex", "Date of Admission", 
                       "Date Discharge", "Condition Index", "Main Condition", 
                       "Other Condition 1", "Other Condition 2", "Other Condition 3",
                       "Other Condition 4", "Other Condition 5")
  records_deaths <- c("Health Board", "UPI", "Surname", "Forename", "Postcode",
                      "Date of Birth", "Age", "Sex", "Date of Death", "Cause Index",
                      "Underlying Cause of Death", "Cause of Death 0", 
                      "Cause of Death 1", "Cause of Death 2", "Cause of Death 3",
                      "Cause of Death 4", "Cause of Death 5", "Cause of Death 6",
                      "Cause of Death 7", "Cause of Death 8", "Cause of Death 9")
  
  ## Styles
  title_style <- createStyle(fontSize = 14, halign = "Left", textDecoration = "bold")
  table_style <- createStyle(valign = "Bottom", halign = "Left",
                             border = "TopBottomLeftRight")
  wrap_style <- createStyle(wrapText = TRUE)
  
  ## Titles
  title <- paste0("Ruptured Anuerysm Audit for ", hb_name)
  date_range <- paste0("Hospital admissions and deaths between ", date_start,
                       " and ", date_end)
  source <- paste0("Data for hospital admissions has been produced from the ",
                   "national inpatients dataset held by PHS (known as SMR01). ",
                   "Data containing information on deaths has been produced from ",
                   "the National Records of Scotland (NRS) death records.")
  source2 <- paste0("SMR01 and NRS datasets were searched for any males, of ", 
                    "age 65 and older, who have presented with a ruptured ",
                    "aneurysm (ICD10 codes I71.3, I71.4, I71.5, I71.6, ",
                    "I71.8, and I71.9).")
  today <- paste0("Workbook created ", Sys.Date())
  
  ## Data
  # Inpatients (extract_simd)
  data1 <- df1 |> 
    filter(hb2019name == hb_name)
  # Deaths (deaths_simd)
  data2 <- df2 |> 
    filter(hb2019name == hb_name)
  
  
  ## Setup workbook
  wb <- createWorkbook()
  options("openxlsx.borderStyle" = "thin",
          "openxlsx.dateFormat" = "dd/mm/yyyy")
  modifyBaseFont(wb, fontSize = 12, fontName = "Arial")
  
  
  ### Notes ----
  addWorksheet(wb, sheetName = "Notes", gridLines = FALSE)
  writeData(wb, "Notes", title, startRow = 1, startCol = 1)
  writeData(wb, "Notes", date_range, startRow = 2, startCol = 1)
  writeData(wb, "Notes", source, startRow = 4, startCol = 1)
  writeData(wb, "Notes", source2, startRow = 5, startCol = 1)
  writeData(wb, "Notes", today, startRow = 7, startCol = 1)
  
  setColWidths(wb, "Notes", cols = 1, widths = "97.00")
  #setRowHeights(wb, "Notes", rows = 4:5, height = 30)
  addStyle(wb, "Notes", wrap_style, rows = 4:5, cols = 1, gridExpand = TRUE)
  addStyle(wb, "Notes", title_style, rows = 1, cols = 1)
  
  
  ### Inpatient Admissions ----
  names(data1) <- records_extract
  addWorksheet(wb, sheetName = "Inpatient Admissions", gridLines = FALSE)
  writeDataTable(wb, "Inpatient Admissions", data1, startRow = 3)
  
  # titles
  title_extract <- paste0("Inpatient admissions between ", date_start, 
                          " and ", date_end)
  writeData(wb, "Inpatient Admissions", title_extract, startRow = 1, startCol = 1)
  addStyle(wb, "Inpatient Admissions", title_style, rows = 1, cols = 1)
  
  # table headers
  addStyle(wb, "Inpatient Admissions", title_style, rows = 3, cols = 1:ncol(data1))

  # tables
  addStyle(wb, "Inpatient Admissions", table_style, rows = 3:(3+nrow(data1)), 
           cols = 1:ncol(data1), gridExpand = TRUE, stack = TRUE)
  setColWidths(wb, "Inpatient Admissions", cols = 1:ncol(data1), 
               widths = "auto")
  
  
  ### Deaths ----
  names(data2) <- records_deaths
  addWorksheet(wb, sheetName = "Deaths", gridLines = FALSE)
  writeDataTable(wb, "Deaths", data2, startRow = 3)

  # titles
  title_deaths <- paste0("Deaths between", date_start, " and ", date_end)
  writeData(wb, "Deaths", title_deaths, startRow = 1, startCol = 1)
  addStyle(wb, "Deaths", title_style, rows = 1, cols = 1)
  
  # table headers
  addStyle(wb, "Deaths", title_style, rows = 3, cols = 1:ncol(data2))

  # tables
  addStyle(wb, "Deaths", table_style, rows = 3:(3+nrow(data2)), 
           cols = 1:ncol(data2), gridExpand = TRUE, stack = TRUE)
  setColWidths(wb, "Deaths", cols = 1:ncol(data2), widths = "auto")
  
  
  ### Save ----
  saveWorkbook(wb, paste0(wd, "/Output/Ruptured_AAA_audit_", hb_name, "_",
                          Sys.Date(), ".xlsx"), overwrite = TRUE)
}


### 2: SMR01 extract ----
# ## A: Call in extract ----
# # Create a connection to SMRA
# SMRA_connection <- odbc::dbConnect(
#   drv = odbc::odbc(),
#   dsn = "SMRA",
#   uid = rstudioapi::showPrompt(title = "Username", message = "Username:"),
#   pwd = rstudioapi::askForPassword("SMRA Password:")
# )
# 
# smr01_query <- tbl(SMRA_connection, "SMR01_PI") %>%
#   #names()
#   select(UPI_NUMBER, DERIVED_CHI, SURNAME, FIRST_FORENAME, POSTCODE, DOB, AGE_IN_YEARS,
#          AGE_IN_MONTHS, SEX, ADMISSION_DATE, DISCHARGE_DATE, MAIN_CONDITION, 
#          OTHER_CONDITION_1, OTHER_CONDITION_2, OTHER_CONDITION_3, OTHER_CONDITION_4,
#          OTHER_CONDITION_5, ADMISSION, DISCHARGE) %>%
#   filter(ADMISSION_DATE >= To_date('2021-06-01', 'YYYY-MM-DD'))
# 
# #smr01_query %>% show_query()
# 
# extract_01 <- collect(smr01_query)
# 
# ## Add in an output for the extract so don't have to connect to SMRA every time
# saveRDS(extract_01, paste0(wd, "/Temp/SMR01_extract.rds"))
# 
# rm(smr01_query)


## B: Refine extract ----
extract_01 <- readRDS(paste0(wd, "/Temp/SMR01_extract.rds")) |> 
  select(-DERIVED_CHI, -AGE_IN_MONTHS, -ADMISSION, -DISCHARGE)
names(extract_01)

# Rename variables
names <- c("upi", "surname", "forename", "postcode", "dob", 
           "age", "sex", "date_admission", "date_discharge",
           "main_condition", "other_condition_1", "other_condition_2", 
           "other_condition_3", "other_condition_4", "other_condition_5")
names(extract_01) <- names

# sex = male
# age = 65+
# ICD10 codes = icd_codes
# date_admission w/in dates _start & _end

# # Should this be for MAIN_CONDITION only or if any condition contains icd_codes?
# # MAIN_CONDITION only
# extract_01a <- extract_01 |> 
#   filter(AGE_IN_YEARS >= 65,
#          SEX == "1",
#          between(ADMISSION_DATE, date_start, date_end)) |> 
#   filter(MAIN_CONDITION %in% icd_codes)
#   
# table(extract_01a$MAIN_CONDITION)

## Include other_condition_x
extract_01 <- extract_01 |> 
  filter(age >= 65,
         sex == "1",
         between(date_admission, date_start, date_end)) |> 
  filter(main_condition %in% icd_codes |
           other_condition_1 %in% icd_codes |
           other_condition_2 %in% icd_codes |
           other_condition_3 %in% icd_codes |
           other_condition_4 %in% icd_codes |
           other_condition_5 %in% icd_codes) |> 
  mutate(surname = str_to_title(surname),
         forename = str_to_title(forename)) |> 
  glimpse()

# table(extract_01b$main_condition)
# table(extract_01b$other_condition_1)
# table(extract_01b$other_condition_2)
# table(extract_01b$other_condition_3)
# table(extract_01b$other_condition_4)
# table(extract_01b$other_condition_5)


## Add column that identifies if condition is main or other
# This is pretty messy, so happy for it to be done more efficiently if possible?
extract_01 <- extract_01 |> 
  mutate(condition = 
           case_when(main_condition %in% icd_codes ~ "main",
                     !(main_condition %in% icd_codes) & 
                       other_condition_1 %in% icd_codes ~ "other 1",
                     !(main_condition %in% icd_codes) & 
                       !(other_condition_1 %in% icd_codes) &
                       other_condition_2 %in% icd_codes ~ "other 2",
                     !(main_condition %in% icd_codes) & 
                       !(other_condition_1 %in% icd_codes) &
                       !(other_condition_2 %in% icd_codes) &
                       other_condition_3 %in% icd_codes ~ "other 3",
                     !(main_condition %in% icd_codes) & 
                       !(other_condition_1 %in% icd_codes) &
                       !(other_condition_2 %in% icd_codes) &
                       !(other_condition_3 %in% icd_codes) &
                       other_condition_4 %in% icd_codes ~ "other 4",
                     !(main_condition %in% icd_codes) & 
                       !(other_condition_1 %in% icd_codes) &
                       !(other_condition_2 %in% icd_codes) &
                       !(other_condition_3 %in% icd_codes) &
                       !(other_condition_3 %in% icd_codes) &
                       other_condition_5 %in% icd_codes ~ "other 5"),
         .after = date_discharge) |> 
  arrange(condition, date_admission)


## C: SIMD ----
# Ensure correct postcode format is used 
extract_01 <- extract_01 |> 
  mutate(postcode = format_postcode(postcode, "pc8"))

# Check for NAs
sum(is.na(extract_01$postcode))

# Combine
extract_simd <- left_join(extract_01, simd, 
                          by = c("postcode" = "pc8")) |> 
  select(hb2019name, upi:other_condition_5) |> 
  arrange(hb2019name, upi, date_admission)

names(extract_simd)

# Check for NAs
table(extract_simd$hb2019name, useNA = "ifany") # 20 NAs
check <- extract_simd[is.na(extract_simd$hb2019name),]
table(check$postcode, useNA = "ifany") # 8 postcodes in England/Wales
rm(check)

extract_simd <- extract_simd |> 
  filter(!is.na(hb2019name)) |> 
  glimpse()
table(extract_simd$hb2019name, useNA = "ifany")

# Keep only last entry for each UPI
length(unique(extract_simd$upi))

# number of observations should match above
extract_simd <- extract_simd |> 
  group_by(upi) |> 
  slice(n()) |> 
  ungroup() |> 
  glimpse()

rm(extract_01)


### 3: Deaths extract ----
# ## A: Call in extract ----
# # Create a connection to SMRA
# SMRA_connection <- odbc::dbConnect(
#   drv = odbc::odbc(),
#   dsn = "SMRA",
#   uid = rstudioapi::showPrompt(title = "Username", message = "Username:"),
#   pwd = rstudioapi::askForPassword("SMRA Password:")
# )
# 
# deaths_query <- tbl(SMRA_connection,
#     dbplyr::in_schema("ANALYSIS", "GRO_DEATHS_C")) %>%
#   #names()
#   select(UPI_NUMBER, CHI, DECEASED_SURNAME, DECEASED_FORENAME, POSTCODE,
#          DATE_OF_BIRTH, AGE, SEX, DATE_OF_DEATH, UNDERLYING_CAUSE_OF_DEATH,
#          CAUSE_OF_DEATH_CODE_0, CAUSE_OF_DEATH_CODE_1, CAUSE_OF_DEATH_CODE_2,
#          CAUSE_OF_DEATH_CODE_3, CAUSE_OF_DEATH_CODE_4, CAUSE_OF_DEATH_CODE_5,
#          CAUSE_OF_DEATH_CODE_6, CAUSE_OF_DEATH_CODE_7, CAUSE_OF_DEATH_CODE_8,
#          CAUSE_OF_DEATH_CODE_9) %>%
#   filter(DATE_OF_DEATH >= To_date("2021-06-01", "YYYY-MM-DD"))
# 
# # Use colnames to check variable names
# colnames(tbl(
#   SMRA_connection,
#   dbplyr::in_schema("ANALYSIS", "GRO_DEATHS_C")
# ))
# 
# # See what the SQL looks like
# deaths_query %>% show_query()
# 
# deaths <- collect(deaths_query)
# 
# ## Add in an output for the extract so don't have to connect to SMRA every time
# saveRDS(deaths, paste0(wd, "/Temp/deaths_extract.rds"))
# 
# rm(SMRA_connection, deaths_query)


## B: Refine extract ----
deaths <- readRDS(paste0(wd, "/Temp/deaths_extract.rds")) |> 
  select(-CHI)
names(deaths)

# Rename variables
names <- c("upi", "surname", "forename", "postcode", "dob", 
           "age", "sex", "date_death", "underlying_cause_death",
           "cause_death_0", "cause_death_1", "cause_death_2", "cause_death_3", 
           "cause_death_4", "cause_death_5", "cause_death_6", "cause_death_7",
           "cause_death_8", "cause_death_9")
names(deaths) <- names
  

# sex = male
# age = 65+
# ICD10 codes = icd_codes
# date_death w/in dates _start & _end
deaths <- deaths |> 
  filter(age >= 65,
         sex == "1",
         between(date_death, date_start, date_end)) |> 
  filter(underlying_cause_death %in% icd_codes |
           cause_death_0 %in% icd_codes |
           cause_death_1 %in% icd_codes |
           cause_death_2 %in% icd_codes |
           cause_death_3 %in% icd_codes |
           cause_death_4 %in% icd_codes |
           cause_death_5 %in% icd_codes |
           cause_death_6 %in% icd_codes |
           cause_death_7 %in% icd_codes |
           cause_death_8 %in% icd_codes |
           cause_death_9 %in% icd_codes) |> 
  mutate(surname = str_to_title(surname),
         forename = str_to_title(forename)) |> 
  glimpse()

# table(deaths$underlying_cause_death)
# table(deaths$cause_death_0)
# table(deaths$cause_death_1)
# table(deaths$cause_death_2)
# table(deaths$cause_death_3)
# table(deaths$cause_death_4)
# table(deaths$cause_death_5)
# table(deaths$cause_death_6)
# table(deaths$cause_death_7) # no ruptured AAA codes
# table(deaths$cause_death_8) # no ruptured AAA codes
# table(deaths$cause_death_9) # no ruptured AAA codes


## Add column that identifies if cause of death is underlying or other
# Again, pretty messy, so happy for it to be done more efficiently...
# Also, is this actually helpful?
cause_variables <- c("underlying", "cause 0", "cause 1", 
                     "cause 2", "cause 3", "cause 4")

deaths <- deaths |> 
  # ID underlying, causes 0,1,2,3,4
  mutate(cause_fatal = 
           case_when(underlying_cause_death %in% icd_codes ~ "underlying",
                     !(underlying_cause_death %in% icd_codes) & 
                       cause_death_0 %in% icd_codes ~ "cause 0",
                     !(underlying_cause_death %in% icd_codes) & 
                       !(cause_death_0 %in% icd_codes) &
                       cause_death_1 %in% icd_codes ~ "cause 1",
                     !(underlying_cause_death %in% icd_codes) & 
                       !(cause_death_0 %in% icd_codes) &
                       !(cause_death_1 %in% icd_codes) &
                       cause_death_2 %in% icd_codes ~ "cause 2",
                     !(underlying_cause_death %in% icd_codes) & 
                       !(cause_death_0 %in% icd_codes) &
                       !(cause_death_1 %in% icd_codes) &
                       !(cause_death_2 %in% icd_codes) &
                       cause_death_3 %in% icd_codes ~ "cause 3",
                     !(underlying_cause_death %in% icd_codes) & 
                       !(cause_death_0 %in% icd_codes) &
                       !(cause_death_1 %in% icd_codes) &
                       !(cause_death_2 %in% icd_codes) &
                       !(cause_death_3 %in% icd_codes) &
                       cause_death_4 %in% icd_codes ~ "cause 4"),
         .after = date_death) |> 
  # ID remaining causes 5,6,7,8,9
  mutate(cause_fatal = case_when(cause_fatal %in% cause_variables ~ cause_fatal,
                                 !(cause_fatal %in% cause_variables) &
                                   cause_death_5 %in% icd_codes ~ "cause 5",
                                 !(cause_fatal %in% cause_variables) &
                                   !(cause_death_5 %in% icd_codes) &
                                   cause_death_6 %in% icd_codes ~ "cause 6",
                                 !(cause_fatal %in% cause_variables) &
                                   !(cause_death_5 %in% icd_codes) &
                                   !(cause_death_6 %in% icd_codes) &
                                   cause_death_7 %in% icd_codes ~ "cause 7",
                                 !(cause_fatal %in% cause_variables) &
                                   !(cause_death_5 %in% icd_codes) &
                                   !(cause_death_6 %in% icd_codes) &
                                   !(cause_death_7 %in% icd_codes) &
                                   cause_death_8 %in% icd_codes ~ "cause 8",
                                 !(cause_fatal %in% cause_variables) &
                                   !(cause_death_5 %in% icd_codes) &
                                   !(cause_death_6 %in% icd_codes) &
                                   !(cause_death_7 %in% icd_codes) &
                                   !(cause_death_8 %in% icd_codes) &
                                   cause_death_9 %in% icd_codes ~ "cause 9")) |> 
  arrange(cause_fatal, date_death)


## C: SIMD ----
# Ensure correct postcode format is used 
deaths <- deaths |> 
  mutate(postcode = format_postcode(postcode, "pc8"))

# Check how many NAs
sum(is.na(deaths$postcode)) # checked single record & has NA in SQL deaths db

# Combine
deaths_simd <- left_join(deaths, simd, 
                          by = c("postcode" = "pc8")) |> 
  select(hb2019name, upi:cause_death_9) |> 
  arrange(hb2019name, upi, date_death)

names(deaths_simd)

# Check for NAs
table(deaths_simd$hb2019name, useNA = "ifany") # just the 1 from above

deaths_simd <- deaths_simd |> 
  filter(!is.na(hb2019name)) |> 
  glimpse()
table(deaths_simd$hb2019name, useNA = "ifany")

# Check no duplicate UPIs
length(unique(deaths_simd$upi)) 
# ideally, this should match number of observations; if not, investigate.


rm(deaths, cause_variables, names, simd_path)


### 5: Output to Excel ----
hb_names <- extract_simd |> 
  distinct(hb2019name) |> 
  pull()

for (hb_name in hb_names) {
  
  write_report(extract_simd, deaths_simd, hb_name)
  
}
