#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# ruptured_AAA_audit.R
# Karen Hotopp
# June 2023
# 
# Ruptured aneurysm audit
# 
# Written/run on Posit WB
# R version 4.1.2
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Methods
# 
# Inpatient Data
# The national inpatients dataset held by PHS (known as SMR01) will be searched 
# for any males, of age 65 and older, who have presented with a ruptured 
# aneurysm (ICD10 codes I71.3, I71.5, and I71.8), in order to remove elective
# surgeries from the search. Records are omitted if the aneurysm is listed as 
# either the 4th or 5th conditions. *Date of death is taken from AAA extract.
# 
# Mortality Data
# National Records of Scotland (NRS) death records will be searched for any males 
# who have presented with an abdominal aortic aneurysm, with or without rupture 
# (ICD10 codes I71.3, I71.4, I71.5, I71.6, I71.8, and I71.9). Individuals aged 
# older than 65 on 1 January 2012 are to be excluded as they would not have been
# part of the AAA screening programme. The remaining records are then matched to 
# patient AAA records to identify date of vascular surgery in order to identify
# if the individual died within 30 days of their surgery. *Date of death is taken 
# from the NRS death records and is supplemented by the AAA extract, where needed
# and possible.
# 
# CHI number will be used to link all datasets, keeping only one record per 
# patient. In order to help screening units review data, the search will be 
# carried out for hospital admissions and/or deaths following the calendar year 
# (January - December).


# install.packages("odbc")
# install.packages("dbplyr")
library(odbc)
library(dplyr)
library(lubridate)
library(stringr)
library(phsmethods)
library(forcats)
library(openxlsx)
library(tidylog)


rm(list=ls())
gc()


### 1: Housekeeping ----
## Variables
year = 2021
date_start <- dmy("01012021") # 1 January
date_end <- dmy("31122021") # 31 December
icd_rupture_codes <- c("I713", "I715", "I718") # inpatients
icd_codes <- c("I713", "I714", "I715", "I716", "I718", "I719") # deaths

extract <- 202209 # the September extract for the year following the year of focus

## Pathways
wd_path <- paste0("/PHI_conf/AAA/Topics/Vascular/Projects/Ruptured AAA Audit")

simd_path <- paste0("/conf/linkage/output/lookups/Unicode/Deprivation",
                    "/postcode_2024_2_simd2020v2.rds")

extract_path <- paste0("/PHI_conf/AAA/Topics/Screening/extracts/",
                       extract, "/output/aaa_extract_", extract, ".rds")

template <- paste0(wd_path, "/Ruptured_AAA_audit_template_YYYY.xlsx")

## SIMD data
simd <- readRDS(simd_path) |> 
  select(pc8, hb2019name)


## Function
write_report <- function(df1, df2, hb_name) {
  
  ### Setup workbook ----
  
  ## Styles
  title_style <- createStyle(fontSize = 14, halign = "Left", textDecoration = "bold")
  table_style <- createStyle(valign = "Bottom", halign = "Left",
                             border = "TopBottomLeftRight")
  #wrap_style <- createStyle(wrapText = TRUE)
  
  ## Titles
  title <- paste0("Ruptured Anuerysm Audit for ", hb_name)
  date_range <- paste0("Hospital admissions and deaths between ", date_start,
                       " and ", date_end)
  today <- paste0("Workbook created ", Sys.Date())
  
  ## Data
  # Inpatients (inpatient_matched)
  data1 <- df1 |> 
    filter(hb2019name == hb_name)
  # Deaths (deaths_matched)
  data2 <- df2 |> 
    filter(hb2019name == hb_name)
  
  
  ## Setup workbook
  wb <- loadWorkbook(template)
  options("openxlsx.borderStyle" = "thin",
          "openxlsx.dateFormat" = "dd/mm/yyyy")
  modifyBaseFont(wb, fontSize = 12, fontName = "Arial")
  
  
  ### Notes ----
  writeData(wb, "Notes", title, startRow = 1, startCol = 1)
  writeData(wb, "Notes", date_range, startRow = 2, startCol = 1)
  writeData(wb, "Notes", today, startRow = 9, startCol = 1)
  
  setColWidths(wb, "Notes", cols = 1, widths = "100.00")
  addStyle(wb, "Notes", title_style, rows = 1, cols = 1)
  
  
  ### Inpatient Admissions ----
  writeData(wb, "Inpatient Admissions", data1, startRow = 2, colNames = FALSE)
  
  # table headers
  addStyle(wb, "Inpatient Admissions", title_style, rows = 1, cols = 1:ncol(data1))
  
  # tables
  addStyle(wb, "Inpatient Admissions", table_style, rows = 2:(2+nrow(data1)), 
           cols = 1:ncol(data1), gridExpand = TRUE, stack = TRUE)
  setColWidths(wb, "Inpatient Admissions", cols = 1:ncol(data1), 
               widths = "auto")
  
  
  ### Deaths ----
  writeData(wb, "Deaths", data2, startRow = 2, colNames = FALSE)
  
  # table headers
  addStyle(wb, "Deaths", title_style, rows = 1, cols = 1:ncol(data2))
  
  # tables
  addStyle(wb, "Deaths", table_style, rows = 2:(2+nrow(data2)), 
           cols = 1:ncol(data2), gridExpand = TRUE, stack = TRUE)
  setColWidths(wb, "Deaths", cols = 1:ncol(data2), widths = "auto")
  
  
  ### Save ----
  saveWorkbook(wb, paste0(wd_path, "/Output/2021 data/Ruptured_AAA_audit_", hb_name, "_",
                          year, ".xlsx"), overwrite = TRUE)
}


### 2: SMR01 extract ----
## A: Call in extract ----
# # Create a connection to SMRA
# SMRA_connection <- odbc::dbConnect(
#   drv = odbc::odbc(),
#   dsn = "SMRA",
#   uid = rstudioapi::showPrompt(title = "Username", message = "Username:"),
#   pwd = rstudioapi::askForPassword("SMRA Password:")
# )
# 
# smr01_query <- tbl(SMRA_connection, "SMR01_PI") %>%
#   #colnames()
#   select(UPI_NUMBER, DERIVED_CHI, SURNAME, FIRST_FORENAME, POSTCODE, DOB, AGE_IN_YEARS,
#          AGE_IN_MONTHS, SEX, ADMISSION_DATE, DISCHARGE_DATE, MAIN_CONDITION,
#          OTHER_CONDITION_1, OTHER_CONDITION_2, OTHER_CONDITION_3, OTHER_CONDITION_4,
#          OTHER_CONDITION_5, ADMISSION, DISCHARGE, LINK_NO, URI) %>%
#   filter(ADMISSION_DATE >= To_date('2021-01-01', 'YYYY-MM-DD')) %>%
#   arrange(LINK_NO, ADMISSION_DATE, DISCHARGE_DATE, ADMISSION, DISCHARGE, URI) %>%
#   select(-LINK_NO, -URI)
# 
# 
# #smr01_query %>% show_query()
# 
# inpatient <- collect(smr01_query)
# 
# ## Add in an output for the extract so don't have to connect to SMRA every time
# saveRDS(inpatient, paste0(wd_path, "/Temp/SMR01_extract_", year, ".rds"))
# 
# rm(smr01_query)


## B: Refine extract ----
inpatient <- readRDS(paste0(wd_path, "/Temp/SMR01_extract_", year, ".rds")) |> 
  select(-DERIVED_CHI, -AGE_IN_MONTHS, -ADMISSION, -DISCHARGE,
         -OTHER_CONDITION_4, -OTHER_CONDITION_5)
names(inpatient)

# Rename variables
names <- c("upi", "surname", "forename", "postcode", "dob", 
           "age", "sex", "date_admission", "date_discharge",
           "main_condition", "other_condition_1", "other_condition_2", 
           "other_condition_3")
names(inpatient) <- names

# sex = male
# age = 65+ and <65 on 1 January 2012
# ICD10 codes = icd_rupture_codes
# date_admission w/in dates _start & _end
# Include other_condition_ 1-3
inpatient <- inpatient |> 
  mutate(age_at_2012 = year(as.period(interval(start = dob, 
                                               end = dmy(01012012))))) |> 
  filter(sex == "1",
         age_at_2012 <= 65,
         age >= 65,
         between(date_admission, date_start, date_end)) |> 
  filter(main_condition %in% icd_rupture_codes |
           other_condition_1 %in% icd_rupture_codes |
           other_condition_2 %in% icd_rupture_codes |
           other_condition_3 %in% icd_rupture_codes) |> 
  mutate(surname = str_to_title(surname),
         forename = str_to_title(forename)) |> 
  glimpse()

table(inpatient$main_condition)
table(inpatient$other_condition_1)
table(inpatient$other_condition_2)
table(inpatient$other_condition_3)
range(inpatient$date_admission)


## Add column that identifies if condition is main or other
inpatient <- inpatient |> 
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
                       other_condition_3 %in% icd_codes ~ "other 3"),
         .after = date_discharge) |> 
  arrange(desc(condition), date_admission)

table(inpatient$condition)


## Add in response rows
inpatient <- inpatient |> 
  mutate(should_on_list = NA,
         why_on_list = NA,
         review_needed = NA,
         notes = NA) |> 
  glimpse()


## C: SIMD ----
# Ensure correct postcode format is used 
inpatient <- inpatient |> 
  mutate(postcode = format_postcode(postcode, "pc8"))

# Check for NAs
sum(is.na(inpatient$postcode))

# Combine
inpatient_simd <- left_join(inpatient, simd, 
                            by = c("postcode" = "pc8")) |> 
  select(hb2019name, upi:notes) |> 
  arrange(hb2019name, upi, date_admission)

names(inpatient_simd)

# Check for NAs
table(inpatient_simd$hb2019name, useNA = "ifany") # 1 NAs
check <- inpatient_simd[is.na(inpatient_simd$hb2019name),]
table(check$postcode, useNA = "ifany") # postcode doesn't exist, but most likely in England
rm(check)


## Keep only last entry for each UPI
length(unique(inpatient_simd$upi))

# number of observations should match above
inpatient_simd <- inpatient_simd |> 
  arrange(desc(condition), date_admission) |> # arrange records such that conditions run low to high
  group_by(upi) |> 
  slice(n()) |> # remaining record is one with highest condition level (i.e. main, etc.)
  ungroup() |> 
  glimpse()

table(inpatient_simd$hb2019name, useNA = "ifany")


## D: Surgical data ----
# Pull in complete extract to identify individuals in the AAA program 
aaa_extract <- readRDS(extract_path) |> 
  select(financial_year, upi, dob, postcode, hbres, date_surgery, 
         hb_surgery, date_death, aneurysm_related) |> 
  mutate(in_aaa_program = "Yes")

# Ensure correct postcode format is used 
aaa_extract <- aaa_extract |> 
  mutate(postcode = format_postcode(postcode, "pc8"))

table(aaa_extract$financial_year, useNA = "ifany")
table(aaa_extract$date_surgery, useNA = "ifany")
table(aaa_extract$date_death, useNA = "ifany")


# Check if individuals had surgery or death (death needs to be >30 days from surgery)
# Also check that dob and HB match
inpatient_matched <- left_join(inpatient_simd, aaa_extract,
                               by = "upi") |> 
  select(financial_year, hb2019name, hbres, upi, dob.x, dob.y, date_admission, 
         date_surgery, date_death, aneurysm_related, surname:postcode.x, 
         postcode.y, age, sex, in_aaa_program, date_discharge:notes) |> 
  mutate(days_to_death = day(as.period(interval(start = date_surgery, 
                                                end = date_death))), .after = date_death) |>
  filter(days_to_death >= 30 | is.na(days_to_death))

inpatient_matched <- inpatient_matched |> 
  arrange(desc(condition), date_admission) |> 
  group_by(upi) |> 
  slice(n()) |> 
  ungroup() 

inpatient_matched <- arrange(inpatient_matched, hb2019name, upi)


# Select final columns for outputting to Excel
inpatient_matched <- inpatient_matched |> 
  mutate(hb2019name = if_else(!is.na(hb2019name), hb2019name, paste0("NHS ", hbres)),
         dob.x = if_else(!is.na(dob.x), dob.x, dob.y),
         postcode.x = if_else(!is.na(postcode.x), postcode.x, postcode.y)) |> 
  select(hb2019name, upi, surname:postcode.x, dob.x, age:in_aaa_program, date_admission, 
         date_discharge:other_condition_3, date_surgery:aneurysm_related,
         should_on_list:notes) |> 
  arrange(hb2019name, upi) |> 
  # prep dates for writing out to Excel
  mutate(dob.x = as.character(dob.x),
         date_admission = as.character(date_admission),
         date_discharge = as.character(date_discharge),
         date_surgery = as.character(date_surgery),
         date_death = as.character(date_death))

table(inpatient_matched$hb2019name, useNA = "ifany")

saveRDS(inpatient_matched, paste0(wd_path, "/Temp/inpatient_matched_", year, ".rds"))
rm(inpatient, inpatient_simd)


### 3: Deaths extract ----
## A: Call in extract ----
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
#   filter(DATE_OF_DEATH >= To_date("2021-01-01", "YYYY-MM-DD"))
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
# saveRDS(deaths, paste0(wd_path, "/Temp/deaths_extract_", year, ".rds"))
# 
# rm(SMRA_connection, deaths_query)


## B: Refine extract ----
deaths <- readRDS(paste0(wd_path, "/Temp/deaths_extract_", year, ".rds")) |> 
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
# age = 65+ and <65 on 1 January 2012
# ICD10 codes = icd_codes
# date_death w/in dates _start & _end
deaths <- deaths |> 
  mutate(age_at_2012 = year(as.period(interval(start = dob, 
                                               end = dmy(01012012))))) |> 
  filter(sex == "1",
         age_at_2012 <= 65,
         age >= 65,
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


table(deaths$age_at_2012, useNA = "ifany")
table(deaths$age, useNA = "ifany")  

range(deaths$date_death)  

table(deaths$underlying_cause_death)
table(deaths$cause_death_0)
table(deaths$cause_death_1)
table(deaths$cause_death_2)
table(deaths$cause_death_3)
table(deaths$cause_death_4)
table(deaths$cause_death_5)
table(deaths$cause_death_6)
table(deaths$cause_death_7) # no ruptured AAA codes 
table(deaths$cause_death_8) # no ruptured AAA codes
table(deaths$cause_death_9)


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
  mutate(cause_fatal = fct_relevel(cause_fatal, "underlying")) |> 
  arrange(cause_fatal, date_death)

table(deaths$cause_fatal, useNA = "ifany")


## Add in response rows
deaths <- deaths |> 
  mutate(surgery_type = NA,
         should_be_on_list = NA,
         why_not_list = NA,
         review_needed = NA,
         report_QPMG = NA,
         notes = NA) |> 
  glimpse()


## C: SIMD ----
# Ensure correct postcode format is used 
deaths <- deaths |> 
  mutate(postcode = format_postcode(postcode, "pc8"))

# Check how many NAs
sum(is.na(deaths$postcode))

# Combine
deaths_simd <- left_join(deaths, simd, 
                         by = c("postcode" = "pc8")) |> 
  select(hb2019name, upi:sex, surgery_type, 
         date_death:cause_death_9, should_be_on_list:notes) |> 
  arrange(hb2019name, upi, date_death)

names(deaths_simd)

# Check for NAs
table(deaths_simd$hb2019name, useNA = "ifany") # 2 NAs
check <- deaths_simd[is.na(deaths_simd$hb2019name),]
table(check$postcode, useNA = "ifany") # no postcodes
rm(check)

# Check for NAs
table(deaths_simd$hb2019name, useNA = "ifany") 

# Check no duplicate UPIs
length(unique(deaths_simd$upi)) 
# ideally, this should match number of observations; if not, investigate.

rm(deaths, cause_variables, names, simd_path)


## D: Surgical data ----
# Pull in complete extract to identify individuals who died within 30 days of surgery 
aaa_extract <- readRDS(extract_path) |> 
  select(financial_year, upi, dob, postcode, hbres, date_surgery, 
         hb_surgery, date_death, aneurysm_related) |> 
  mutate(in_aaa_program = "Yes") |> 
  filter(financial_year %in% c("2020/21", "2021/22")) |> # update so that full year of investigation is covered
  # filter out anyone who died in a different year, but keep NAs
  mutate(year = year(date_death)) |> 
  filter(year %in% c(NA, 2021)) # update year to current investigation year (variable year not working??)

# Ensure correct postcode format is used 
aaa_extract <- aaa_extract |> 
  mutate(postcode = format_postcode(postcode, "pc8"))

table(aaa_extract$financial_year, useNA = "ifany")
table(aaa_extract$date_death, useNA = "ifany")
table(aaa_extract$year, useNA = "ifany")


## Check if individuals had surgery or death (death needs to be >30 days after surgery)
# Also check that dob, date of death, and HB match
deaths_matched <- full_join(deaths_simd, aaa_extract,
                            by = "upi") |> 
  select(financial_year, year, hb2019name, hbres, upi, dob.x, dob.y, 
         date_surgery, date_death.x, date_death.y, aneurysm_related,
         surname:postcode.x, postcode.y, age, sex, in_aaa_program,
         surgery_type, cause_fatal:notes) 

# Create age at death variable
# Select final columns for outputting to Excel
deaths_matched <- deaths_matched |> 
  # remove ampersand from aaa_extract HB names
  mutate(hbres = case_when(hbres == "Ayrshire & Arran" ~ "Ayrshire and Arran",
                           hbres == "Dumfries & Galloway" ~ "Dumfries and Galloway",
                           hbres == "Greater Glasgow & Clyde" ~ "Greater Glasgow and Clyde",
                           TRUE ~ hbres)) |> 
  # where NRS data is NA, fill with AAA extract information
  mutate(hb2019name = if_else(!is.na(hb2019name), hb2019name, paste0("NHS ", hbres)),
         dob.x = if_else(!is.na(dob.x), dob.x, dob.y),
         postcode.x = if_else(!is.na(postcode.x), postcode.x, postcode.y)) |> 
  # make special note where NRS is missing date of death but AAA is not
  mutate(notes = if_else(is.na(date_death.x) & !is.na(date_death.y), 
                         "Date of death missing from NRS, taken from AAA extract", 
                         NA),
         date_death.x = if_else(!is.na(date_death.x), date_death.x, date_death.y)) |> 
  # everyone should have a recorded date of death now, can remove NAs
  filter(!is.na(date_death.x)) |> 
  select(hb2019name, upi, surname:postcode.x, dob.x, age, sex:surgery_type,
         date_surgery, date_death.x, cause_fatal:notes) |> 
  mutate(days_to_death = day(as.period(interval(start = date_surgery, 
                                                end = date_death.x))), .after = date_death.x) |>
  filter(days_to_death < 30 | is.na(days_to_death)) |> 
  arrange(hb2019name, cause_fatal, upi) |> 
  # prep dates for writing out to Excel
  mutate(dob.x = as.character(dob.x),
         date_surgery = as.character(date_surgery),
         date_death.x = as.character(date_death.x))

deaths_matched <- deaths_matched |> 
  group_by(upi) |> 
  slice(n()) |> 
  ungroup() 

deaths_matched <- arrange(deaths_matched, hb2019name, upi)

table(deaths_matched$hb2019name, useNA = "ifany")

saveRDS(deaths_matched, paste0(wd_path, "/Temp/deaths_matched_", year, ".rds"))
rm(deaths_simd, simd, aaa_extract)


### 4: Note double entries ----
## Need to identify and make a note of any individual who is on both the 
## inpatients and the deaths list
duplicates <- inner_join(inpatient_matched, deaths_matched, by = "upi") |> 
  select(upi) |> 
  mutate(dups = "Patient is in both lists")

## Inpatients
inpatient_matched <- left_join(inpatient_matched, duplicates) |> 
  select(hb2019name, upi, dups, surname:notes)

## Deaths
deaths_matched <- left_join(deaths_matched, duplicates) |> 
  select(hb2019name, upi, dups, surname:notes)


### 5: Output to Excel ----
## Create checking workbooks to send to NHS Fife & Tayside collective
fife_tay <- c("NHS Fife", "NHS Tayside")
gram_os <- c("NHS Grampian", "NHS Orkney", "NHS Shetland")

for (hb_name in fife_tay) {
  
  write_report(inpatient_matched, deaths_matched, hb_name)
  
}

write_report(inpatient_matched, deaths_matched, "NHS Grampian")


# ## Create vector for workbooks to all HBs
# hb_names <- simd |>
#   distinct(hb2019name) |>
#   pull()
# 
# for (hb_name in hb_names) {
# 
#   write_report(inpatient_matched, deaths_matched, hb_name)
# 
# }



### 6: Find cases ----
aaa_extract <- readRDS(extract_path)

pick_1 <- aaa_extract[aaa_extract$upi == "",]
pick_2 <- aaa_extract[aaa_extract$upi == "",]

surname <- c("", "")

review <- rbind(pick_1, pick_2) |> 
  mutate(surname = surname,
         .after = upi) |> 
  select(-financial_quarter, -chi, -eligibility_period, 
         -age65_onstartdate, -over65_onstartdate, -dob_eligibility, 
         -ca2019, -simd2020v2_hb2019_quintile, -first_outcome)

names(review)

# Update variable responses
review <- review |> 
  mutate(pat_elig = case_when(pat_elig=="01" ~ "Eligible, in cohort",
                              pat_elig=="02" ~ "Eligible, under previous surveillance in NHS Highland/Western Isles prior to national screening programme ",
                              pat_elig=="03" ~ "Eligible, self-referral"),
         screen_type = case_when(screen_type=="01" ~ "Initial",
                                 screen_type=="02" ~ "Surveillance",
                                 screen_type=="03" ~ "QA initial",
                                 screen_type=="04" ~ "QA surveillance"),
         att_dna = case_when(att_dna=="01" ~ "Non-responder/DNA, initial",
                             att_dna=="02" ~ "Non-responder/DNA, surveillance",
                             att_dna=="03" ~ "Appointment cancelled, patient",
                             att_dna=="04" ~ "Appointment cancelled/postponed, healthcare",
                             att_dna=="05" ~ "Attended and seen"),
         screen_result = case_when(
           screen_result=="01" ~ "Positive (AAA >= 3.0cm)",
           screen_result=="02" ~ "Negative (AAA < 3.0cm)",
           screen_result=="03" ~ "Technical failure",
           screen_result=="04" ~ "Non-visualization, longitudinal/transverse plane",
           screen_result=="05" ~ "External positive (AAA >= 3.0cm)",
           screen_result=="06" ~ "External negative (AAA < 3.0cm)"),
         screen_exep = case_when(
           screen_result=="01" ~ "Declined screening during attendance",
           screen_result=="02" ~ "Aorta non-visualised, technical failure",
           screen_result=="03" ~ "Aorta non-visualised, physical barrier",
           screen_result=="04" ~ "Clinically unsuitable for portable screening",
           screen_result=="05" ~ "Incomplete measurements",
           screen_result=="06" ~ "Too many measurements"),
         followup_recom = case_when(
           followup_recom=="01" ~ "3 months",
           followup_recom=="02" ~ "12 months",
           followup_recom=="03" ~ "Discharge",
           followup_recom=="04" ~ "Refer to vascular",
           followup_recom=="05" ~ "Immediate recall",
           followup_recom=="06" ~ "No further recall"),
         result_outcome = case_when(
           result_outcome=="01" ~ "Declined vascular referral",
           result_outcome=="02" ~ "Referred in error: Vascular appointment not required",
           result_outcome=="03" ~ "DNA outpatien service: Self-discharge",
           result_outcome=="04" ~ "DNA outpatien service: Died w/in 10 working days of referral",
           result_outcome=="05" ~ "DNA outpatien service: Died more than 10 working days of referral",
           result_outcome=="06" ~ "Referred in error: As determined by vascular services",
           result_outcome=="07" ~ "Died before surgical assessment completed",
           result_outcome=="08" ~ "Unfit for surgery",
           result_outcome=="09" ~ "Refer to another specialty",
           result_outcome=="10" ~ "Awaiting further AAA growth",
           result_outcome=="11" ~ "Appropriate for surgery: Patient declined surgery",
           result_outcome=="12" ~ "Appropriate for surgery: Died before treatment",
           result_outcome=="13" ~ "Appropriate for surgery: Self-discharge",
           result_outcome=="14" ~ "Appropriate for surgery: Patient deferred surgery",
           result_outcome=="15" ~ "Appropriate for surgery: AAA repaired and survived 30 days",
           result_outcome=="16" ~ "Appropriate for surgery: Died w/in 30 days of treatment",
           result_outcome=="17" ~ "Appropriate for surgery: Final outcome pending",
           result_outcome=="18" ~ "Ongoing assessment by vascular",
           result_outcome=="19" ~ "Final outcome pending",
           result_outcome=="20" ~ "Other final outcome"),
         referral_error_manage = case_when(
           referral_error_manage=="01" ~ "Discharged",
           referral_error_manage=="02" ~ "Surveillance 3 months",
           referral_error_manage=="03" ~ "Surveillance 12 months"),
         hb_surgery = case_when(hb_surgery=="A" ~ "Ayrshire & Arran",
                                hb_surgery=="B" ~ "Borders",
                                hb_surgery=="F" ~ "Fife",
                                hb_surgery=="G" ~ "Greater Glasgow & Clyde",
                                hb_surgery=="H" ~ "Highland",
                                hb_surgery=="L" ~ "Lanarkshire",
                                hb_surgery=="N" ~ "Grampian",
                                hb_surgery=="R" ~ "Orkney",
                                hb_surgery=="S" ~ "Lothian",
                                hb_surgery=="T" ~ "Tayside",
                                hb_surgery=="V" ~ "Forth Valley",
                                hb_surgery=="W" ~ "Western Isles",
                                hb_surgery=="Y" ~ "Dumfries & Galloway",
                                hb_surgery=="Z" ~ "Shetland",
                                hb_surgery=="D" ~ "Cumbria"),
         surg_method = case_when(
           surg_method=="01" ~ "Endovascular surgery",
           surg_method=="02" ~ "Open surgery",
           surg_method=="03" ~ "Proceedure abandoned"),
         audit_flag = case_when(audit_flag=="01" ~ "Yes",
                                audit_flag=="02" ~ "No"),
         audit_result = case_when(audit_result=="01" ~ "Standard met",
                                  audit_result=="02" ~ "Standard not met"))

# Rename variables
names <- c("financial year", "quarter", "upi", "surname", "dob", "sex", "postcode",
           "practice code", "practice name", "eligibility", "HB of residence", 
           "HB of screen", "SIMD", "location code", "screen type", 
           "date offer sent", "date screened", "age at screening", "attendance", 
           "screening result", "screen exceptions", "follow-up recommendation", 
           "APL measurement", "APT measurement", "largest measurement", 
           "AAA size", "AAA size group", "date result", "result verified", 
           "date verified", "date referral generated", "date referral actual", 
           "date seen outpatient", "referral outcome", "referred in error management", 
           "date surgery", "financial year surgery", "HB of surgery", "surgical methiod", 
           "date death", "aneurysm related", "flagged for audit", "audit result", 
           "audit fail reason", "audit fail detail 1", "audit fail detail 2", 
           "audit fail detail 3", "audit fail detail 4", "audit fail detail 5", 
           "audit outcome", "audit batch fail", "audit batch outcome")

names(review) <- names


# Write out
write.xlsx(review, paste0(wd_path, 
                          "/Output/2022 data/Fife_and_Tayside_review_2022.xlsx"))

