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
# either the 4th or 5th conditions.
# 
# Mortality Data
# National Records of Scotland (NRS) death records will be searched for any males 
# who have presented with an abdominal aortic aneurysm, with or without rupture 
# (ICD10 codes I71.3, I71.4, I71.5, I71.6, I71.8, and I71.9). Individuals aged 
# older than 65 on 1 January 2012 are to be excluded as they would not have been
# part of the AAA screening programme. The remaining records are then matched to 
# patient AAA records to identify date of vascular surgery in order to identify
# if the individual died within 30 days of their surgery.
# 
# CHI number will be used to link all datasets, keeping only one record per 
# patient. In order to help screening units review data, the search will be 
# carried out for hospital admissions and/or deaths following the calendar year 
# (January - December).


# install.packages("odbc")
# install.packages("dbplyr")
# library(odbc)
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
date_start <- dmy("01012022") # 1 January
date_end <- dmy("31122022") # 31 December
icd_rupture_codes <- c("I713", "I715", "I718") # inpatients
icd_codes <- c("I713", "I714", "I715", "I716", "I718", "I719") # deaths

extract <- 202309 # the September extract for the year following the year of focus

## Pathways
wd_path <- paste0("/PHI_conf/AAA/Topics/Projects/20210929-Ruptured-AAAs-Audit")

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
  # source <- paste0("Data for hospital admissions has been produced from the ",
  #                  "national inpatients dataset held by PHS (known as SMR01). ",
  #                  "Data containing information on deaths has been produced from ",
  #                  "the National Records of Scotland (NRS) death records.")
  # source2 <- paste0("SMR01 and NRS datasets were searched for any males, of ", 
  #                   "age 65 and older, who have presented with a ruptured ",
  #                   "aneurysm (ICD10 codes I71.3, I71.4, I71.5, I71.6, ",
  #                   "I71.8, and I71.9).")
  today <- paste0("Workbook created ", Sys.Date())
  
  ## Data
  # Inpatients (extract_simd)
  data1 <- df1 |> 
    filter(hb2019name == hb_name)
  # Deaths (deaths_simd)
  data2 <- df2 |> 
    filter(hb2019name == hb_name)
  
  
  ## Setup workbook
  # wb <- createWorkbook()
  wb <- loadWorkbook(template)
  options("openxlsx.borderStyle" = "thin",
          "openxlsx.dateFormat" = "dd/mm/yyyy")
  modifyBaseFont(wb, fontSize = 12, fontName = "Arial")
  
  
  ### Notes ----
  #addWorksheet(wb, sheetName = "Notes", gridLines = FALSE)
  writeData(wb, "Notes", title, startRow = 1, startCol = 1)
  writeData(wb, "Notes", date_range, startRow = 2, startCol = 1)
  # writeData(wb, "Notes", source, startRow = 4, startCol = 1)
  # writeData(wb, "Notes", source2, startRow = 5, startCol = 1)
  writeData(wb, "Notes", today, startRow = 7, startCol = 1)
  
  setColWidths(wb, "Notes", cols = 1, widths = "100.00")
  #setRowHeights(wb, "Notes", rows = 4:5, height = 30)
  addStyle(wb, "Notes", title_style, rows = 1, cols = 1)
  # addStyle(wb, "Notes", wrap_style, rows = 4:5, cols = 1, gridExpand = TRUE)
  
  
  ### Inpatient Admissions ----
  # names(data1) <- records_extract
  # addWorksheet(wb, sheetName = "Inpatient Admissions", gridLines = FALSE)
  writeData(wb, "Inpatient Admissions", data1, startRow = 2, colNames = FALSE)
  
  # # titles
  # title_extract <- paste0("Inpatient admissions between ", date_start, 
  #                         " and ", date_end)
  # writeData(wb, "Inpatient Admissions", title_extract, startRow = 1, startCol = 1)
  # addStyle(wb, "Inpatient Admissions", title_style, rows = 1, cols = 1)
  
  # table headers
  addStyle(wb, "Inpatient Admissions", title_style, rows = 1, cols = 1:ncol(data1))
  
  # tables
  addStyle(wb, "Inpatient Admissions", table_style, rows = 2:(2+nrow(data1)), 
           cols = 1:ncol(data1), gridExpand = TRUE, stack = TRUE)
  setColWidths(wb, "Inpatient Admissions", cols = 1:ncol(data1), 
               widths = "auto")
  
  
  ### Deaths ----
  # names(data2) <- records_deaths
  # addWorksheet(wb, sheetName = "Deaths", gridLines = FALSE)
  writeData(wb, "Deaths", data2, startRow = 2, colNames = FALSE)
  
  # # titles
  # title_deaths <- paste0("Deaths between", date_start, " and ", date_end)
  # writeData(wb, "Deaths", title_deaths, startRow = 1, startCol = 1)
  # addStyle(wb, "Deaths", title_style, rows = 1, cols = 1)
   
  # table headers
  addStyle(wb, "Deaths", title_style, rows = 1, cols = 1:ncol(data2))
  
  # tables
  addStyle(wb, "Deaths", table_style, rows = 2:(2+nrow(data2)), 
           cols = 1:ncol(data2), gridExpand = TRUE, stack = TRUE)
  setColWidths(wb, "Deaths", cols = 1:ncol(data2), widths = "auto")
  
  
  ### Save ----
  saveWorkbook(wb, paste0(wd_path, "/Output/Ruptured_AAA_audit_", hb_name, "_",
                          Sys.Date(), ".xlsx"), overwrite = TRUE)
}


### 2: SMR01 extract ----
## A: Call in extract ----
# Create a connection to SMRA
SMRA_connection <- odbc::dbConnect(
  drv = odbc::odbc(),
  dsn = "SMRA",
  uid = rstudioapi::showPrompt(title = "Username", message = "Username:"),
  pwd = rstudioapi::askForPassword("SMRA Password:")
)

smr01_query <- tbl(SMRA_connection, "SMR01_PI") %>%
  colnames()
#   select(UPI_NUMBER, DERIVED_CHI, SURNAME, FIRST_FORENAME, POSTCODE, DOB, AGE_IN_YEARS,
#          AGE_IN_MONTHS, SEX, ADMISSION_DATE, DISCHARGE_DATE, MAIN_CONDITION,
#          OTHER_CONDITION_1, OTHER_CONDITION_2, OTHER_CONDITION_3, OTHER_CONDITION_4,
#          OTHER_CONDITION_5, ADMISSION, DISCHARGE) %>%
#   filter(ADMISSION_DATE >= To_date('2022-01-01', 'YYYY-MM-DD'))
# 
#
# #smr01_query %>% show_query()
# 
# inpatient <- collect(smr01_query)
# 
# ## Add in an output for the extract so don't have to connect to SMRA every time
# saveRDS(inpatient, paste0(wd_path, "/Temp/SMR01_extract.rds"))
# 
# rm(smr01_query)


## B: Refine extract ----
inpatient <- readRDS(paste0(wd_path, "/Temp/SMR01_extract.rds")) |> 
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
table(inpatient_simd$hb2019name, useNA = "ifany") # 8 NAs
check <- inpatient_simd[is.na(inpatient_simd$hb2019name),]
table(check$postcode, useNA = "ifany") # 1 postcode in England/Wales
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


## D: Surgical data ----
# Pull in complete extract to identify individuals in the AAA program 
aaa_extract <- readRDS(extract_path) |> 
  select(financial_year, upi, dob, hbres, date_surgery, 
         hb_surgery, date_death, aneurysm_related) |> 
  mutate(in_aaa_program = "Yes")


table(aaa_extract$financial_year, useNA = "ifany")
table(aaa_extract$date_surgery, useNA = "ifany")
table(aaa_extract$date_death, useNA = "ifany")


# Check if individuals had surgery or death (death needs to be >30 surgery)
# Also check that dob and HB match
inpatient_matched <- left_join(inpatient_simd, aaa_extract,
                               by = "upi") |> 
  select(financial_year, hb2019name, hbres, upi, dob.x, dob.y,
         date_admission, date_surgery, date_death, aneurysm_related, 
         surname:postcode, age, sex, in_aaa_program, date_discharge:notes) |> 
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
         dob.x = if_else(!is.na(dob.x), dob.x, dob.y)) |> 
  select(hb2019name, upi, surname:postcode, dob.x, age:in_aaa_program, date_admission, 
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

rm(inpatient, inpatient_simd)

fife <- inpatient_matched[inpatient_matched$hb2019name == "NHS Fife",]

### 3: Deaths extract ----
## A: Call in extract ----
# Create a connection to SMRA
# SMRA_connection <- odbc::dbConnect(
#   drv = odbc::odbc(),
#   dsn = "SMRA",
#   uid = rstudioapi::showPrompt(title = "Username", message = "Username:"),
#   pwd = rstudioapi::askForPassword("SMRA Password:")
# )

# deaths_query <- tbl(SMRA_connection,
#     dbplyr::in_schema("ANALYSIS", "GRO_DEATHS_C")) %>%
#   #names()
#   select(UPI_NUMBER, CHI, DECEASED_SURNAME, DECEASED_FORENAME, POSTCODE,
#          DATE_OF_BIRTH, AGE, SEX, DATE_OF_DEATH, UNDERLYING_CAUSE_OF_DEATH,
#          CAUSE_OF_DEATH_CODE_0, CAUSE_OF_DEATH_CODE_1, CAUSE_OF_DEATH_CODE_2,
#          CAUSE_OF_DEATH_CODE_3, CAUSE_OF_DEATH_CODE_4, CAUSE_OF_DEATH_CODE_5,
#          CAUSE_OF_DEATH_CODE_6, CAUSE_OF_DEATH_CODE_7, CAUSE_OF_DEATH_CODE_8,
#          CAUSE_OF_DEATH_CODE_9) %>%
#   filter(DATE_OF_DEATH >= To_date("2022-01-01", "YYYY-MM-DD"))
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
# saveRDS(deaths, paste0(wd_path, "/Temp/deaths_extract.rds"))
# 
# rm(SMRA_connection, deaths_query)


## B: Refine extract ----
deaths <- readRDS(paste0(wd_path, "/Temp/deaths_extract.rds")) |> 
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
 
table(deaths$underlying_cause_death)
table(deaths$cause_death_0)
table(deaths$cause_death_1)
table(deaths$cause_death_2)
table(deaths$cause_death_3)
table(deaths$cause_death_4)
table(deaths$cause_death_5)
table(deaths$cause_death_6)
table(deaths$cause_death_7) 
table(deaths$cause_death_8) # no ruptured AAA codes
table(deaths$cause_death_9) # no ruptured AAA codes


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
table(deaths_simd$hb2019name, useNA = "ifany") 

# Check no duplicate UPIs
length(unique(deaths_simd$upi)) 
# ideally, this should match number of observations; if not, investigate.

rm(deaths, cause_variables, names, simd_path)


## D: Surgical data ----
# Pull in complete extract to identify individuals who died within 30 days of surgery 
aaa_extract <- readRDS(extract_path) |> 
  select(financial_year, upi, dob, hbres, date_surgery, 
         hb_surgery, date_death, aneurysm_related) |> 
  mutate(in_aaa_program = "Yes") |> 
  mutate(year = year(date_death)) |> 
  filter(year == 2022)


table(aaa_extract$financial_year, useNA = "ifany")
table(aaa_extract$date_death, useNA = "ifany")
table(aaa_extract$year, useNA = "ifany")


## Check if individuals had surgery or death (death needs to be >30 days after surgery)
# Also check that dob, date of death, and HB match
deaths_matched <- full_join(deaths_simd, aaa_extract,
                            by = "upi") |> 
  select(financial_year, year, hb2019name, hbres, upi, dob.x, dob.y, 
         date_surgery, date_death.x, date_death.y, aneurysm_related,
         surname:postcode, age, sex, in_aaa_program,
         surgery_type, cause_fatal:notes) 

# Create age at death variable
# Select final columns for outputting to Excel
deaths_matched <- deaths_matched |> 
  mutate(hb2019name = if_else(!is.na(hb2019name), hb2019name, paste0("NHS ", hbres)),
         dob.x = if_else(!is.na(dob.x), dob.x, dob.y),
         date_death.x = if_else(!is.na(date_death.x), date_death.x, date_death.y)) |> 
  select(hb2019name, upi, surname:postcode, dob.x, age, sex:surgery_type,
         date_surgery, date_death.x, cause_fatal:notes) |> 
  mutate(days_to_death = day(as.period(interval(start = date_surgery, 
                                              end = date_death.x))), .after = date_death.x) |>
  filter(days_to_death < 30 | is.na(days_to_death)) |> 
  arrange(hb2019name, cause_fatal, upi) |> 
  # prep dates for writing out to Excel
  mutate(dob.x = as.character(dob.x),
         date_surgery = as.character(date_surgery),
         date_death.x = as.character(date_death.x))

table(deaths_matched$hb2019name, useNA = "ifany")

rm(deaths_simd)

fife <- deaths_matched[deaths_matched$hb2019name == "NHS Fife",]


### 4: Output to Excel ----
## Create checking workbooks to send to NHS Fife & Tayside collective
fife_tay <- c("NHS Fife", "NHS Tayside")

for (hb_name in fife_tay) {
  
  write_report(inpatient_matched, deaths_matched, hb_name)
  
}


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

