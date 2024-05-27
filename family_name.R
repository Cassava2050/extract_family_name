# Load necessary libraries
library(readxl)
library(tidyverse)
library(openxlsx)
library(here)

# Function to load multiple sheets from an Excel file
# This function reads multiple sheets from an Excel file (.xlsx) 
# and imports them into a list of data frames. 
read_excel_sheets <- function(filename) {
  # Read all sheet names from the .xlsx file
  all_sheets <- readxl::excel_sheets(filename)
  
  # Import each sheet into a list of data frames using readxl::read_excel
  df_list <- lapply(all_sheets, function(sheet) readxl::read_excel(filename, sheet = sheet))
  
  # Assign sheet names as names of the list elements
  names(df_list) <- all_sheets
  
  # Break up the list and create a data frame for each sheet with names matching sheet names
  list2env(df_list, envir = .GlobalEnv)
}

# Execute the function to read Excel sheets
# Change the filename to the correct path of the Excel file
read_excel_sheets("./data/GS_Cohorts1_2_3.xlsx")

# Extract names of data frames stored in the environment
df_names <- ls(pattern = "^[0-9]")

# Convert the named data frames into a list
df_list <- lapply(df_names, get)

# Define a function to process each data frame
process_df <- function(df) {
  df %>%
    separate(accession_name, c("family", "offspring_code"), sep = "-", remove = FALSE) %>%
    mutate(family = str_replace_all(family, "[A-Za-z]$", "")) %>%
    relocate(family, .after = accession_name)
}

# Apply the processing function to each data frame in the list
processed_gs_cohorts <- map(df_list, process_df)

# Save the processed data frames
output_dir <- here::here("output/")
dir.create(output_dir, showWarnings = FALSE)

meta_file_name <- paste0(output_dir, "2023_family_name_", Sys.Date(), ".xlsx")
write.xlsx(processed_gs_cohorts, file = meta_file_name)
