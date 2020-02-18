library(pdftools)
library(readxl)
library(dplyr)
library(stringr)
library(stringdist)
library(fuzzywuzzyR)
library(purrr)

location <- "home" #set "home", "laptop", "work" or "work-lap"

if (location == "home") {
  oned_path <- "D:/OneDrive - Aalborg Universitet/CALDISS_projects/"

} else if (location == "work") {
  oned_path <- "D:/OneDrive/OneDrive - Aalborg Universitet/CALDISS_projects/"
  
} else if (location == "work-lap") {
  oned_path <- "C:/Users/kgk/OneDrive - Aalborg Universitet/CALDISS_projects/"
  
} else {
  print("Specify location")
}

#Filepaths
data_path <- paste0(oned_path, "eu_organization-class_dps_F20/data/")
mat_path <- paste0(oned_path, "eu_organization-class_dps_F20/materials/")
out_path <- paste0(oned_path, "eu_organization-class_dps_F20/output/")

#read organisations overview
org1_df <- read_excel(paste0(mat_path, "eu-transp-reg_orgs-cat1.xls"))
org2_df <- read_excel(paste0(mat_path, "eu-transp-reg_orgs-cat2.xls"))
org3_df <- read_excel(paste0(mat_path, "eu-transp-reg_orgs-cat3.xls"))
org4_df <- read_excel(paste0(mat_path, "eu-transp-reg_orgs-cat4.xls"))
org5_df <- read_excel(paste0(mat_path, "eu-transp-reg_orgs-cat5.xls"))
org6_df <- read_excel(paste0(mat_path, "eu-transp-reg_orgs-cat6.xls"))

allorgs_df <- union_all(org1_df, org2_df) %>%
  union_all(org3_df) %>%
  union_all(org4_df) %>%
  union_all(org5_df) %>%
  union_all(org6_df)

colnames(allorgs_df)[which(colnames(allorgs_df) == "(Organisation) name")] <- "organisation_name"

#read list of participants
part_list <- pdf_text((paste0(data_path, "List-of-participants10.pdf")))
part_list <- paste(part_list, collapse = "")
orgs <- unlist(str_extract_all(part_list, "(?<=\\w\\s{5,50})\\S+?(\\s\\w+)*\\w(?=\r\n)"))

match_org <- function(org, orglist = unlist(str_to_lower(orgs))) {
  org = tolower(org)
  match = GetCloseMatches(string = org, sequence_strings = orglist, n = 1, cutoff = 0.85)
  match = unlist(match)
  match = ifelse(is.null(match), NA, match)
  return(match)
}

allorgs_matched <- allorgs_df
allorgs_matched$org_match <- map_chr(allorgs_matched$organisation_name, match_org)

allorgs_matched <- allorgs_matched %>%
  select(organisation_name, org_match, Section) %>%
  filter(!(is.na(org_match)))

setwd(out_path)
write.table(allorgs_matched, "org_match.csv", sep = ";", col.names = TRUE, row.names = FALSE)
