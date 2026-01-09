#This code is inspired by the code developed by Teresa Schultz for her FSCI 2025 Workshop available here: https://github.com/schauch/OpenAlexRFSCI2025

#Install and include necessary packages. Install only needs to be run once.
install.packages("openalexR") #Package to query OpenAlex
install.packages("tidyverse") #Package to work with data more easily
install.packages("openxlsx2") #Package to export to Excel

#Load the necessary packages
library(openalexR)
library(tidyverse)
library(openxlsx2)

options(openalexR.mailto = "jason.friedman@usask.ca") #Add email address to use polite pool / Only needs to be done once

#Query String: Modify the variables below to adjust query
QueryEntity = "works" #Entity
QueryInstitution_OpenAlex_ID = "I32625721" #OpenAlex Institution ID is University of Saskatchewan. Must use capital I for filters to work.
QueryType = "article|review" #Type
QuerySourceType = "journal|conference" #Source Type
QueryStartDate = "2023-01-01" #Publication Start Date
QueryEndDate = "2023-12-31" #Publication End Date

InstitutionFilterString = paste0("https://openalex.org/", QueryInstitution_OpenAlex_ID) #Adds OpenAlex URL string to institution ID for filtering

#Filter string Add/remove column headers as desired. Columns listed below are removed from the results.
FilterColumns = c(
  'abstract', 
  'fwci', 
  'pdf_url', 
  'first_page', 
  'last_page', 
  'volume', 
  'issue', 
  'any_repository_has_fulltext', 
  'cited_by_api_url',
  'ids', 
  'referenced_works', 
  'related_works', 
  'concepts', 
  'counts_by_year',
  'topics',
  'keywords',
  'grants'
)

#This code uses the query variables above to call the OpenAlex API
Institutional_Works <- oa_fetch(
  entity = QueryEntity, #query type
  corresponding_institution_ids = QueryInstitution_OpenAlex_ID, 
  type = QueryType, #type of item
  primary_location.source.type = QuerySourceType, #limit to journal articles
  from_publication_date = QueryStartDate, #start range
  to_publication_date = QueryEndDate, #end range
#  mailto = oa_email(), #To use polite API if not set above
  verbose = TRUE,
#  options = list("data-version" = 1), #Toggle to use version 2/Walden of OpenAlex
)

#Records query information, time stamp, and warning in data frame for future Export
CurrentDate <- substr(Sys.time(),1,10)
QueryWarnings = names(last.warning) #Save warning message
QueryWarnings <- QueryWarnings[!QueryWarnings %in% c("Note: `oa_fetch` and `oa2df` now return new names for some columns in openalexR v2.0.0.\n      See NEWS.md for the list of changes.\n      Call `get_coverage()` to view the all updated columns and their original names in OpenAlex.\n\033[90mThis warning is displayed once every 8 hours.\033[39m")] #This code filters out a specific openalexr warning about column name changes.
if (length(QueryWarnings) == 0) {QueryWarnings = "No warnings"} #Enters "No warnings" if there are no warnings
QueryStructure <- c("Entity", "Institution ID", "Type", "Source Type", "Start Date", "End Date", "Timestamp", "Warnings") #Query labels
QueryValues <- c(QueryEntity, QueryInstitution_OpenAlex_ID, QueryType, QuerySourceType, QueryStartDate, QueryEndDate, CurrentDate, QueryWarnings) #Query values
Query <- data.frame(Request = QueryStructure, Value = QueryValues) #Create data frame for Query information

#Generate initial works list
Institutional_FilteredWorks <- Institutional_Works %>% select(-any_of(FilterColumns)) #Removes columns indicated above
Institutional_FilteredWorksOnly <- Institutional_FilteredWorks %>% select(-any_of(c("authorships","apc"))) #Removes nested data authorships/apc for a clean works list for export
Institutional_OAWorksOnly <- filter(Institutional_FilteredWorksOnly,is_oa == TRUE) #Generate OA works list
Institutional_GoldHybrid <- filter(Institutional_FilteredWorksOnly,oa_status == "gold" | oa_status == "hybrid") #Generate Gold/Hybrid only works list (possible APC paid)

#Generate initial authors list
Institutional_AllAuthors <- unnest(Institutional_FilteredWorks, authorships, names_sep = "_") #Unnests authorships so that all authors are listed
Institutional_AllAuthorsOnly <- Institutional_AllAuthors %>% select(-any_of(c("authorships_affiliations","apc"))) #Removes nested data affiliations and apc for a clean authors list for export

#Generate affiliations list
AllAuthorAffiliations <- unnest(Institutional_AllAuthors, authorships_affiliations, names_sep = "_") #Unnests affiliations so that all affiliations are listed
AllAuthorAffiliations_NoAPC <- AllAuthorAffiliations %>% select(!"apc") #Removes APC data for clean export
Institutional_AuthorsOnly <- filter(AllAuthorAffiliations_NoAPC, authorships_affiliations_id == paste0("https://openalex.org/", QueryInstitution_OpenAlex_ID)) #Filters affiliations to institution only
Institutional_GoldHybridAuthorsOnly <- filter(Institutional_AuthorsOnly, oa_status == "gold" | oa_status == "hybrid")

#Generate corresponding authors list
Corresponding_only <- filter(Institutional_AllAuthors, authorships_is_corresponding == "TRUE") #Filters list to only include corresponding authors
Corresponding_AuthorsAffilations <- unnest(Corresponding_only, authorships_affiliations, names_sep = "_") #Unnests affiliations so that all affiliations are listed
Corresponding_AuthorsAffilations_NoAPC <- Corresponding_AuthorsAffilations %>% select(!"apc") #Removes APC data for clean export

#Generate institutional corresponding authors list
Institutional_CorrespondingAuthors <- filter(Corresponding_AuthorsAffilations, authorships_affiliations_id == paste0("https://openalex.org/", QueryInstitution_OpenAlex_ID)) #Filters corresponding authors list for institutional list
Institutional_CorrespondingAuthors_NoAPC <- Institutional_CorrespondingAuthors %>% select(!"apc") #Removes APC data for clean export
Institutional_GoldHybrid_CorrespondingAuthors_NoAPC <- filter(Institutional_CorrespondingAuthors_NoAPC, oa_status == "gold" | oa_status == "hybrid") #Filters institutional corresponding authors list for gold and hybrid only

#Generate Paid/List APC list
Institutional_APCs <- unnest(Institutional_FilteredWorks, apc, names_sep = "_") #Unnests APC data for export
Institutional_APCS_NoAffiliation <- Institutional_APCs %>% select(!"authorships") #Removes author data for clean export
Institutional_GoldHybrid_APCs <- filter(Institutional_APCS_NoAffiliation, oa_status == "gold" | oa_status == "hybrid") #Selects only gold and hybrid articles

#Build data frame for guide Worksheet
WorksheetNames <- c("Query",
                    "Works", 
                    "OAWorks",
                    "GoldHybrid",
                    "AllAuthors",
                    "AllAffiliations",
                    "InstAuthors", 
                    "InstAuthorsGoldHybrid",
                    "Corresponding", 
                    "InstCorresponding", 
                    "InstCorrespondingGoldHybrid",
                    "APCs",
                    "GoldHybridAPCs")
WorksheetDescriptions <- c("Query Information (includes warnings for articles with more than 100 authors)",
                           "All institutional Works",
                           "All open access Works",
                           "All gold and hybrid Instituional Works (as defined by OpenAlex) https://docs.openalex.org/api-entities/works/work-object#oa_status",
                           "All authors for all institutional works",
                           "All affiliations for all authors for all institutional works",
                           "All affiliated authors for the institution for all works",
                           "All affiliated authors for the institutional for Gold and Hybrid works",
                           "All corresponding authors for all works Note: This data is a work in progress https://docs.openalex.org/api-entities/works/work-object/authorship-object#is_corresponding",
                           "All corresponding authors affiliated with the institution for all works",
                           "All corresponding authors affiliated with the institution for all Gold and Hybrid works",
                           "All list and 'paid' APC data for all institutional works Note: 'paid' APC data often uses list price: https://docs.openalex.org/api-entities/works/work-object#apc_paid",
                           "All list and 'paid' APC data for all gold and hybrid works Note: 'paid' APC data often uses list price: https://docs.openalex.org/api-entities/works/work-object#apc_paid")
GuideSheet <- data.frame(Worksheet = WorksheetNames, Description = WorksheetDescriptions)

#Generate Excel file with worksheets for each data frame
ListofWorksheets <- list("Guide" = GuideSheet,
  "Query" = Query,
  "Works" = Institutional_FilteredWorksOnly, 
  "OAWorks" = Institutional_OAWorksOnly,
  "GoldHybrid" = Institutional_GoldHybrid,
  "AllAuthors" = Institutional_AllAuthorsOnly,
  "AllAffiliations" = AllAuthorAffiliations_NoAPC,
  "InstAuthors" = Institutional_AuthorsOnly, 
  "InstAuthorsGoldHybrid" = Institutional_GoldHybridAuthorsOnly,
  "Corresponding" = Corresponding_AuthorsAffilations_NoAPC, 
  "InstCorresponding" = Institutional_CorrespondingAuthors_NoAPC, 
  "InstGoldHybridCorresponding" = Institutional_GoldHybrid_CorrespondingAuthors_NoAPC,
  "APCs" = Institutional_APCS_NoAffiliation,
  "GoldHybridAPCs" = Institutional_GoldHybrid_APCs
  )

# Export file -------------------------------------------------------------

OutputFile <- paste(CurrentDate, "OpenAlex_TA_Corresponding", QueryStartDate, QueryEndDate, ".xlsx", sep = "_") #Build output file name using query values
OutputPath <- paste("./data/", OutputFile) #File path for results

write_xlsx(ListofWorksheets, OutputPath)
