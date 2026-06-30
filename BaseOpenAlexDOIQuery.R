#This code is inspired by the code developed by Teresa Schultz for her FSCI 2025 Workshop available here: https://github.com/schauch/OpenAlexRFSCI2025

# Install necessary packages by removing the # symbol before the next three lines and running them. Packages only need to be installed once.
# install.packages("openalexR") #Package to query OpenAlex
# install.packages("tidyverse") #Package to work with data more easily
# install.packages("openxlsx2") #Package to export to Excel

#Load the necessary packages
library(openalexR)
library(tidyverse)
library(openxlsx2)

#Query Settings
#As of February 13, 2026, OpenAlex requires an API key (available for free).
#To find your API key, Go to openalex.org and create a free account
#Copy your API key from https://openalex.org/settings/api-key

file.edit("~/.Rprofile") #Opens the Rprofile file. You will need to copy and paste options(openalexR.apikey = "YOUR API KEY") into the file. Make sure to include the quotation marks around the API key. This only needs to be done once (or if you receive a new API key)

#Query String: Modify the variables below to adjust query
QueryEntity = "works" #Entity
QueryDOI = c() #List DOIs here in quotes separated by commas

#This code uses the query variables above to call the OpenAlex API
DOI_Works <- oa_fetch(
  entity = QueryEntity, #query type
  doi = QueryDOI,
  verbose = TRUE,
#  options = list("data-version" = 1), #Toggle to use version 1/old version of OpenAlex
)

#Process data

#Records query information, time stamp, and warning in data frame for future Export
CurrentTime <- substr(Sys.time(),1,19) #Get current time to the second
CurrentTime <- str_replace(CurrentTime,' ','_') #Replace space with underscore
CurrentTime <- str_replace_all(CurrentTime,':','_') #Replace : with underscore
if (exists("last.warning")) { #Checks if a warning exists
  QueryWarnings = names(last.warning) #Save warning message
  QueryWarnings <- QueryWarnings[!QueryWarnings %in% c("\033[38;5;232m\033[33m!\033[38;5;232m `oa_fetch()` and `oa2df()` now return new names for some columns in openalexR\n  v2.0.0.\n\033[36mℹ\033[38;5;232m See NEWS.md for the list of changes.\n\033[36mℹ\033[38;5;232m Call `get_coverage()` to view all updated columns and their original names in\n  OpenAlex.\033[39m\n\033[90mThis warning is displayed once every 8 hours.\033[39m")] #This code filters out a specific openalexr warning about column name changes.
  if (length(QueryWarnings) == 0) {QueryWarnings = "No warnings"} #Enters "No warnings" if there are no warnings
} else {
  QueryWarnings = "No warnings" #Enters "No warnings" if there are no warnings
}
QueryStructure <- c("Entity", "Timestamp", "Warnings") #Query labels
QueryValues <- c(QueryEntity, CurrentTime, QueryWarnings) #Query values
Query <- data.frame(Request = QueryStructure, Value = QueryValues) #Create data frame for Query information

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
  'awards',
  'funders',
  'sustainable_development_goals'
)


#Generate initial works list
DOI_FilteredWorks <- DOI_Works %>% select(-any_of(FilterColumns)) #Removes columns indicated above
DOI_FilteredWorksOnly <- DOI_FilteredWorks %>% select(-any_of(c("authorships","apc"))) #Removes nested data authorships/apc for a clean works list for export
DOI_OAWorksOnly <- filter(DOI_FilteredWorksOnly,is_oa_anywhere == TRUE) #Generate OA works list
DOI_GoldHybrid <- filter(DOI_FilteredWorksOnly,oa_status == "gold" | oa_status == "hybrid") #Generate Gold/Hybrid only works list (possible APC paid)

#Generate initial authors list
DOI_AllAuthors <- unnest(DOI_FilteredWorks, authorships, names_sep = "_") #Unnests authorships so that all authors are listed
DOI_AllAuthorsOnly <- Institutional_DOIAuthors %>% select(-any_of(c("authorships_affiliations","apc"))) #Removes nested data affiliations and apc for a clean authors list for export

#Generate affiliations list
AllAuthorAffiliations <- unnest(DOI_AllAuthors, authorships_affiliations, names_sep = "_") #Unnests affiliations so that all affiliations are listed
AllAuthorAffiliations_NoAPC <- AllAuthorAffiliations %>% select(!"apc") #Removes APC data for clean export

#Generate corresponding authors list
Corresponding_only <- filter(DOI_AllAuthors, authorships_is_corresponding == "TRUE") #Filters list to only include corresponding authors
Corresponding_AuthorsAffilations <- unnest(Corresponding_only, authorships_affiliations, names_sep = "_") #Unnests affiliations so that all affiliations are listed
Corresponding_AuthorsAffilations_NoAPC <- Corresponding_AuthorsAffilations %>% select(!"apc") #Removes APC data for clean export

#Generate Paid/List APC list
DOI_APCs <- unnest(DOI_FilteredWorks, apc, names_sep = "_") #Unnests APC data for export
DOI_APCS_NoAffiliation <- DOI_APCs %>% select(!"authorships") #Removes author data for clean export
DOI_GoldHybrid_APCs <- filter(DOI_APCS_NoAffiliation, oa_status == "gold" | oa_status == "hybrid") #Selects only gold and hybrid articles

#Structure data for Excel

#Build data frame for guide Worksheet
WorksheetNames <- c("Query",
                    "Works", 
                    "OAWorks",
                    "GoldHybrid",
                    "AllAuthors",
                    "AllAffiliations",
                    "Corresponding", 
                    "APCs",
                    "GoldHybridAPCs")
WorksheetDescriptions <- c("Query Information (includes warnings for articles with more than 100 authors)",
                           "All works",
                           "All open access Works",
                           "All affiliations for all authors for all works",
                           "All corresponding authors for all works Note: This data is a work in progress https://docs.openalex.org/api-entities/works/work-object/authorship-object#is_corresponding",
                           "All list and 'paid' APC data for all institutional works Note: 'paid' APC data often uses list price: https://docs.openalex.org/api-entities/works/work-object#apc_paid",
                           "All list and 'paid' APC data for all gold and hybrid works Note: 'paid' APC data often uses list price: https://docs.openalex.org/api-entities/works/work-object#apc_paid")
GuideSheet <- data.frame(Worksheet = WorksheetNames, Description = WorksheetDescriptions)

#Generate Excel file with worksheets for each data frame
ListofWorksheets <- list("Guide" = GuideSheet,
                         "Query" = Query,
                         "Works" = DOI_FilteredWorksOnly, 
                         "OAWorks" = DOI_OAWorksOnly,
                         "GoldHybrid" = DOI_GoldHybrid,
                         "AllAuthors" = DOI_AllAuthorsOnly,
                         "AllAffiliations" = AllAuthorAffiliations_NoAPC,
                         "Corresponding" = DOI_AuthorsAffilations_NoAPC, 
                         "APCs" = DOI_APCS_NoAffiliation,
                         "GoldHybridAPCs" = DOI_GoldHybrid_APCs
)


#Export

OutputFile <- paste(CurrentTime, "OpenAlex_TA", "DOI", ".xlsx", sep = "_") #Build output file name using query values
OutputPath <- paste0("./data/", OutputFile) #File path for results

write_xlsx(ListofWorksheets, OutputPath)

paste("The code successfully completed. The Excel file is located here:",OutputPath )
