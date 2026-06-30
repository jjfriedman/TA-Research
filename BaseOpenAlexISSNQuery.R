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
QueryEntity = "sources" #Entity
QueryISSN = c(
              "2292-1141" #Default is USURJ
) #List ISSNs here in quotes separated by commas

#This code uses the query variables above to call the OpenAlex API
ISSN_Sources <- oa_fetch(
  entity = QueryEntity, #query type
  issn = QueryISSN,
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
  'host_organization_lineage', 
  'summary_stats', 
  'ids', 
  'societies', 
  'alternate_titles', 
  'counts_by_year', 
  'topics'
)


#Generate initial sources list
ISSN_FilteredSources <- ISSN_Sources %>% select(-any_of(FilterColumns)) #Removes columns indicated above
ISSN_FilteredSourcesOnly <- ISSN_FilteredSources %>% select(-any_of(c("issn","apc_prices"))) #Removes nested data authorships/apc for a clean works list for export
ISSN_OASourcesOnly <- filter(ISSN_FilteredSourcesOnly,is_oa == TRUE) #Generate OA works list

#Generate All ISSN list
ISSN_AllISSNs <- unnest(ISSN_FilteredSources, issn, names_sep = "_") #Unnests ISSNs so that all ISSNs are listed
ISSN_AllISSNsOnly <- ISSN_AllISSNs %>% select(!"apc_prices") #Removes nested data affiliations and apc for a clean authors list for export

#Generate All ISSNs All APCs list
ISSN_AllAPCsAllISSNs <- unnest(ISSN_AllISSNs, apc_prices, names_sep = "_") #Unnests affiliations so that all affiliations are listed

#Structure data for Excel

#Build data frame for guide Worksheet
WorksheetNames <- c("Query",
                    "Sources", 
                    "OASources",
                    "AllISSNs",
                    "AllISSNsAndAPCs"
                    )
WorksheetDescriptions <- c("Query Information (includes warnings)",
                           "All sources",
                           "All open access sources",
                           "All ISSNs for all sources",
                           "All APC prices for all ISSNs for all sources"
                           )
GuideSheet <- data.frame(Worksheet = WorksheetNames, Description = WorksheetDescriptions)

#Generate Excel file with worksheets for each data frame
ListofWorksheets <- list("Guide" = GuideSheet,
                         "Query" = Query,
                         "Sources" = ISSN_FilteredSourcesOnly, 
                         "OASources" = ISSN_OASourcesOnly,
                         "AllISSNs" = ISSN_AllISSNsOnly,
                         "AllISSNsAndAPCs" = ISSN_AllAPCsAllISSNs
                          )


#Export

OutputFile <- paste(CurrentTime, "OpenAlex_TA", "ISSN", ".xlsx", sep = "_") #Build output file name using query values
OutputPath <- paste0("./data/", OutputFile) #File path for results

write_xlsx(ListofWorksheets, OutputPath)

paste("The code successfully completed. The Excel file is located here:",OutputPath )
