#This code is inspired by the code developed by Teresa Schultz for her FSCI 2025 Workshop available here: https://github.com/schauch/OpenAlexRFSCI2025

#Install and include necessary packages. Install only needs to be run once.
#install.packages("openalexR") 
#install.packages("tidyverse")

library(openalexR)
library(tidyverse)

#Use Current date and time to name directory
CurrentWD = getwd() #Get current working directory
CurrentTime <- substr(Sys.time(),1,19) #Get current time to the second
CurrentTime <- str_replace(CurrentTime,' ','_') #Replace space with underscore
CurrentTime <- str_replace_all(CurrentTime,':','_') #Replace : with underscore
OutputPath <- paste("Output", CurrentTime,sep = '_') #File path for results
dir.create(OutputPath) #Create output directory
setwd(OutputPath) #Set new working directory for output

options(openalexR.mailto = "jason.friedman@usask.ca") #Add email address to use polite pool / Only needs to be done once

Institution_OpenAlex_ID = "I32625721" #OpenAlex Institution ID is USask
InstitutionFilterString = paste0("https://openalex.org/", Institution_OpenAlex_ID) #Adds OpenAlex URL string to institution ID for filtering

#Query String

Institutional_Works <- oa_fetch(
  entity = "works", #query type
  authorships.institutions.lineage = Institution_OpenAlex_ID, 
  type = "article", #type of item
  primary_location.source.type = "journal", #limit to journal articles
  from_publication_date = "2023-01-01", #start range
  to_publication_date = "2023-12-31", #end range
#  mailto = oa_email(), #To use polite API
#  per_page = 25,
  verbose = TRUE
)

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

#Generate initial works list
Institutional_FilteredWorks <- Institutional_Works %>% select(-any_of(FilterColumns)) #Removes columns indicated above
Institutional_FilteredWorksOnly <- Institutional_FilteredWorks %>% select(-any_of(c("authorships","apc"))) #Removes nested data authorships/apc for a clean works list for export
write.csv(Institutional_FilteredWorksOnly,"WorksList.csv") #Exports works list

#Generate initial authors list
Institutional_Authors <- unnest(Institutional_FilteredWorks, authorships, names_sep = "_") #Unnests authorships so that all authors are listed
Institutional_AuthorsOnly <- Institutional_Authors %>% select(-any_of(c("authorships_affiliations","apc"))) #Removes nested data affiliations and apc for a clearn authors list for export
write.csv(Institutional_AuthorsOnly,"AuthorsList.csv") #Exports authors list

#Generate corresponding authors list
Corresponding_only <- filter(Institutional_Authors, authorships_is_corresponding == "TRUE") #Filters list to only include corresponding authors
Corresponding_AuthorsAffilations <- unnest(Corresponding_only, authorships_affiliations, names_sep = "_") #Unnests affiliations so that all affiliations are listed
Corresponding_AuthorsAffilations_NoAPC <- Corresponding_AuthorsAffilations %>% select(!"apc") #Removes APC data for clean export
write.csv(Corresponding_AuthorsAffilations_NoAPC,"CorrespondingAuthorsList.csv") #Exports corresponding authors list

#Generate institutional corresponding authors list
Institutional_CorrespondingAuthors <- filter(Corresponding_AuthorsAffilations, authorships_affiliations_id == paste0("https://openalex.org/", Institution_OpenAlex_ID)) #Filters corresponding authors list 
Institutional_CorrespondingAuthors_NoAPC <- Institutional_CorrespondingAuthors %>% select(!"apc") #Removes APC data for clean export
write.csv(Institutional_CorrespondingAuthors_NoAPC,"InstitutionalCorrespondingAuthors.csv") #Exports institutional corresponding authors list

#Generate Paid/List APC list
#Institutional_OA <- filter(Institutional_AuthorsOnly, is_oa == "TRUE")
Institutional_APCs <- unnest(Institutional_CorrespondingAuthors, apc, names_sep = "_") #Unnests APC data for export
#Corresponding_only <- filter(Institutional_APCs, authorships_is_corresponding == "TRUE")

# Institutional_APCs <- Institutional_APCs %>% select(-any_of('authorships_affiliation_raw'))

write.csv(Institutional_APCs,"InstitutionalAPCList.csv") #Writes institutional corresponding author APC data
#write.csv(Corresponding_only, CorrespondingOutputPath)

setwd(CurrentWD) #Return to old working directory
                                                  
