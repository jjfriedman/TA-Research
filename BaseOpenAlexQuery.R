#This code is based on the code developed by Teresa Schultz for her FSCI 2025 Workshop available here: https://github.com/schauch/OpenAlexRFSCI2025


file.edit("~/.Rprofile") #Add email address to use polite pool / Only needs to be done once

#Query String

Institutional_Works <- oa_fetch(
  entity = "works", #query type
  authorships.institutions.lineage = "i32625721", #OpenAlex Institution ID is USask
  type = "article", #type of item
  primary_location.source.type = "journal", #limit to journal articles
  from_publication_date = "2024-06-01", #start range
  to_publication_date = "2024-07-31", #end range
  mailto = oa_email(), #To use polite API
#  per_page = 25,
  verbose = TRUE
)

#Filter string Add/remove column headers as desired.
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

Institutional_FilteredWorks <- Institutional_Works %>% select(-any_of(FilterColumns))

Institutional_Authors <- unnest(Institutional_FilteredWorks, authorships, names_sep = "_")
Institutional_AuthorsAffilations <- unnest(Institutional_Authors, authorships_affiliations, names_sep = "_")
Institutional_APCs <- unnest(Institutional_AuthorsAffilations, apc, names_sep = "_")

# Institutional_APCs <- Institutional_APCs %>% select(-any_of('authorships_affiliation_raw'))

write.csv(Institutional_APCs, "OutputFiles/Institutional_DataRawFlattened.csv")
                                                    
