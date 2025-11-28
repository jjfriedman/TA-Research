This is R code designed to query the [OpenAlex](https://openalex.org/) [API](https://docs.openalex.org/)
and retrieve institutional research outputs. It then takes those outputs and converts it to an Excel file for ease of use.

This code is inspired by the code developed by Teresa Schultz for her FSCI 2025 Workshop available here: (https://github.com/schauch/OpenAlexRFSCI2025)

<h2>Getting Started</h2>

For the basics of getting started, this [guide](https://rstudio-education.github.io/hopr/starting.html) may be helpful.

First off, you'll need to install [R](https://cran.r-project.org/).

You'll need software to edit and run code. I recommend (and use) the free and open source [RStudio](https://posit.co/products/open-source/rstudio/?sid=1).

The code should do this for you (lines 4-6), but you'll also need install 3 packages:
[openalexR](https://github.com/ropensci/openalexR)
[tidayverse](https://tidyverse.org/)
[openxlsx2](https://janmarvin.github.io/openxlsx2/)

**Important**
On line 13, you'll need to enter your email address to use the OpenAlex [polite pool](https://docs.openalex.org/how-to-use-the-api/rate-limits-and-authentication#the-polite-pool). 

options(openalexR.mailto = "youremailaddress@goeshere.com")

This code does not support API key access.

The main query code is on lines 16-21:

#Query String: Modify the variables below to adjust query
QueryEntity = "works" #Entity
QueryInstitution_OpenAlex_ID = "I32625721" #OpenAlex Institution ID is University of Saskatchewan. Must use capital I for filters to work.
QueryType = "article" #Type
QuerySourceType = "journal" #Source Type
QueryStartDate = "2024-01-01" #Publication Start Date
QueryEndDate = "2024-12-31" #Publication End Date

This code will only work with the [works entity](https://docs.openalex.org/api-entities/works).

In order to design your query, the easiest way is to use the OpenAlex web interface, copy the API link, and paste the relevant components into the appropriate fields. For the [default example](https://openalex.org/works?page=1&filter=authorships.institutions.lineage:i32625721,type:types/article,primary_location.source.type:source-types/journal,publication_year:2024) above, click on the
