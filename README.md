This is R code designed to query the [OpenAlex](https://openalex.org/) [API](https://docs.openalex.org/)
and retrieve institutional research outputs. It then takes those outputs and converts it to an Excel file for ease of use.

This code is inspired by the code developed by Teresa Schultz for her FSCI 2025 Workshop available here: (https://github.com/schauch/OpenAlexRFSCI2025).

I am *NOT* an R coder. This code represents my attempt to try to solve two problems with the OpenAlex web interface:
1. That OpenAlex often doesn't play nice with Excel.
2. That it takes a fire hose approach that can require removing a lot of columns depending on your needs.

If you are an R coder, or just someone who knows more about R than I do, I welcome any feedback you have.

I also want to acknowledge and thank the wonderful people at [OpenAlex](https://openalex.org/team) and the [openalexR project](https://docs.ropensci.org/openalexR/).

## Installation

I wrote this code so that hopefully it can be used without any specific R knowledge or coding skills. The idea is that by replacing a few variables at the top, you'll be able to get a useful Excel file.

For the basics of getting started, this [guide](https://rstudio-education.github.io/hopr/starting.html) may be helpful.

1\. Install [R](https://cran.r-project.org/).

2\. Install [RStudio](https://posit.co/products/open-source/rstudio/?sid=1) to edit and run the project.

3\. [Download, clone, or fork](https://docs.github.com/en/get-started/start-your-journey/downloading-files-from-github) the repository files

### Dependencies:

The code uses three R packages (instructions for installing and loading them are in the script):

[openalexR](https://cran.r-project.org/web/packages/openalexR/index.html) - Information and recent updates available on [GitHub](https://github.com/ropensci/openalexR)

[tidyverse](https://tidyverse.org/)

[openxlsx2](https://cran.r-project.org/web/packages/openxlsx2/index.html)

## Usage

The main file is BaseOpenAlexQuery.R. There is a similar, experimental file called CorrespondingInstitutionOpenAlexQuery.R. However, I do *NOT* recommend using this file other than as a comparison, because OpenAlex [corresponding author data](https://docs.openalex.org/api-entities/works/work-object/authorship-object#is_corresponding) is not reliable.

**Important**
OpenAlex announced via their users list that an API key will be required starting February 13.

In order to access your API key, you need to register a free account with OpenAlex and go to the [Settings page](https://openalex.org/settings/api) to access your API key.

Following the recommendation of OpenAlexR, the API key is stored in the .Rprofile file.

This line needs to be added: options(openalexR.apikey = "YOUR API KEY")

The main query code is on lines 20-26:

#Query String: Modify the variables below to adjust query  
QueryEntity = "works" #Entity  
QueryInstitution_OpenAlex_ID = "I32625721" #OpenAlex Institution ID is University of Saskatchewan. Must use capital I for filters to work.  
QueryType = "article" #Type  
QuerySourceType = "journal" #Source Type  
QueryStartDate = "2024-01-01" #Publication Start Date  
QueryEndDate = "2024-12-31" #Publication End Date

This code will only work with the [works entity](https://docs.openalex.org/api-entities/works).

In order to design your query, the easiest way is to use the OpenAlex web interface, copy the API link, and paste the relevant components into the appropriate fields. For the [default example](https://openalex.org/works?page=1&filter=authorships.institutions.lineage:i32625721,type:types/article,primary_location.source.type:source-types/journal,publication_year:2024) above,
click on the three dots and select "Show API query". You can then replace the sample ID above with your institution's ID. **Please note, you must capitalize the i in the institutional identifier for the code to work.

You can also modify the type and source type (or remove them) if you don't want to use the defaults.

![Screenshot](/images/OpenAlexScreenshot.png)

Once that's done you're ready to run the code.

The final Excel file is named according to query values and saved in the data folder. The naming convention is CurrentDate_OpenAlex_TA_QueryStartDate_QueryEndDate.xlsx, with all dates in YYYYMMDD format.

The Excel file consists of 14 worksheets:

1.  Guide: A listing of all of the worksheets  
2.  Query:	Query Information (includes warnings for articles with more than 100 authors)  
3.  Works: All institutional Works
4.  OAWorks: All open access Works
5.  GoldHybrid:	All gold and hybrid institutional Works (as defined by OpenAlex) https://docs.openalex.org/api-entities/works/work-object#oa_status  
6.  AllAuthors:	All authors for all institutional works  
7.  AllAffiliations:	All affiliations for all authors for all institutional works  
8.  InstAuthors:	All affiliated authors for the institution for all works  
9.  InstAuthorsGoldHybrid:	All affiliated authors for the institutional for Gold and Hybrid works  
10. Corresponding:	All corresponding authors for all works Note: This data is a work in progress https://docs.openalex.org/api-entities/works/work-object/authorship-object#is_corresponding  
11. InstCorresponding:	All corresponding authors affiliated with the institution for all works  
12. InstCorrespondingGoldHybrid:	All corresponding authors affiliated with the institution for all Gold and Hybrid works  
13. APCs:	All list and 'paid' APC data for all institutional works Note: 'paid' APC data often uses list price: https://docs.openalex.org/api-entities/works/work-object#apc_paid  
14. GoldHybridAPCs:	All list and 'paid' APC data for all gold and hybrid works Note: 'paid' APC data often uses list price: https://docs.openalex.org/api-entities/works/work-object#apc_paid  

<h2>For more advanced users:</h2>

There is a filter string to remove columns in lines 27-43. Those columns can be commented out or removed to add additional columns to the Excel file.
Where there is only one variable for the column, they should be integrated into the Excel file without issue.
Where there are multiple variables for the column/field, openxlsx2 will attempt to format the json to work with Excel, but the results may be messy or the code may fail entirely.

The code in this repository is Copyright Jason Friedmen, licensed under the MIT License.
