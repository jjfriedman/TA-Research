# Developed by Jason Friedman with assistance from Microsoft M365 Copilot (GPT-5 reasoning model)

# ============================================================
# KBART serials -> OpenAlex source matches -> work COUNTS only

# ------------------------------------------------------------
# SECTION 0: Load packages
# ------------------------------------------------------------

# Install necessary packages by removing the # symbol before the next three lines and running them. Packages only need to be installed once.
# install.packages("openalexR") #Package to query OpenAlex
# install.packages("tidyverse") #Package to work with data more easily
# install.packages("openxlsx2") #Package to export to Excel

#Load the necessary packages
library(openalexR)
library(tidyverse)
library(openxlsx2)


# ------------------------------------------------------------
# SECTION 1: User settings
# ------------------------------------------------------------

# Set the path to your KBART tab-delimited file.
kbart_file <- "KBART file path"

# Set the path to the output Excel workbook.
output_xlsx <- "Excel file output"

# Optional manual upper limit for open-ended coverage windows.
# Leave as NA_character_ to keep open-ended rows truly open-ended.
open_end_to <- NA_character_

# Optional test mode.
# Set to NULL to run the entire file.
# Set to an integer like 10 or 25 to test a subset first.
test_n <- NULL

#Flag for if you want the OpenAlex queries to run in verbose mode or not
OpenAlexVerboseMode = TRUE

# ------------------------------------------------------------
# SECTION 2: Read KBART as character data
# ------------------------------------------------------------

# Read the KBART file as tab-delimited text.
# Read every column as character so dates and ISSNs are not mangled on import.
kbart_raw <- readr::read_tsv(
  file = kbart_file,
  col_types = cols(.default = col_character()),
  na = c("", "NA")
)

# Define the KBART columns that are required regardless of which identifier is used.
required_kbart_cols <- c(
  "publication_title",
  "date_first_issue_online",
  "date_last_issue_online"
)

# Find any required columns that are missing.
missing_kbart_cols <- setdiff(required_kbart_cols, names(kbart_raw))

# Stop early if required KBART columns are missing.
if (length(missing_kbart_cols) > 0) {
  stop(
    "Missing required KBART columns: ",
    paste(missing_kbart_cols, collapse = ", ")
  )
}

# If online_identifier is absent, create it as all NA so fallback logic still works.
if (!"online_identifier" %in% names(kbart_raw)) {
  kbart_raw$online_identifier <- NA_character_
}

# If print_identifier is absent, create it as all NA so fallback logic still works.
if (!"print_identifier" %in% names(kbart_raw)) {
  kbart_raw$print_identifier <- NA_character_
}

# Stop if neither identifier column contains any usable values.
if (
  all(is.na(kbart_raw$online_identifier) | str_squish(kbart_raw$online_identifier) == "") &&
  all(is.na(kbart_raw$print_identifier) | str_squish(kbart_raw$print_identifier) == "")
) {
  stop("Neither online_identifier nor print_identifier contains usable values.")
}


# ------------------------------------------------------------
# SECTION 3: Parse KBART dates and create clean working table
# ------------------------------------------------------------

# Parse the optional manual upper-limit date once.
# This accepts YYYY-MM-DD, YYYY-MM, or YYYY if open_end_to is used.
open_end_to_date <- suppressWarnings(
  as.Date(
    parse_date_time(
      open_end_to,
      orders = c("ymd", "Ymd", "Y-m", "Ym", "Y")
    )
  )
)

# Create the clean KBART working table with explicit provenance names.
kbart <- kbart_raw %>%
  mutate(
    # Create a stable row ID so every KBART row can be traced across sheets.
    kbart_row_id = row_number(),

    # Copy the publication title into an explicit provenance-labeled column.
    publication_title_kbart = publication_title,

    # Normalize the online identifier by trimming extra whitespace.
    online_identifier_kbart = str_squish(online_identifier),

    # Normalize the print identifier by trimming extra whitespace.
    print_identifier_kbart = str_squish(print_identifier),

    # Normalize the KBART lower-date text.
    date_first_issue_online_kbart = str_squish(date_first_issue_online),

    # Normalize the KBART upper-date text.
    date_last_issue_online_kbart = str_squish(date_last_issue_online),

    # Choose which identifier will be used for OpenAlex source lookup.
    # Prefer online_identifier when it exists.
    # Otherwise fall back to print_identifier.
    identifier_used_for_lookup_kbart = case_when(
      !is.na(online_identifier_kbart) & online_identifier_kbart != "" ~ online_identifier_kbart,
      (is.na(online_identifier_kbart) | online_identifier_kbart == "") &
        !is.na(print_identifier_kbart) & print_identifier_kbart != "" ~ print_identifier_kbart,
      TRUE ~ NA_character_
    ),

    # Record which identifier type was used for lookup for transparency.
    identifier_type_used_for_lookup_kbart = case_when(
      !is.na(online_identifier_kbart) & online_identifier_kbart != "" ~ "online_identifier",
      (is.na(online_identifier_kbart) | online_identifier_kbart == "") &
        !is.na(print_identifier_kbart) & print_identifier_kbart != "" ~ "print_identifier",
      TRUE ~ NA_character_
    ),

    # Build a safe lower-bound text value first.
    # YYYY-MM-DD stays as-is.
    # YYYY-MM becomes the first day of the month.
    # YYYY becomes the first day of the year.
    lower_bound_text = case_when(
      str_detect(date_first_issue_online_kbart, "^\\d{4}-\\d{2}-\\d{2}$") ~ date_first_issue_online_kbart,
      str_detect(date_first_issue_online_kbart, "^\\d{4}-\\d{2}$") ~ paste0(date_first_issue_online_kbart, "-01"),
      str_detect(date_first_issue_online_kbart, "^\\d{4}$") ~ paste0(date_first_issue_online_kbart, "-01-01"),
      TRUE ~ NA_character_
    ),

    # Convert the safe lower-bound text to Date.
    lower_bound_date = as.Date(lower_bound_text),

    # Build a helper text value used only for YYYY-MM upper dates.
    upper_year_month_start_text = case_when(
      str_detect(date_last_issue_online_kbart, "^\\d{4}-\\d{2}$") ~ paste0(date_last_issue_online_kbart, "-01"),
      TRUE ~ NA_character_
    ),

    # Convert that helper text to Date.
    upper_year_month_start_date = as.Date(upper_year_month_start_text),

    # Build a safe upper-bound text value first.
    # YYYY-MM-DD stays as-is.
    # YYYY-MM becomes the last day of the month.
    # YYYY becomes the last day of the year.
    upper_bound_text = case_when(
      str_detect(date_last_issue_online_kbart, "^\\d{4}-\\d{2}-\\d{2}$") ~ date_last_issue_online_kbart,
      str_detect(date_last_issue_online_kbart, "^\\d{4}-\\d{2}$") ~ as.character(
        ceiling_date(upper_year_month_start_date, unit = "month") - days(1)
      ),
      str_detect(date_last_issue_online_kbart, "^\\d{4}$") ~ paste0(date_last_issue_online_kbart, "-12-31"),
      TRUE ~ NA_character_
    ),

    # Convert the safe upper-bound text to Date.
    upper_bound_date_kbart = as.Date(upper_bound_text),

    # Create the effective upper date used in OpenAlex queries.
    # If KBART upper date exists, use it.
    # Else if KBART upper date is blank and open_end_to is set, use open_end_to.
    # Else keep NA, meaning open-ended.
    upper_bound_date_effective = case_when(
      !is.na(upper_bound_date_kbart) ~ upper_bound_date_kbart,
      is.na(upper_bound_date_kbart) & !is.na(open_end_to_date) ~ open_end_to_date,
      TRUE ~ as.Date(NA)
    )
  ) %>%
  select(
    # Keep only the explicit clean columns needed downstream.
    kbart_row_id,
    publication_title_kbart,
    online_identifier_kbart,
    print_identifier_kbart,
    identifier_used_for_lookup_kbart,
    identifier_type_used_for_lookup_kbart,
    date_first_issue_online_kbart,
    date_last_issue_online_kbart,
    lower_bound_date,
    upper_bound_date_effective
  )

# Apply test mode if requested.
if (!is.null(test_n)) {
  kbart <- kbart %>% slice_head(n = test_n)
}

# Print a quick structure preview of the KBART working table.
print(glimpse(kbart))


# ------------------------------------------------------------
# SECTION 4: Create empty log tables
# ------------------------------------------------------------

# Create the main query log.
# This captures all major steps, including successful steps, warnings, and errors.
query_log <- tibble(
  log_time = character(),
  step = character(),
  kbart_row_id = integer(),
  publication_title_kbart = character(),
  online_identifier_kbart = character(),
  print_identifier_kbart = character(),
  identifier_used_for_lookup_kbart = character(),
  identifier_type_used_for_lookup_kbart = character(),
  openalex_source_id = character(),
  status = character(),
  detail = character()
)

# Create a companion issue log.
# This is a filtered log that only stores warnings and errors.
query_issues <- tibble(
  log_time = character(),
  step = character(),
  kbart_row_id = integer(),
  publication_title_kbart = character(),
  online_identifier_kbart = character(),
  print_identifier_kbart = character(),
  identifier_used_for_lookup_kbart = character(),
  identifier_type_used_for_lookup_kbart = character(),
  openalex_source_id = character(),
  status = character(),
  detail = character()
)


# ------------------------------------------------------------
# SECTION 5: Resolve KBART rows to OpenAlex source matches
# ------------------------------------------------------------

# Create an empty table to collect all source matches.
source_matches <- tibble()

# Loop over each KBART row one at a time.
for (i in seq_len(nrow(kbart))) {

  # Pull out the current KBART row.
  row_i <- kbart[i, ]

  # Pull out basic values used repeatedly below.
  kbart_row_id_i <- row_i$kbart_row_id
  publication_title_i <- row_i$publication_title_kbart
  online_id_i <- row_i$online_identifier_kbart
  print_id_i <- row_i$print_identifier_kbart
  identifier_used_i <- row_i$identifier_used_for_lookup_kbart
  identifier_type_i <- row_i$identifier_type_used_for_lookup_kbart

  # If there is no usable identifier at all, log that fact and skip the row.
  if (is.na(identifier_used_i) || identifier_used_i == "") {

    query_log <- query_log %>%
      add_row(
        log_time = as.character(Sys.time()),
        step = "source_lookup",
        kbart_row_id = kbart_row_id_i,
        publication_title_kbart = publication_title_i,
        online_identifier_kbart = online_id_i,
        print_identifier_kbart = print_id_i,
        identifier_used_for_lookup_kbart = identifier_used_i,
        identifier_type_used_for_lookup_kbart = identifier_type_i,
        openalex_source_id = NA_character_,
        status = "skipped",
        detail = "Both online_identifier and print_identifier are blank; source lookup skipped."
      )

    query_issues <- query_issues %>%
      add_row(
        log_time = as.character(Sys.time()),
        step = "source_lookup",
        kbart_row_id = kbart_row_id_i,
        publication_title_kbart = publication_title_i,
        online_identifier_kbart = online_id_i,
        print_identifier_kbart = print_id_i,
        identifier_used_for_lookup_kbart = identifier_used_i,
        identifier_type_used_for_lookup_kbart = identifier_type_i,
        openalex_source_id = NA_character_,
        status = "warning",
        detail = "Both online_identifier and print_identifier are blank; source lookup skipped."
      )

    next
  }

  # Query the OpenAlex sources endpoint using the chosen ISSN.
  # This is usually online_identifier, but may be print_identifier if online is blank.
  source_res <- tryCatch(
    {
      oa_fetch(
        entity = "sources",
        issn = identifier_used_i,
        output = "tibble",
        verbose = OpenAlexVerboseMode
      )
    },
    error = function(e) e
  )

  # If the source lookup produced an error, log it and continue.
  if (inherits(source_res, "error")) {

    query_log <- query_log %>%
      add_row(
        log_time = as.character(Sys.time()),
        step = "source_lookup",
        kbart_row_id = kbart_row_id_i,
        publication_title_kbart = publication_title_i,
        online_identifier_kbart = online_id_i,
        print_identifier_kbart = print_id_i,
        identifier_used_for_lookup_kbart = identifier_used_i,
        identifier_type_used_for_lookup_kbart = identifier_type_i,
        openalex_source_id = NA_character_,
        status = "error",
        detail = conditionMessage(source_res)
      )

    query_issues <- query_issues %>%
      add_row(
        log_time = as.character(Sys.time()),
        step = "source_lookup",
        kbart_row_id = kbart_row_id_i,
        publication_title_kbart = publication_title_i,
        online_identifier_kbart = online_id_i,
        print_identifier_kbart = print_id_i,
        identifier_used_for_lookup_kbart = identifier_used_i,
        identifier_type_used_for_lookup_kbart = identifier_type_i,
        openalex_source_id = NA_character_,
        status = "error",
        detail = conditionMessage(source_res)
      )

    next
  }

  # If openalexR returned NULL for a no-match case, convert it to an empty tibble.
  # This avoids nrow(NULL) problems and lets the script continue cleanly.
  if (is.null(source_res)) {
    source_res <- tibble()
  }

  # If the returned object is not a data frame/tibble, log it as unexpected.
  # This is a safety check so odd return objects do not crash the script.
  if (!is.data.frame(source_res)) {

    query_log <- query_log %>%
      add_row(
        log_time = as.character(Sys.time()),
        step = "source_lookup",
        kbart_row_id = kbart_row_id_i,
        publication_title_kbart = publication_title_i,
        online_identifier_kbart = online_id_i,
        print_identifier_kbart = print_id_i,
        identifier_used_for_lookup_kbart = identifier_used_i,
        identifier_type_used_for_lookup_kbart = identifier_type_i,
        openalex_source_id = NA_character_,
        status = "warning",
        detail = paste("Source lookup returned a non-data-frame object for identifier:", identifier_used_i)
      )

    query_issues <- query_issues %>%
      add_row(
        log_time = as.character(Sys.time()),
        step = "source_lookup",
        kbart_row_id = kbart_row_id_i,
        publication_title_kbart = publication_title_i,
        online_identifier_kbart = online_id_i,
        print_identifier_kbart = print_id_i,
        identifier_used_for_lookup_kbart = identifier_used_i,
        identifier_type_used_for_lookup_kbart = identifier_type_i,
        openalex_source_id = NA_character_,
        status = "warning",
        detail = paste(
          "Source lookup returned a non-data-frame object for identifier:",
          identifier_used_i,
          "| class =",
          paste(class(source_res), collapse = "; ")
        )
      )

    next
  }

  # If zero source matches were found, log that explicitly and continue.
  if (nrow(source_res) == 0) {

    query_log <- query_log %>%
      add_row(
        log_time = as.character(Sys.time()),
        step = "source_lookup",
        kbart_row_id = kbart_row_id_i,
        publication_title_kbart = publication_title_i,
        online_identifier_kbart = online_id_i,
        print_identifier_kbart = print_id_i,
        identifier_used_for_lookup_kbart = identifier_used_i,
        identifier_type_used_for_lookup_kbart = identifier_type_i,
        openalex_source_id = NA_character_,
        status = "ok_zero_matches",
        detail = paste(
          "No OpenAlex source matches found for identifier used for lookup:",
          identifier_used_i,
          "(",
          identifier_type_i,
          ")"
        )
      )

    query_issues <- query_issues %>%
      add_row(
        log_time = as.character(Sys.time()),
        step = "source_lookup",
        kbart_row_id = kbart_row_id_i,
        publication_title_kbart = publication_title_i,
        online_identifier_kbart = online_id_i,
        print_identifier_kbart = print_id_i,
        identifier_used_for_lookup_kbart = identifier_used_i,
        identifier_type_used_for_lookup_kbart = identifier_type_i,
        openalex_source_id = NA_character_,
        status = "warning",
        detail = paste(
          "No OpenAlex source matches found for identifier used for lookup:",
          identifier_used_i,
          "(",
          identifier_type_i,
          ")"
        )
      )

    next
  }

  # Start cleaning the source result table.
  source_res_clean <- source_res

  # Create an explicit full OpenAlex source ID column if the API returned 'id'.
  if ("id" %in% names(source_res_clean)) {
    source_res_clean <- source_res_clean %>% mutate(openalex_source_id = id)
  } else {
    source_res_clean <- source_res_clean %>% mutate(openalex_source_id = NA_character_)
  }

  # Create a short source ID for works filters.
  source_res_clean <- source_res_clean %>%
    mutate(openalex_source_id_short = str_remove(openalex_source_id, "^https://openalex.org/"))

  # Create a labeled source display-name column if present.
  if ("display_name" %in% names(source_res_clean)) {
    source_res_clean <- source_res_clean %>% mutate(openalex_source_display_name = display_name)
  } else {
    source_res_clean <- source_res_clean %>% mutate(openalex_source_display_name = NA_character_)
  }

  # Create a labeled ISSN-L column if present.
  if ("issn_l" %in% names(source_res_clean)) {
    source_res_clean <- source_res_clean %>% mutate(openalex_source_issn_l = issn_l)
  } else {
    source_res_clean <- source_res_clean %>% mutate(openalex_source_issn_l = NA_character_)
  }

  # Flatten the source ISSN field into a readable character column.
  # OpenAlex may return ISSNs as a vector/list, so collapse them with semicolons.
  if ("issn" %in% names(source_res_clean)) {
    source_res_clean <- source_res_clean %>%
      mutate(
        openalex_source_issn = purrr::map_chr(
          issn,
          ~ {
            if (is.null(.x)) {
              NA_character_
            } else if (length(.x) == 1 && is.character(.x)) {
              .x
            } else {
              paste(.x, collapse = "; ")
            }
          }
        )
      )
  } else {
    source_res_clean <- source_res_clean %>% mutate(openalex_source_issn = NA_character_)
  }

  # Keep only the explicit columns wanted in the source_matches sheet.
  source_res_clean <- source_res_clean %>%
    transmute(
      kbart_row_id = kbart_row_id_i,
      publication_title_kbart = publication_title_i,
      online_identifier_kbart = online_id_i,
      print_identifier_kbart = print_id_i,
      identifier_used_for_lookup_kbart = identifier_used_i,
      identifier_type_used_for_lookup_kbart = identifier_type_i,
      date_first_issue_online_kbart = row_i$date_first_issue_online_kbart,
      date_last_issue_online_kbart = row_i$date_last_issue_online_kbart,
      lower_bound_date = row_i$lower_bound_date,
      upper_bound_date_effective = row_i$upper_bound_date_effective,
      openalex_source_id,
      openalex_source_id_short,
      openalex_source_display_name,
      openalex_source_issn,
      openalex_source_issn_l
    )

  # Append the matches for this KBART row.
  source_matches <- bind_rows(source_matches, source_res_clean)

  # Add a log entry showing how many source matches were found.
  query_log <- query_log %>%
    add_row(
      log_time = as.character(Sys.time()),
      step = "source_lookup",
      kbart_row_id = kbart_row_id_i,
      publication_title_kbart = publication_title_i,
      online_identifier_kbart = online_id_i,
      print_identifier_kbart = print_id_i,
      identifier_used_for_lookup_kbart = identifier_used_i,
      identifier_type_used_for_lookup_kbart = identifier_type_i,
      openalex_source_id = NA_character_,
      status = "ok",
      detail = paste("Matched", nrow(source_res_clean), "OpenAlex source record(s).")
    )
}

# Print a quick structure preview of the source match table.
print(glimpse(source_matches))


# ------------------------------------------------------------
# SECTION 6: One-row diagnostic BEFORE full counting
# ------------------------------------------------------------

# Create an empty diagnostic table.
preflight_diagnostic <- tibble(
  diagnostic_item = character(),
  diagnostic_value = character()
)

# Run diagnostics only if we have at least one source match.
if (nrow(source_matches) > 0) {

  # Take the first source match as a test case.
  diag_row <- source_matches %>% slice(1)

  # Pull out the test values.
  diag_kbart_row_id <- diag_row$kbart_row_id
  diag_source_id <- diag_row$openalex_source_id
  diag_source_id_short <- diag_row$openalex_source_id_short
  diag_lower <- diag_row$lower_bound_date
  diag_upper <- diag_row$upper_bound_date_effective

  # Record the diagnostic identifiers and query dates.
  preflight_diagnostic <- preflight_diagnostic %>%
    add_row(diagnostic_item = "kbart_row_id", diagnostic_value = as.character(diag_kbart_row_id)) %>%
    add_row(diagnostic_item = "publication_title_kbart", diagnostic_value = as.character(diag_row$publication_title_kbart)) %>%
    add_row(diagnostic_item = "online_identifier_kbart", diagnostic_value = as.character(diag_row$online_identifier_kbart)) %>%
    add_row(diagnostic_item = "print_identifier_kbart", diagnostic_value = as.character(diag_row$print_identifier_kbart)) %>%
    add_row(diagnostic_item = "identifier_used_for_lookup_kbart", diagnostic_value = as.character(diag_row$identifier_used_for_lookup_kbart)) %>%
    add_row(diagnostic_item = "identifier_type_used_for_lookup_kbart", diagnostic_value = as.character(diag_row$identifier_type_used_for_lookup_kbart)) %>%
    add_row(diagnostic_item = "openalex_source_id", diagnostic_value = as.character(diag_source_id)) %>%
    add_row(diagnostic_item = "openalex_source_id_short", diagnostic_value = as.character(diag_source_id_short)) %>%
    add_row(diagnostic_item = "lower_bound_used", diagnostic_value = as.character(diag_lower)) %>%
    add_row(diagnostic_item = "upper_bound_used", diagnostic_value = as.character(diag_upper))

  # Build arguments for the total-count diagnostic query.
  diag_total_args <- list(
    entity = "works",
    output = "list",
    verbose = OpenAlexVerboseMode
  )
  diag_total_args[["primary_location.source.id"]] <- diag_source_id_short
  if (!is.na(diag_lower)) {
    diag_total_args[["from_publication_date"]] <- as.character(diag_lower)
  }
  if (!is.na(diag_upper)) {
    diag_total_args[["to_publication_date"]] <- as.character(diag_upper)
  }
  diag_total_args$count_only <- TRUE

  # Run the total-count diagnostic query.
  diag_total_res <- tryCatch(
    {
      do.call(oa_fetch, diag_total_args)
    },
    error = function(e) e
  )

  # Record the result structure so parser behavior can be inspected later.
  if (inherits(diag_total_res, "error")) {
    preflight_diagnostic <- preflight_diagnostic %>%
      add_row(diagnostic_item = "total_works_count_test", diagnostic_value = paste("ERROR:", conditionMessage(diag_total_res)))
  } else {
    preflight_diagnostic <- preflight_diagnostic %>%
      add_row(diagnostic_item = "total_works_count_test_class", diagnostic_value = paste(class(diag_total_res), collapse = "; ")) %>%
      add_row(diagnostic_item = "total_works_count_test_names", diagnostic_value = paste(names(diag_total_res), collapse = "; ")) %>%
      add_row(diagnostic_item = "total_works_count_test_str", diagnostic_value = paste(capture.output(str(diag_total_res, max.level = 2)), collapse = " | "))
  }

  # Build arguments for the OA-only count diagnostic query.
  diag_oa_args <- list(
    entity = "works",
    output = "list",
    verbose = OpenAlexVerboseMode
  )
  diag_oa_args[["primary_location.source.id"]] <- diag_source_id_short
  if (!is.na(diag_lower)) {
    diag_oa_args[["from_publication_date"]] <- as.character(diag_lower)
  }
  if (!is.na(diag_upper)) {
    diag_oa_args[["to_publication_date"]] <- as.character(diag_upper)
  }
  diag_oa_args[["open_access.is_oa"]] <- TRUE
  diag_oa_args$count_only <- TRUE

  # Run the OA-only count diagnostic query.
  diag_oa_res <- tryCatch(
    {
      do.call(oa_fetch, diag_oa_args)
    },
    error = function(e) e
  )

  # Record the OA-only result structure.
  if (inherits(diag_oa_res, "error")) {
    preflight_diagnostic <- preflight_diagnostic %>%
      add_row(diagnostic_item = "oa_works_count_test", diagnostic_value = paste("ERROR:", conditionMessage(diag_oa_res)))
  } else {
    preflight_diagnostic <- preflight_diagnostic %>%
      add_row(diagnostic_item = "oa_works_count_test_class", diagnostic_value = paste(class(diag_oa_res), collapse = "; ")) %>%
      add_row(diagnostic_item = "oa_works_count_test_names", diagnostic_value = paste(names(diag_oa_res), collapse = "; ")) %>%
      add_row(diagnostic_item = "oa_works_count_test_str", diagnostic_value = paste(capture.output(str(diag_oa_res, max.level = 2)), collapse = " | "))
  }

  # Build arguments for the OA-status grouped diagnostic query.
  diag_status_args <- list(
    entity = "works",
    output = "list",
    verbose = OpenAlexVerboseMode
  )
  diag_status_args[["primary_location.source.id"]] <- diag_source_id_short
  if (!is.na(diag_lower)) {
    diag_status_args[["from_publication_date"]] <- as.character(diag_lower)
  }
  if (!is.na(diag_upper)) {
    diag_status_args[["to_publication_date"]] <- as.character(diag_upper)
  }
  diag_status_args$group_by <- "oa_status"

  # Run the OA-status grouped diagnostic query.
  diag_status_res <- tryCatch(
    {
      do.call(oa_fetch, diag_status_args)
    },
    error = function(e) e
  )

  # Record the grouped result structure.
  if (inherits(diag_status_res, "error")) {
    preflight_diagnostic <- preflight_diagnostic %>%
      add_row(diagnostic_item = "oa_status_group_test", diagnostic_value = paste("ERROR:", conditionMessage(diag_status_res)))
  } else {
    preflight_diagnostic <- preflight_diagnostic %>%
      add_row(diagnostic_item = "oa_status_group_test_class", diagnostic_value = paste(class(diag_status_res), collapse = "; ")) %>%
      add_row(diagnostic_item = "oa_status_group_test_names", diagnostic_value = paste(names(diag_status_res), collapse = "; ")) %>%
      add_row(diagnostic_item = "oa_status_group_test_str", diagnostic_value = paste(capture.output(str(diag_status_res, max.level = 3)), collapse = " | "))
  }

  # Log completion of the preflight diagnostic.
  query_log <- query_log %>%
    add_row(
      log_time = as.character(Sys.time()),
      step = "preflight_diagnostic",
      kbart_row_id = diag_kbart_row_id,
      publication_title_kbart = as.character(diag_row$publication_title_kbart),
      online_identifier_kbart = as.character(diag_row$online_identifier_kbart),
      print_identifier_kbart = as.character(diag_row$print_identifier_kbart),
      identifier_used_for_lookup_kbart = as.character(diag_row$identifier_used_for_lookup_kbart),
      identifier_type_used_for_lookup_kbart = as.character(diag_row$identifier_type_used_for_lookup_kbart),
      openalex_source_id = diag_source_id,
      status = "ok",
      detail = "Preflight diagnostic completed; see preflight_diagnostic sheet."
    )
}


# ------------------------------------------------------------
# SECTION 7: Count works for each KBART row / matched source
# ------------------------------------------------------------

# Create an empty final count table.
counts_by_kbart_row <- tibble()

# Loop over each KBART row / matched source pair.
for (i in seq_len(nrow(source_matches))) {

  # Pull out the current source-match row.
  row_i <- source_matches[i, ]

  # Pull out identifiers and dates used repeatedly below.
  kbart_row_id_i <- row_i$kbart_row_id
  publication_title_i <- row_i$publication_title_kbart
  online_id_i <- row_i$online_identifier_kbart
  print_id_i <- row_i$print_identifier_kbart
  identifier_used_i <- row_i$identifier_used_for_lookup_kbart
  identifier_type_i <- row_i$identifier_type_used_for_lookup_kbart
  source_id_i <- row_i$openalex_source_id
  source_id_short_i <- row_i$openalex_source_id_short
  lower_i <- row_i$lower_bound_date
  upper_i <- row_i$upper_bound_date_effective

  # Start each count as NA so failures remain visible.
  total_works_count <- NA_integer_
  oa_works_count <- NA_integer_
  oa_status_diamond_count <- NA_integer_
  oa_status_gold_count <- NA_integer_
  oa_status_green_count <- NA_integer_
  oa_status_hybrid_count <- NA_integer_
  oa_status_bronze_count <- NA_integer_
  oa_status_closed_count <- NA_integer_

  # If the lower bound is missing, log a warning.
  if (is.na(lower_i)) {
    query_issues <- query_issues %>%
      add_row(
        log_time = as.character(Sys.time()),
        step = "count_setup",
        kbart_row_id = kbart_row_id_i,
        publication_title_kbart = publication_title_i,
        online_identifier_kbart = online_id_i,
        print_identifier_kbart = print_id_i,
        identifier_used_for_lookup_kbart = identifier_used_i,
        identifier_type_used_for_lookup_kbart = identifier_type_i,
        openalex_source_id = source_id_i,
        status = "warning",
        detail = "date_first_issue_online is missing or unparsable; counting without a lower bound."
      )
  }

  # Add a setup log entry.
  query_log <- query_log %>%
    add_row(
      log_time = as.character(Sys.time()),
      step = "count_setup",
      kbart_row_id = kbart_row_id_i,
      publication_title_kbart = publication_title_i,
      online_identifier_kbart = online_id_i,
      print_identifier_kbart = print_id_i,
      identifier_used_for_lookup_kbart = identifier_used_i,
      identifier_type_used_for_lookup_kbart = identifier_type_i,
      openalex_source_id = source_id_i,
      status = "ok",
      detail = paste("Using matched source ID:", source_id_short_i)
    )

  # --------------------------------------------------------
  # 7A: Query total works count
  # --------------------------------------------------------

  # Build the argument list for the total-count query.
  total_args <- list(
    entity = "works",
    output = "list",
    verbose = OpenAlexVerboseMode
  )
  total_args[["primary_location.source.id"]] <- source_id_short_i
  if (!is.na(lower_i)) {
    total_args[["from_publication_date"]] <- as.character(lower_i)
  }
  if (!is.na(upper_i)) {
    total_args[["to_publication_date"]] <- as.character(upper_i)
  }
  total_args$count_only <- TRUE

  # Run the total-count query.
  total_res <- tryCatch(
    {
      do.call(oa_fetch, total_args)
    },
    error = function(e) e
  )

  # If the total-count query errored, log it.
  if (inherits(total_res, "error")) {

    query_log <- query_log %>%
      add_row(
        log_time = as.character(Sys.time()),
        step = "count_total",
        kbart_row_id = kbart_row_id_i,
        publication_title_kbart = publication_title_i,
        online_identifier_kbart = online_id_i,
        print_identifier_kbart = print_id_i,
        identifier_used_for_lookup_kbart = identifier_used_i,
        identifier_type_used_for_lookup_kbart = identifier_type_i,
        openalex_source_id = source_id_i,
        status = "error",
        detail = conditionMessage(total_res)
      )

    query_issues <- query_issues %>%
      add_row(
        log_time = as.character(Sys.time()),
        step = "count_total",
        kbart_row_id = kbart_row_id_i,
        publication_title_kbart = publication_title_i,
        online_identifier_kbart = online_id_i,
        print_identifier_kbart = print_id_i,
        identifier_used_for_lookup_kbart = identifier_used_i,
        identifier_type_used_for_lookup_kbart = identifier_type_i,
        openalex_source_id = source_id_i,
        status = "error",
        detail = conditionMessage(total_res)
      )

  } else {

    # If the response is a list and has a top-level count field, use it.
    if (is.list(total_res) && "count" %in% names(total_res)) {
      total_works_count <- as.integer(total_res$count)
    }

    # If the response is a scalar numeric, use that as a fallback parser path.
    if (is.na(total_works_count) && is.numeric(total_res) && length(total_res) == 1) {
      total_works_count <- as.integer(total_res)
    }

    # Log an unexpected structure if the parser still failed.
    if (is.na(total_works_count)) {

      query_log <- query_log %>%
        add_row(
          log_time = as.character(Sys.time()),
          step = "count_total",
          kbart_row_id = kbart_row_id_i,
          publication_title_kbart = publication_title_i,
          online_identifier_kbart = online_id_i,
          print_identifier_kbart = print_id_i,
          identifier_used_for_lookup_kbart = identifier_used_i,
          identifier_type_used_for_lookup_kbart = identifier_type_i,
          openalex_source_id = source_id_i,
          status = "warning",
          detail = "Total count response structure was not recognized; total_works_count left as NA."
        )

      query_issues <- query_issues %>%
        add_row(
          log_time = as.character(Sys.time()),
          step = "count_total",
          kbart_row_id = kbart_row_id_i,
          publication_title_kbart = publication_title_i,
          online_identifier_kbart = online_id_i,
          print_identifier_kbart = print_id_i,
          identifier_used_for_lookup_kbart = identifier_used_i,
          identifier_type_used_for_lookup_kbart = identifier_type_i,
          openalex_source_id = source_id_i,
          status = "warning",
          detail = paste(
            "Unrecognized total-count response structure:",
            paste(capture.output(str(total_res, max.level = 2)), collapse = " | ")
          )
        )

    } else {

      # Log the successful total-count extraction.
      query_log <- query_log %>%
        add_row(
          log_time = as.character(Sys.time()),
          step = "count_total",
          kbart_row_id = kbart_row_id_i,
          publication_title_kbart = publication_title_i,
          online_identifier_kbart = online_id_i,
          print_identifier_kbart = print_id_i,
          identifier_used_for_lookup_kbart = identifier_used_i,
          identifier_type_used_for_lookup_kbart = identifier_type_i,
          openalex_source_id = source_id_i,
          status = "ok",
          detail = paste("total_works_count =", total_works_count)
        )
    }
  }

  # --------------------------------------------------------
  # 7B: Query OA-only works count
  # --------------------------------------------------------

  # Build the argument list for the OA-only count query.
  oa_args <- list(
    entity = "works",
    output = "list",
    verbose = OpenAlexVerboseMode
  )
  oa_args[["primary_location.source.id"]] <- source_id_short_i
  if (!is.na(lower_i)) {
    oa_args[["from_publication_date"]] <- as.character(lower_i)
  }
  if (!is.na(upper_i)) {
    oa_args[["to_publication_date"]] <- as.character(upper_i)
  }
  oa_args[["open_access.is_oa"]] <- TRUE
  oa_args$count_only <- TRUE

  # Run the OA-only count query.
  oa_res <- tryCatch(
    {
      do.call(oa_fetch, oa_args)
    },
    error = function(e) e
  )

  # If the OA-only query errored, log it.
  if (inherits(oa_res, "error")) {

    query_log <- query_log %>%
      add_row(
        log_time = as.character(Sys.time()),
        step = "count_oa",
        kbart_row_id = kbart_row_id_i,
        publication_title_kbart = publication_title_i,
        online_identifier_kbart = online_id_i,
        print_identifier_kbart = print_id_i,
        identifier_used_for_lookup_kbart = identifier_used_i,
        identifier_type_used_for_lookup_kbart = identifier_type_i,
        openalex_source_id = source_id_i,
        status = "error",
        detail = conditionMessage(oa_res)
      )

    query_issues <- query_issues %>%
      add_row(
        log_time = as.character(Sys.time()),
        step = "count_oa",
        kbart_row_id = kbart_row_id_i,
        publication_title_kbart = publication_title_i,
        online_identifier_kbart = online_id_i,
        print_identifier_kbart = print_id_i,
        identifier_used_for_lookup_kbart = identifier_used_i,
        identifier_type_used_for_lookup_kbart = identifier_type_i,
        openalex_source_id = source_id_i,
        status = "error",
        detail = conditionMessage(oa_res)
      )

  } else {

    # If the response is a list with a top-level count field, use it.
    if (is.list(oa_res) && "count" %in% names(oa_res)) {
      oa_works_count <- as.integer(oa_res$count)
    }

    # If the response is a scalar numeric, use it as a fallback parser path.
    if (is.na(oa_works_count) && is.numeric(oa_res) && length(oa_res) == 1) {
      oa_works_count <- as.integer(oa_res)
    }

    # Log an unexpected structure if the parser still failed.
    if (is.na(oa_works_count)) {

      query_log <- query_log %>%
        add_row(
          log_time = as.character(Sys.time()),
          step = "count_oa",
          kbart_row_id = kbart_row_id_i,
          publication_title_kbart = publication_title_i,
          online_identifier_kbart = online_id_i,
          print_identifier_kbart = print_id_i,
          identifier_used_for_lookup_kbart = identifier_used_i,
          identifier_type_used_for_lookup_kbart = identifier_type_i,
          openalex_source_id = source_id_i,
          status = "warning",
          detail = "OA count response structure was not recognized; oa_works_count left as NA."
        )

      query_issues <- query_issues %>%
        add_row(
          log_time = as.character(Sys.time()),
          step = "count_oa",
          kbart_row_id = kbart_row_id_i,
          publication_title_kbart = publication_title_i,
          online_identifier_kbart = online_id_i,
          print_identifier_kbart = print_id_i,
          identifier_used_for_lookup_kbart = identifier_used_i,
          identifier_type_used_for_lookup_kbart = identifier_type_i,
          openalex_source_id = source_id_i,
          status = "warning",
          detail = paste(
            "Unrecognized OA-count response structure:",
            paste(capture.output(str(oa_res, max.level = 2)), collapse = " | ")
          )
        )

    } else {

      # Log the successful OA-count extraction.
      query_log <- query_log %>%
        add_row(
          log_time = as.character(Sys.time()),
          step = "count_oa",
          kbart_row_id = kbart_row_id_i,
          publication_title_kbart = publication_title_i,
          online_identifier_kbart = online_id_i,
          print_identifier_kbart = print_id_i,
          identifier_used_for_lookup_kbart = identifier_used_i,
          identifier_type_used_for_lookup_kbart = identifier_type_i,
          openalex_source_id = source_id_i,
          status = "ok",
          detail = paste("oa_works_count =", oa_works_count)
        )
    }
  }

  # --------------------------------------------------------
  # 7C: Query grouped counts by OA status
  # --------------------------------------------------------

  # Build the argument list for the grouped OA-status query.
  status_args <- list(
    entity = "works",
    output = "list",
    verbose = OpenAlexVerboseMode
  )
  status_args[["primary_location.source.id"]] <- source_id_short_i
  if (!is.na(lower_i)) {
    status_args[["from_publication_date"]] <- as.character(lower_i)
  }
  if (!is.na(upper_i)) {
    status_args[["to_publication_date"]] <- as.character(upper_i)
  }
  status_args$group_by <- "oa_status"

  # Run the grouped OA-status query.
  status_res <- tryCatch(
    {
      do.call(oa_fetch, status_args)
    },
    error = function(e) e
  )

  # Create an empty grouped-count table.
  status_tbl <- tibble()

  # If the first grouped query errored, try the OpenAlex fallback group field name.
  if (inherits(status_res, "error")) {

    status_args_fallback <- status_args
    status_args_fallback$group_by <- "open_access.oa_status"
    status_res_fallback <- tryCatch(
      {
        do.call(oa_fetch, status_args_fallback)
      },
      error = function(e) e
    )

    # If fallback also errored, log the failure.
    if (inherits(status_res_fallback, "error")) {

      query_log <- query_log %>%
        add_row(
          log_time = as.character(Sys.time()),
          step = "count_oa_status",
          kbart_row_id = kbart_row_id_i,
          publication_title_kbart = publication_title_i,
          online_identifier_kbart = online_id_i,
          print_identifier_kbart = print_id_i,
          identifier_used_for_lookup_kbart = identifier_used_i,
          identifier_type_used_for_lookup_kbart = identifier_type_i,
          openalex_source_id = source_id_i,
          status = "error",
          detail = paste(
            "Both group_by attempts failed:",
            conditionMessage(status_res),
            "||",
            conditionMessage(status_res_fallback)
          )
        )

      query_issues <- query_issues %>%
        add_row(
          log_time = as.character(Sys.time()),
          step = "count_oa_status",
          kbart_row_id = kbart_row_id_i,
          publication_title_kbart = publication_title_i,
          online_identifier_kbart = online_id_i,
          print_identifier_kbart = print_id_i,
          identifier_used_for_lookup_kbart = identifier_used_i,
          identifier_type_used_for_lookup_kbart = identifier_type_i,
          openalex_source_id = source_id_i,
          status = "error",
          detail = paste(
            "Both group_by attempts failed:",
            conditionMessage(status_res),
            "||",
            conditionMessage(status_res_fallback)
          )
        )

    } else {

      # If fallback returned an unnamed list of group records, convert it row-by-row.
      if (
        is.list(status_res_fallback) &&
        length(status_res_fallback) > 0 &&
        is.list(status_res_fallback[[1]]) &&
        "key" %in% names(status_res_fallback[[1]]) &&
        "count" %in% names(status_res_fallback[[1]])
      ) {
        status_tbl <- purrr::map_dfr(
          status_res_fallback,
          ~ tibble(
            key = as.character(.x$key),
            count = as.integer(.x$count)
          )
        )
      }

      # If fallback returned a data frame, use it as-is.
      if (nrow(status_tbl) == 0 && is.data.frame(status_res_fallback)) {
        status_tbl <- as_tibble(status_res_fallback)
      }

      # Log that the fallback field name was used.
      query_log <- query_log %>%
        add_row(
          log_time = as.character(Sys.time()),
          step = "count_oa_status",
          kbart_row_id = kbart_row_id_i,
          publication_title_kbart = publication_title_i,
          online_identifier_kbart = online_id_i,
          print_identifier_kbart = print_id_i,
          identifier_used_for_lookup_kbart = identifier_used_i,
          identifier_type_used_for_lookup_kbart = identifier_type_i,
          openalex_source_id = source_id_i,
          status = "ok_fallback_attempt",
          detail = "Used fallback group_by = open_access.oa_status"
        )
    }

  } else {

    # If the first grouped response is an unnamed list of group records, convert it row-by-row.
    if (
      is.list(status_res) &&
      length(status_res) > 0 &&
      is.list(status_res[[1]]) &&
      "key" %in% names(status_res[[1]]) &&
      "count" %in% names(status_res[[1]])
    ) {
      status_tbl <- purrr::map_dfr(
        status_res,
        ~ tibble(
          key = as.character(.x$key),
          count = as.integer(.x$count)
        )
      )
    }

    # If the first grouped response is already a data frame, use it directly.
    if (nrow(status_tbl) == 0 && is.data.frame(status_res)) {
      status_tbl <- as_tibble(status_res)
    }
  }

  # If we have a usable grouped table, normalize its keys and fill explicit status columns.
  if (nrow(status_tbl) > 0) {

    status_tbl <- status_tbl %>%
      mutate(
        key = tolower(as.character(key)),
        count = as.integer(count)
      )

    # Fill each requested OA status explicitly.
    if ("diamond" %in% status_tbl$key) {
      oa_status_diamond_count <- status_tbl$count[match("diamond", status_tbl$key)]
    } else {
      oa_status_diamond_count <- 0L
    }

    if ("gold" %in% status_tbl$key) {
      oa_status_gold_count <- status_tbl$count[match("gold", status_tbl$key)]
    } else {
      oa_status_gold_count <- 0L
    }

    if ("green" %in% status_tbl$key) {
      oa_status_green_count <- status_tbl$count[match("green", status_tbl$key)]
    } else {
      oa_status_green_count <- 0L
    }

    if ("hybrid" %in% status_tbl$key) {
      oa_status_hybrid_count <- status_tbl$count[match("hybrid", status_tbl$key)]
    } else {
      oa_status_hybrid_count <- 0L
    }

    if ("bronze" %in% status_tbl$key) {
      oa_status_bronze_count <- status_tbl$count[match("bronze", status_tbl$key)]
    } else {
      oa_status_bronze_count <- 0L
    }

    if ("closed" %in% status_tbl$key) {
      oa_status_closed_count <- status_tbl$count[match("closed", status_tbl$key)]
    } else {
      oa_status_closed_count <- 0L
    }

    # Log successful grouped retrieval.
    query_log <- query_log %>%
      add_row(
        log_time = as.character(Sys.time()),
        step = "count_oa_status",
        kbart_row_id = kbart_row_id_i,
        publication_title_kbart = publication_title_i,
        online_identifier_kbart = online_id_i,
        print_identifier_kbart = print_id_i,
        identifier_used_for_lookup_kbart = identifier_used_i,
        identifier_type_used_for_lookup_kbart = identifier_type_i,
        openalex_source_id = source_id_i,
        status = "ok",
        detail = "Grouped OA-status counts retrieved."
      )

  } else {

    # Log that the grouped query did not yield a usable table.
    query_log <- query_log %>%
      add_row(
        log_time = as.character(Sys.time()),
        step = "count_oa_status",
        kbart_row_id = kbart_row_id_i,
        publication_title_kbart = publication_title_i,
        online_identifier_kbart = online_id_i,
        print_identifier_kbart = print_id_i,
        identifier_used_for_lookup_kbart = identifier_used_i,
        identifier_type_used_for_lookup_kbart = identifier_type_i,
        openalex_source_id = source_id_i,
        status = "warning",
        detail = "Grouped OA-status query did not yield a usable table; status counts left as NA."
      )

    query_issues <- query_issues %>%
      add_row(
        log_time = as.character(Sys.time()),
        step = "count_oa_status",
        kbart_row_id = kbart_row_id_i,
        publication_title_kbart = publication_title_i,
        online_identifier_kbart = online_id_i,
        print_identifier_kbart = print_id_i,
        identifier_used_for_lookup_kbart = identifier_used_i,
        identifier_type_used_for_lookup_kbart = identifier_type_i,
        openalex_source_id = source_id_i,
        status = "warning",
        detail = "Grouped OA-status query did not yield a usable table."
      )
  }

  # Build the final output row for this KBART row / matched source pair.
  out_row <- tibble(
    kbart_row_id = kbart_row_id_i,
    publication_title_kbart = publication_title_i,
    online_identifier_kbart = online_id_i,
    print_identifier_kbart = print_id_i,
    identifier_used_for_lookup_kbart = identifier_used_i,
    identifier_type_used_for_lookup_kbart = identifier_type_i,
    date_first_issue_online_kbart = row_i$date_first_issue_online_kbart,
    date_last_issue_online_kbart = row_i$date_last_issue_online_kbart,
    query_lower_bound_used = as.character(lower_i),
    query_upper_bound_used = as.character(upper_i),
    openalex_source_id = source_id_i,
    openalex_source_id_short = source_id_short_i,
    openalex_source_display_name = row_i$openalex_source_display_name,
    openalex_source_issn = row_i$openalex_source_issn,
    openalex_source_issn_l = row_i$openalex_source_issn_l,
    total_works_count = total_works_count,
    oa_works_count = oa_works_count,
    oa_status_diamond_count = oa_status_diamond_count,
    oa_status_gold_count = oa_status_gold_count,
    oa_status_green_count = oa_status_green_count,
    oa_status_hybrid_count = oa_status_hybrid_count,
    oa_status_bronze_count = oa_status_bronze_count,
    oa_status_closed_count = oa_status_closed_count
  )

  # Append the row to the final count table.
  counts_by_kbart_row <- bind_rows(counts_by_kbart_row, out_row)
}

# Print a quick structure preview of the final count table.
print(glimpse(counts_by_kbart_row))


# ------------------------------------------------------------
# SECTION 8: Create a simple summary sheet
# ------------------------------------------------------------

# Create a compact summary table for the workbook.
summary_tbl <- tibble(
  metric = c(
    "kbart_rows_total",
    "kbart_rows_with_nonblank_online_identifier",
    "kbart_rows_with_nonblank_print_identifier",
    "kbart_rows_with_nonblank_lookup_identifier",
    "source_match_rows_total",
    "kbart_rows_with_at_least_one_source_match",
    "kbart_rows_without_source_match",
    "count_rows_total",
    "query_log_rows_total",
    "query_issue_rows_total",
    "test_mode_used",
    "test_n_rows_setting",
    "open_end_to_setting"
  ),
  value = c(
    nrow(kbart),
    sum(!is.na(kbart$online_identifier_kbart) & kbart$online_identifier_kbart != ""),
    sum(!is.na(kbart$print_identifier_kbart) & kbart$print_identifier_kbart != ""),
    sum(!is.na(kbart$identifier_used_for_lookup_kbart) & kbart$identifier_used_for_lookup_kbart != ""),
    nrow(source_matches),
    n_distinct(source_matches$kbart_row_id),
    nrow(kbart) - n_distinct(source_matches$kbart_row_id),
    nrow(counts_by_kbart_row),
    nrow(query_log),
    nrow(query_issues),
    !is.null(test_n),
    ifelse(is.null(test_n), NA_character_, as.character(test_n)),
    ifelse(is.na(open_end_to), NA_character_, open_end_to)
  )
)


# ------------------------------------------------------------
# SECTION 9: Export all sheets to Excel
# ------------------------------------------------------------

# ------------------------------------------------------------
# XML-safe text sanitization BEFORE workbook export
# ------------------------------------------------------------


sanitize_for_excel_xml <- function(x) {
  # Force to character in case factors or other text-like objects appear
  x <- as.character(x)
  
  # Work only on non-missing values
  non_na <- !is.na(x)
  
  # First: coerce/clean to valid UTF-8
  # from = "" means "use current/native encoding"
  # sub = "" drops bytes that cannot be converted
  x[non_na] <- iconv(x[non_na], from = "", to = "UTF-8", sub = "")
  
  # Recompute non-missing values in case iconv produced any NA values
  non_na <- !is.na(x)
  
  # Second: remove XML-illegal low control characters
  x[non_na] <- gsub("[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F]", "", x[non_na], perl = TRUE)
  
  # Return cleaned vector
  x
}

counts_by_kbart_row <- counts_by_kbart_row %>%
  mutate(across(where(is.character), sanitize_for_excel_xml))

source_matches <- source_matches %>%
  mutate(across(where(is.character), sanitize_for_excel_xml))

query_log <- query_log %>%
  mutate(across(where(is.character), sanitize_for_excel_xml))

query_issues <- query_issues %>%
  mutate(across(where(is.character), sanitize_for_excel_xml))

summary_tbl <- summary_tbl %>%
  mutate(across(where(is.character), sanitize_for_excel_xml))

preflight_diagnostic <- preflight_diagnostic %>%
  mutate(across(where(is.character), sanitize_for_excel_xml))

# Create a new workbook.
wb <- wb_workbook()

# Add and write the main count sheet.
wb <- wb_add_worksheet(wb, "counts_by_kbart_row")
wb <- wb_add_data(wb, sheet = "counts_by_kbart_row", x = counts_by_kbart_row)

# Add and write the source match sheet.
wb <- wb_add_worksheet(wb, "source_matches")
wb <- wb_add_data(wb, sheet = "source_matches", x = source_matches)

# Add and write the query log sheet.
wb <- wb_add_worksheet(wb, "query_log")
wb <- wb_add_data(wb, sheet = "query_log", x = query_log)

# Add and write the query issues sheet.
wb <- wb_add_worksheet(wb, "query_issues")
wb <- wb_add_data(wb, sheet = "query_issues", x = query_issues)

# Add and write the summary sheet.
wb <- wb_add_worksheet(wb, "summary")
wb <- wb_add_data(wb, sheet = "summary", x = summary_tbl)

# Add and write the preflight diagnostic sheet.
wb <- wb_add_worksheet(wb, "preflight_diagnostic")
wb <- wb_add_data(wb, sheet = "preflight_diagnostic", x = preflight_diagnostic)

# Save the workbook.
wb_save(wb, file = output_xlsx, overwrite = TRUE)

# Print a completion message.
message("Done. Workbook written to: ", output_xlsx)
