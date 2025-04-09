#' Create a Table of Third Party IP (Table 3) of the R Packages and Quarto Template Used for an AAGI IPPO Document
#'
#' @param project_path The path to the directory of the project to generate the
#'  tables for, defaults to the current working directory and recurses into
#'  subfolders.
#' @param file_path The path to the directory where the table will be saved,
#'  defaults to the project's "1 Documentation" directory.
#' @param quarto Logical, whether to site the \acronym{AAGI} Quarto template.
#' Defaults to `FALSE`.
#'
#' @examples
#' create_table_3()
#'
#' @return nothing, called for its side effects of writing an Excel file to disk
#' for inclusion in the IPPO document spreadsheet.
#'
#' @author Zhanglong Cao, \email{zhanglong.cao@@curtin.edu.au} and Adam H.
#' Sparks, \email{adam.sparks@@curtin.edu.au}

create_table_3 <- function(
    project_path = getwd(),
    file_path = "1 Documentation",
    quarto = FALSE,
    digger = FALSE) {
  # Check if the path is a valid directory
  if (!dir.exists(project_path)) {
    stop("The specified path does not exist.")
  }

  # Set the working directory to the specified path
  withr::withdir(project_path)

  pkgs <- funspottr::spot_packages(
    path = project_path,
    recursive = TRUE
  )

  if (digger) {
    digger <- data.frame(
      Package = "DiGGer",
      Author = "Neil Coombes [aut, ctb]",
      Description = "searches for A–efficient experimental designs under specified blocking and correlation.",
      License = "NSW DPI Freeware Licence"
    )
  }
}

cran_database <- tools::CRAN_package_db()[, c(
  "Package",
  "Author",
  "Description",
  "License"
)]


db_DiGGer <- data.frame(
  Package = "DiGGer",
  Author = "Neil Coombes [aut, ctb]",
  Description = "searches for A–efficient experimental designs under specified blocking and correlation.",
  License = "NSW DPI Freeware Licence"
)
# db_AAGIThemes <- data.frame(Package = "AAGIThemes",
#                              Author = "AAGI",
#                              Description = "AAGI Branding for R Graphical and Tabular Outputs. An R programming language software package to help ensure consistency in AAGI branding across product outputs from AAGI participants.",
#                              License = "MIT Licence and is available from https://github.com/AAGI-AUS/AAGIThemes")
#
# db_AAGIPalettes <- data.frame(Package = "AAGIPalettes",
#                             Author = "AAGI",
#                             Description = "AAGI Colours and Colour Palettes for R. An R programming language software package to help ensure consistency in AAGI branding across product outputs from AAGI partners",
#                             License = "MIT Licence and is available fromhttps://github.com/AAGI–AUS/AAGIPalettes")

db_Quarto <- data.frame(
  Package = "Quarto",
  Author = "Posit Software, PBC",
  Description = "Quarto, open–source tools for scientific and technical publishing",
  License = "version 1.3 (and earlier) is licensed under the GNU GPL v2. Quarto version 1.4 is licensed under the MIT"
)


cran_database <- rbind(cran_database, db_DiGGer, db_Quarto)


## for example
cran_database[which(cran_database$Package == "Quarto"), ]
cran_database[which(cran_database$Package == "quarto"), ]

cran_database[which(cran_database$Package == "DiGGer"), ]


## default two rows
top2rows <- data.frame(
  No. = 1:2,
  `AAGI Project Code` = c("AAGI-ALL-SP-003", "AAGI-ALL-SP-003"),
  `Owner/s` = c("R Foundation", "Posit Software, PBC"),
  Description = c(
    "R programming language, a statistical analysis environment that provides data manipulation, calculation and graphical display.",
    "RStudio, an integrated development environment for R, a programming language for statistical computing and graphics."
  ),
  `Date made available to Project` = as.Date(c(
    "2023-07-18",
    "2023-07-18"
  )),
  `Name of party making Third Party IP available (if not the owner(s))` = c(
    "CU",
    "CU"
  ),
  `Arrangements applicable to the provision of Third Party IP for the Project` = c(
    "R is made freely available for use and redistribution under the Free Software Foundation’s GNU General Public Licence v2.0.",
    "RStudio is made freely available for use and distribution under the GNU AFFERO GENERAL PUBLIC LICENSE Version 3, 19 November 2007"
  ),
  `Restrictions / limitations on use for dissemination or Commercialisation of Project Outputs` = c(
    "Subject to the terms of the GNU General Public Licence v2.0.",
    "Subject to the terms of the GNU AFFERO GENERAL PUBLIC LICENSE Version 3, 19."
  )
)

### extract package information from CRAN database
add_package_info <- function(df, db, packages) {
  for (pkg in packages) {
    pkg_info <- db[db$Package == pkg, ]
    if (nrow(pkg_info) > 0) {
      author_clean <- gsub("\\s*<[^>]+>", "", pkg_info$Author)
      author_clean <- gsub("\\(\\)", "", author_clean)
      author_clean <- gsub(
        "\\s+",
        " ",
        gsub("\n", " ", author_clean)
      )
      description_clean <- gsub(
        "\\s*<[^>]+>",
        "",
        pkg_info$Description
      )
      description_clean <- gsub(
        "\\(\\)",
        "",
        description_clean
      )
      description_clean <- gsub(
        "\\s+",
        " ",
        gsub("\n", " ", description_clean)
      )

      new_row <- data.frame(
        No. = nrow(df) + 1,
        `AAGI Project Code` = "AAGI-ALL-SP-003",
        `Owner/s` = author_clean,
        Description = paste0(
          "{",
          pkg_info$Package,
          "} ",
          description_clean
        ),
        `Date made available to Project` = as.Date(
          "2023-07-18"
        ),
        `Name of party making Third Party IP available (if not the owner(s))` = "CU",
        `Arrangements applicable to the provision of Third Party IP for the Project` = paste0(
          "{",
          pkg_info$Package,
          "} is made freely available for use and redistribution under the",
          pkg_info$License,
          "Licence"
        ),
        `Restrictions / limitations on use for dissemination or Commercialisation of Project Outputs` = paste(
          "Subject to the terms of the",
          pkg_info$License,
          "Licence."
        )
      )
      df <- rbind(df, new_row)
    }
  }
  return(df)
}


extract_library_packages <- function(code) {
  matches <- gregexpr("library\\(([^)]+)\\)", code)
  packages <- regmatches(code, matches)
  packages <- unlist(packages)
  packages <- gsub("library\\(|\\)|#", "", packages)
  packages <- trimws(packages)
  packages <- unique(packages)
  return(packages)
}

mypackages <- "
library(readxl)
library(dplyr)
library(stringr)
library(tidyr)
library(ggplot2)
library(shiny)
library(shinydashboard)
library(sf)
library(leaflet)
library(shinycssloaders)
library(highcharter)
library(purrr)
library(DT)
library(scales)
library(RColorBrewer)
library(quarto)
"


package_names <- extract_library_packages(mypackages)

table3_packages <- add_package_info(top2rows, cran_database, package_names)

write.csv(table3_packages, file = "IPPO_Table3.csv", row.names = FALSE)


######## a list of projects ##############

myprojects <- read.csv("MyProjectList.csv")
myprojects <- myprojects[myprojects$Packages != "", ]

for (i in 1:nrow(myprojects)) {
  project_code <- myprojects$ProjectCode[i]
  packages <- myprojects$Packages[i]
  package_names <- extract_library_packages(packages)
  table3_packages <- add_package_info(
    top2rows,
    cran_database,
    package_names
  )

  output_file <- file.path(paste0("Table3/Table3_", project_code, ".csv"))
  write.csv(table3_packages, output_file, row.names = FALSE)
}


## additional packages
library(AAGIThemes)
library(AAGIPalettes)
library(DiGGer)
