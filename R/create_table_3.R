#' Create a Table of Third Party IP (Table 3) of the R Packages and Quarto Template Used for an AAGI IPPO Document
#'
#' @param project_path The path to the directory of the project to generate the
#'  tables for, defaults to the current working directory and recurses into
#'  subfolders.
#' @param file_path The path to the directory where the table will be saved,
#'  defaults to the project's "1 Documentation" directory.
#' @param quarto Logical, whether to cite the \acronym{AAGI} Quarto template.
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
  digger = FALSE
) {
  # Check if the path is a valid directory
  if (!dir.exists(project_path)) {
    cli::cli_abort("The specified path does not exist.")
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
      Description = "Searches for efficient experimental designs under specified blocking and correlation.",
      License = "NSW DPI Freeware Licence"
    )
  }
}
