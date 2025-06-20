#' Collate AAGI-CU IPPO register Excel workbooks for reporting
#'
#' Traverses directories and imports IPPO register Excel(TM) workbook files and
#'  collates them into a single unified reporting format.
#'
#' @param dir_path_in A string value that provides the path to the top-level
#'  directory of the R: drive holding the AAGI-CU Service and Support project
#'  files.
#' @param dir_path_out A string value that provides the path to the desired
#'  directory to write the output file to. _It will be created if it does
#'  not exist_. The file will be named "AAGI-CU-IPPO Register.docx" by default.
#'  Any currently existing IPPO register will be overwritten.
#'
#' @examplesIf interactive()
#'  # for macOS
#'  library(fs)
#'  R_drive <- "/Volumes/0-9/AAGI_CCDM_CBADA-GIBBEM-SE21982/",
#'  collate_ippo(
#'    dir_path_in = path(R_drive, "Projects"),
#'    dir_path_out = path(R_drive, "Reports")
#'  )
#'
#'  @returns Nothing, called for its side-effects of generating an MS Word(TM)
#'   Document of \acronym{IPPO} details for the purposes of \acronym{AAGI}
#'   reporting.
#'

collate_ippo <- function(dir_path_in, dir_path_out) {
  if (isFALSE(fs::dir_exists(dir_path_in))) {
    cli::cli_abort("{.var dir_path_in} does not exist; cannot proceed")
  }

  if (isFALSE(fs::dir_exists(dir_path_out))) {
    cli::cli_warn("{.var dir_path_out} does not exist; it will be created")
    fs::dir_create(dir_path_out)
  }
}
