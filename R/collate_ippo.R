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
#'  R_drive <- "/Volumes/0-9/AAGI_CCDM_CBADA-GIBBEM-SE21982/"
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

  # create a list of completed projects that are no longer actively supported
  completed <- fs::dir_ls(path(dir_path_in, "02 Archived Completed"))

  # create a list of active projects and filter out non-proj dirs
  active <- fs::dir_ls(path(dir_path_in))
  active <- active[!grepl("Projects/\\d{2} ", active)]
  active <- active[!grepl("Projects/AAGI", active, fixed = TRUE)]
  active <- active[!grepl("Projects/RiskWise Program", active, fixed = TRUE)]

  # point to the IPPO registers
  ippo_paths <- paste0(c(completed, active), "/1 Documentation/")
  ippo_registers <- lapply(
    X = ippo_paths,
    FUN = fs::dir_ls,
    regexp = "AAGI-CU-.*IPPO.*\\.xlsx$"
  )

  merged <- Map(
    function(x, y) if (length(x) == 0L) y else x,
    ippo_registers,
    ippo_paths
  )

  # Regular expression pattern
  pattern <- "(?<=/Projects/).*?(?=/1 Documentation/)"

  # Apply regex using vapply
  project_names <- vapply(
    merged,
    function(x) {
      regmatches(x, regexpr(pattern, x, perl = TRUE))
    },
    character(1L)
  )
  project_names <- gsub(
    pattern = "02 Archived Completed/",
    replacement = "Completed - ",
    x = project_names
  )
}
