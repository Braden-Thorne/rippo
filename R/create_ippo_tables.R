#' Create IPPO tables list
#'
#' Traverses directories and imports IPPO register Excel(TM) workbook files and
#'  creates a list of tables corresponding to Tables 1-5 in the \acronym{AAGI}
#'  \acronym{IPPO} register for all AAGI-CU Service and Support projects.
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
#'  create_ippo_tables(
#'    dir_path_in = path(R_drive, "Projects"),
#'    dir_path_out = path(R_drive, "Reports")
#'  )
#'
#' @returns A `list` object that contains five [tibble::tibble()] objects that
#'  contain either tables for the IPPO register or a list of projects lacking
#'  an IPPO register.
#'

create_ippo_tables <- function(dir_path_in, dir_path_out) {
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
    ippo_paths <- file.path(c(completed, active), "/1 Documentation/")
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

    # TODO: remove any temporary Excel files, e.g., those starting a filename
    # with a "~" that exist in the directory.

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
    names(merged) <- project_names

    has_ippo <- merged[grep(".xlsx", merged)]
    no_ippo <- names(merged[names(merged) %notin% names(has_ippo)])

    table_1 <- lapply(
        X = has_ippo,
        FUN = readxl::read_excel,
        sheet = 2L,
        col_types = c("numeric", "text", "text", "date", "text")
    )
    table_2 <- lapply(
        X = has_ippo,
        FUN = readxl::read_excel,
        sheet = 3L,
        col_types = c("numeric", "text", "text", "date", "text")
    )
    table_3 <- lapply(
        X = has_ippo,
        FUN = readxl::read_excel,
        sheet = 4L,
        col_types = c(
            "numeric",
            "text",
            "text",
            "text",
            "date",
            "text",
            "text",
            "text"
        )
    )
    table_4 <- lapply(
        X = has_ippo,
        FUN = readxl::read_excel,
        sheet = 5L,
        col_types = c("numeric", "text", "text", "date", "text", "text")
    )
    table_5 <- lapply(
        X = has_ippo,
        FUN = readxl::read_excel,
        sheet = 6L,
        col_types = c("numeric", "text", "text", "date", "text", "text")
    )

    tables <- list(table_1, table_2, table_3, table_4, table_5)
    tables_out <- lapply(
        X = tables,
        FUN = purrr::list_rbind,
        names_to = "Project Code"
    )

    return(list("tables" = tables, "No_IPPO" = no_ippo))
}
