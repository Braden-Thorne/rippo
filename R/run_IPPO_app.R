#' Run the Shiny IPPO app
#'
#' Runs a shiny app where a user can build, edit and download an IPPO registry.
#'
#' @importFrom utils installed.packages
#'
#' @examples
#' \dontrun{
#' run_IPPO_app()
#' }
#' @export
run_IPPO_app <- function() {
  appDir <- system.file("IPPO_app", package = "rippo")
  
  installed_pkgs <- rownames(installed.packages())
  not_found <- setdiff(c("shiny", "shinyWidgets", "bslib", "dplyr", "openxlsx"), grep("(^shiny$)|(^shinyWidgets$)|(^bslib$)|(^dplyr$)|(^openxlsx$)", installed_pkgs, value = TRUE))
  if (length(not_found) > 0) {
    stop("Package ", not_found[1], " is needed for this function to work. Please install it.", call. = FALSE)
  }
  
  if (appDir == "") {
    stop("Could not find IPPO_app. Try re-installing `rippo`.", call. = FALSE)
  }
  else {
    ### Can add a resource path here if required using shiny::addResourcePath('path', system.file('path', package = 'rippo'))
    if(interactive()) {
      shiny::runApp(appDir)
    }
    else(return(FALSE))
  }
}