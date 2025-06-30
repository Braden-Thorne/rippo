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

  if (appDir == "") {
    cli::cli_abort("Could not find {.code IPPO_app}. Try re-installing {.pkg rippo}.")
  }
  else {
    ### Can add a resource path here if required using shiny::addResourcePath('path', system.file('path', package = 'rippo'))
    if(interactive()) {
      shiny::runApp(appDir)
    }
    else(return(FALSE))
  }
}
