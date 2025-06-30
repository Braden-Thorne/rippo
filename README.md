---
output: github_document
---

<!-- README.md is generated from README.Rmd. Please edit that file -->



# rippo

<!-- badges: start -->
<!-- badges: end -->

The goal of rippo is to ease data entry and tracking of AAGI-CU Intellectual Property and Project Outputs for reporting purposes.

## Installation

You can install the development version of rippo like so:

```r
o <- options() # store default options

options(pkg.build_vignettes = TRUE)

if (!require("pak"))
  install.packages("pak")
pak::pak("AAGI-AUS/rippo")

options(o) # reset options
```

## Example

This is a basic example which shows you how to solve a common problem:

```r
# create a Word document of IPPO tables from each of the project's registers

library(rippo)
create_ippo_tables()
```

## Contributions 

All contributions are appreciated, but please make sure to follow the [Contribution Guidelines](.github/CONTRIBUTING.md).

### Code of Conduct

Please note that the {read.abares} project is released with a [Contributor Code of Conduct](https://AAGI-AUS.github.io/rippo/CODE_OF_CONDUCT.md).
By contributing to this project, you agree to abide by its terms.
