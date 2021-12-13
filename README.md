
<!-- README.md is generated from README.Rmd. Please edit that file -->

# bccashinypivottabler

<!-- badges: start -->

[![R-CMD-check](https://github.com/yzr1996/bccashinypivottabler/workflows/R-CMD-check/badge.svg)](https://github.com/yzr1996/bccashinypivottabler/actions)
<!-- badges: end -->

The `bccashinypivottabler` is based on [shinypivottabler](https://github.com/datastorm-open/shinypivottabler), which enables pivot table to be created with a just few lines. Compared with`shinypivottabler`, `bccashinypivottabler` is simplier to use but less functions.

## Installation

You can install the development version of bccashinypivottabler from
[GitHub](https://github.com/) with:

``` r
# install.packages("devtools")
devtools::install_github("yzr1996/bccashinypivottabler")
```

## Example

This is a basic example which shows you how to use ``bccashinypivottabler' in Shiny App:

``` r
require(bccashinypivottabler)

n <- 10000000

# create artificial dataset
data <- data.frame("gr1" = sample(c("A", "B", "C", "D"), size = n,
                                 prob = rep(1, 4), replace = T),
                   "gr2" = sample(c("E", "F", "G", "H"), size = n,
                                 prob = rep(1, 4), replace = T),
                   "gr3" = sample(c("I", "J", "K", "L"), size = n,
                                 prob = rep(1, 4), replace = T),
                   "gr4" = sample(c("M", "N", "O", "P"), size = n,
                                 prob = rep(1, 4), replace = T),
                   "value1" = rnorm(n),
                   "value2" = runif(n))


server = function(input, output, session) {
  shiny::callModule(module = shinypivottabler2,
                    id = "id",
                    data = data,
                    pivot_cols = c("gr1", "gr2", "gr3", "gr4"))
}

ui = shiny::fluidPage(
  shinypivottablerUI2(id = "id")
)

shiny::shinyApp(ui = ui, server = server)

```
