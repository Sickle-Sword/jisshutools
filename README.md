# jissutools

<!-- README.md is generated from README.qmd. Please edit that file -->

## Overview

調査実習で使える関数を作ってみました。とりあえずは一つですが、やる気があれば更新していきます。

## Installation

以下のコードを実行してください

``` r
if(!require(remotes)) install.packages("remotes")
remotes::install_github("Kentaro-Kamada/jisshutools")

library(jisshutools)
```

## Usage

- 2変数のクロス表

``` r
# Create a sample data
data <-
  data.frame(
    x = c(rep('A', 50), rep('B', 50)),
    y = c(rep(1:5, 20))
  )

# Create a cross table
result <- data |> jisshu_cross(x, y)

# Save the cross table as an excel file
result$save('hoge.xlsx')
```

- 線形回帰

``` r
# Create a sample data
data <- tibble::tibble(
  x1 = rnorm(100, mean = 0, sd = 1),
  x2 = rnorm(100, mean = 0, sd = 1),
  y = 3 + 0.3*x1 + 0.5*x2 + rnorm(100, mean = 0, sd = 0.1)
)

# Create a regression table
model <- lm(y ~ x1 + x2, data = data)
result <- jisshu_reg(model)

# Save the regression table as an excel file
result$save('hoge.xlsx')
```
