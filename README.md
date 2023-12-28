# jissutools

<!-- README.md is generated from README.qmd. Please edit that file -->

## Overview

調査実習で使える関数を作ってみました。とりあえずは一つですが、やる気があれば更新していきます。

## Installation

以下のコードを実行してください

``` r
if(!require(remotes)) install.packages("remotes")
remotes::install_github("Kentaro-Kamada/jisshutools")
```

## Usage

- 2変数のクロス表

``` r
library(jisshutools)

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
