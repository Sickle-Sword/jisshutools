# jissutools


<!-- README.md is generated from README.qmd. Please edit that file -->

## Overview

調査実習で使える関数を作ってみました。やる気があれば多項ロジットなどもつくります。

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

- 回帰分析（以下をサポート）
  - 線形回帰：`lm`
  - 2項ロジスティック回帰：`glm`
  - 多項ロジスティック回帰：`nnet::multinom`

``` r
data <- tibble::tibble(
  x1 = rnorm(100, mean = 0, sd = 1),
  x2 = rnorm(100, mean = 0, sd = 1),
  y = 0.2 + 0.3*x1 + 0.5*x2 + rnorm(100, mean = 0, sd = 0.1),
  y_bin = rbinom(100, 1, plogis(y)),
  y_multinom = cut(
    y,
    breaks = quantile(y, probs = c(0, 0.25, 0.75, 1)),
    labels = c('Q1', 'Q2', 'Q3'),
    include.lowest = TRUE
  )
)


# Create a regression table
# Linear regression
model_lm <- lm(y ~ x1 + x2, data = data)
result_lm <- jisshu_reg(model_lm)

# Logistic regression
model_glm <- glm(y_bin ~ x1 + x2, data = data, family = binomial(link = 'logit'))
result_glm <- jisshu_reg(model_glm)

# Multinomial logistic regression
model_multinom <- nnet::multinom(y_multinom ~ x1 + x2, data = data, model = TRUE)
result_multinom <- jisshu_reg(model_multinom)

# Glance the regression table
result_lm$open()
result_glm$open()
result_multinom$open()


# Save the regression table as an excel file
result_lm$save('hoge_lm.xlsx')
result_glm$save('hoge_glm.xlsx')
result_multinom$save('hoge_multinom.xlsx')
```
