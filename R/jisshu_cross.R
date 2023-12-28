library(tidyverse)
library(gtsummary)
library(DescTools)
library(janitor)
library(gt)
library(openxlsx2)

fnc <- function(.data, .x, .y, cramer = TRUE, p.value = TRUE){
  .x <- enquo(.x)
  .y <- enquo(.y)

  .contents_.y <- pull(.data, !!.y)
  if(any(class(.contents_.y) == 'factor')){
    .contents_.y <-
      fct_na_value_to_level(.contents_.y, level = 'NA_') |>
      fct_unique() |>
      as.character()
  } else {
    .contents_.y <-
      factor(.contents_.y) |>
      fct_na_value_to_level(level = 'NA_') |>
      fct_unique() |>
      as.character()
  }

  # クロス表作成
  .tabyl <- tabyl(.data, !!.x, !!.y)

  # 度数を抽出しておく
  N <-
    .tabyl |>
    adorn_totals(where = c('row', 'col')) |>
    pull(Total)

  # 統計量
  if(any(is.na(select(.data, !!.x, !!.y)))){
    .p.value <- NA
    .cramer <- NA
    .sig <- '判定不能'
  } else {
    .p.value <- chisq.test(.tabyl) |> pluck('p.value')
    .cramer <- DescTools::CramerV(.tabyl)
    .sig <- case_when(
      .p.value < 0.001 ~ '0.1％水準で有意',
      .p.value < 0.01 & .p.value >= 0.001 ~ '1％水準で有意',
      .p.value < 0.05 & .p.value >= 0.01 ~ '5％水準で有意',
      .default = '有意差なし'
    )
  }

  .crosstab_raw <-
    .tabyl |>
    adorn_totals(where = c('row', 'col')) |>
    adorn_percentages(denominator = 'row') |>
    as_tibble() |>
    # NAを「NA_」に置換
    mutate(across(.cols = 1, .fns = \(x) replace_na(x, replace = 'NA_'))) |>
    # 「Total」を「合計」に置換
    mutate(across(.cols = 1, .fns = \(x) case_match(x, 'Total' ~ '合計', .default = x))) |>
    rename(合計 = Total) |>
    # xのほうに「（％）」を追加
    mutate(across(.cols = 1, .fns = \(x) str_c(x, '（％）'))) |>
    # 度数を追加
    mutate(N = str_c('（', N, '）')) |>
    # formatting
    mutate(across(where(is.numeric), \(x) sprintf(fmt = '%.1f', x*100)))

  # 統計量を追加
  txt <- str_glue('クラメールのV：{sprintf(fmt = "%.3f", .cramer)}　　{.sig}　{scales::pvalue(.p.value, add_p = TRUE)}')

  # .wb <-
  #   wb_workbook() |>
  #   # シート追加
  #   wb_add_worksheet(sheet = 'クロス表') |>
  #   # データ追加
  #   wb_add_data(x = combined_table, start_row = 2) |>

  return(.crosstab_raw)
}


format.pval(c(0.0001, 0.001, 0.01, 0.05, 0.1, 0.2, 0.3), digits = 3, eps = 0.0001)
scales::pvalue(c(0.0001, 0.001, 0.01, 0.05, 0.1, 0.2, 0.3), add_p = TRUE)


a <- fnc(.data = starwars, .x = sex, .y = gender, cramer = TRUE, p.value = TRUE)

ycont <- starwars$gender |>
  fct_na_value_to_level(level = 'NA_') |>
  fct_unique() |>
  as.character()

wb <-
  wb_workbook() |>
  wb_add_worksheet(sheet = 'クロス表') |>
  wb_add_data(x = a, start_row = 2) |>
  # ヘッダーの追加
  wb_add_data(x = 'gender', dims = 'B1')

sheetdata <- wb_data(wb)


wb$add_data(x = 'ねこ', dims = wb_dims(from_row = dim(sheetdata)[1] + 2))


wb$merge_cells(dims = wb_dims(rows = 1, cols = 2:(1 + length(ycont))))

wb$open()
