#' 2 way cross table for chosa jisshu
#'
#' 調査実習用の2次元クロス表を作成する関数です。
#'
#' データと変数2つを指定するとクロス表が作成されます。
#'
#' クラメールのVとカイ2乗検定のp値も算出されます。2×2のクロス表の場合はイェーツの連続修正が入ります。
#'
#' エクセルへの保存はexampleを参照してください。
#'
#' @param .data a `tibble` or `data.frame`
#' @param .x a column name in `.data`
#' @param .y a column name in `.data`
#' @return an object of class `wbWorkbook` exported from `openxlsx2` package. The object can be easily saved as an excel file. See examples for details.
#'
#' @examples
#' # Create a sample data
#' data <-
#'   data.frame(
#'     x = c(rep('A', 50), rep('B', 50)),
#'     y = c(rep(1:5, 20))
#'   )
#'
#' # Create a cross table
#' result <- data |> jisshu_cross(x, y)
#'
#' # Save the cross table as an excel file
#' \dontrun{
#' result$save('hoge.xlsx')
#' }
#'
#' @import dplyr
#' @importFrom rlang enquo as_label
#' @importFrom forcats fct_drop fct_unique fct_na_value_to_level
#' @importFrom stringr str_glue str_c
#' @importFrom tidyr pivot_wider replace_na
#' @importFrom tibble enframe as_tibble
#' @importFrom scales number pvalue
#' @importFrom DescTools CramerV
#' @importFrom janitor tabyl adorn_totals adorn_percentages chisq.test
#' @importFrom openxlsx2 wb_workbook wb_add_worksheet wb_add_data wb_dims wb_merge_cells wb_data
#'
#'@export
#'


jisshu_cross <- function(.data, .x, .y) {
  .x <- enquo(.x)
  .y <- enquo(.y)

  .contents_.y <- pull(.data, !!.y)
  if(!inherits(.contents_.y, 'factor')) .contents_.y <- factor(.contents_.y)
  .contents_.y <-
    fct_na_value_to_level(.contents_.y, level = 'NA_') |>
    fct_drop() |>
    fct_unique() |>
    as.character()

  # クロス表作成
  .tabyl <- tabyl(.data, !!.x, !!.y)

  # 度数を抽出しておく
  N <-
    .tabyl |>
    adorn_totals(where = c('row', 'col')) |>
    pull(Total)

  # 統計量
  if(any(is.na(select(.data, !!.x, !!.y)))){
    .txt <- 'NAが含まれているため、統計量は算出できません。'
  } else {
    .p.value <- chisq.test(.tabyl, correct = TRUE)$p.value
    .cramer <- CramerV(pull(.data, !!.x), pull(.data, !!.y))
    .sig <- case_when(
      .p.value < 0.001 ~ '0.1％水準で有意',
      .p.value < 0.01 & .p.value >= 0.001 ~ '1％水準で有意',
      .p.value < 0.05 & .p.value >= 0.01 ~ '5％水準で有意',
      .default = '有意差なし'
    )
    # 統計量を追加
    .txt <- str_glue('クラメールのV：{sprintf(fmt = "%.3f", .cramer)}　　{.sig}　{pvalue(.p.value, add_p = TRUE)}') |> as.character()
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
    mutate(across(where(is.numeric), \(x) number(x*100, accuracy = 0.1)))

  # excelに書き出し
  .wb <-
    wb_workbook() |>
    # シート追加
    wb_add_worksheet(sheet = 'クロス表') |>
    # データ追加
    wb_add_data(x = .crosstab_raw, start_row = 2)

  # ヘッダーの追加
  .header <-
    c(as_label(.x), rep(as_label(.y), length(.contents_.y)), '合計', 'N') |>
    enframe() |>
    pivot_wider()
  .wb$add_data(
    x = .header, col_names = FALSE,
    dims = wb_dims(x = .header, select = 'col_names')
  )

  # p値などを追加（ヘッダー2行 + 行数 + 1）
  .wb$add_data(x = .txt, dims = wb_dims(from_row = 2 + nrow(.crosstab_raw) + 1))

  # セル結合
  .wb$merge_cells(dims = 'A1:A2')$
    merge_cells(dims = wb_dims(rows = 1:2, cols = 2 + length(.contents_.y)))$
    merge_cells(dims = wb_dims(rows = 1:2, cols = 3 + length(.contents_.y)))$
    # 列名のセル結合
    merge_cells(dims = wb_dims(rows = 1, cols = 2:(1 + length(.contents_.y))))$
    # p値などのセル結合
    merge_cells(dims = wb_dims(from_row = 2 + nrow(.crosstab_raw) + 1, cols = 1:ncol(.crosstab_raw)))

  # セルの書式設定
  # すべて中央揃え
  .wb$add_cell_style(
    dims = wb_dims(x = wb_data(.wb), select = 'x', from_row = 1),
    apply_alignment = TRUE,
    horizontal = 'center',
    vertical = 'center',
  )$
    # xのカテゴリのみ左揃え
    add_cell_style(
      dims = wb_dims(cols = 1, rows = 3:(2 + nrow(.crosstab_raw))),
      apply_alignment = TRUE,
      horizontal = 'left',
      vertical = 'center',
    )$
    # p値などのセルのみ右揃え
    add_cell_style(
      dims = wb_dims(from_row = 2 + nrow(.crosstab_raw) + 1),
      apply_alignment = TRUE,
      horizontal = 'right',
      vertical = 'center',
    )

  # 罫線
  .wb$add_border(
    dims = wb_dims(x = wb_data(.wb), select = 'x'),
    top_border = 'thin', bottom_border = 'thin', left_border = NULL, right_border = NULL
  )$
    add_border(
      dims = wb_dims(cols = 1:ncol(wb_data(.wb)), rows = nrow(wb_data(.wb))),
      top_border = NULL, bottom_border = 'thin', left_border = NULL, right_border = NULL
    )$
    add_border(
      dims = wb_dims(rows = 2, cols = 1:ncol(wb_data(.wb))),
      top_border = NULL, bottom_border = 'thin', left_border = NULL, right_border = NULL
    )$
    add_border(
      dims = wb_dims(rows = 1, cols = 2:(1 + length(.contents_.y))),
      top_border = 'thin', bottom_border = 'thin', left_border = NULL, right_border = NULL
    )

  print(list(.crosstab_raw, .txt))
  return(.wb)
}


