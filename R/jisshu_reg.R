#' regression table for tyosa jisshu
#'
#' 調査実習用の回帰分析結果を作成する関数です。
#'
#' `lm`や`glm`の結果から、執筆要項に沿ったエクセルの表が作成できます。
#'
#' エクセルへの保存はexampleを参照してください。
#'
#' @param object an object for which a summary is desired. For example, a model object, such as those returned by `lm` and `glm`.
#'
#' @examples
#' # Create a sample data
#' data <- tibble::tibble(
#'   x1 = rnorm(100, mean = 0, sd = 1),
#'   x2 = rnorm(100, mean = 0, sd = 1),
#'   y = 3 + 0.3*x1 + 0.5*x2 + rnorm(100, mean = 0, sd = 0.1)
#' )
#'
#' # Create a regression table
#' model <- lm(y ~ x1 + x2, data = data)
#' result <- jisshu_reg(model)
#'
#' # Save the regression table as an excel file
#' \dontrun{
#' result$save('hoge.xlsx')
#' }
#'
#' @export
#'

jisshu_reg <- function(object) {
  UseMethod('jisshu_reg')
}

#'
#' @rdname jisshu_reg
#' @import dplyr
#' @importFrom broom tidy glance
#' @importFrom scales number
#' @importFrom tibble tibble
#' @importFrom openxlsx2 wb_workbook wb_add_worksheet wb_add_data wb_dims wb_data
#'
#' @export
#'

jisshu_reg.lm <- function(object) {

  .glance_tab <- glance(object)
  .coef_tab <-
    tidy(object) |>
    mutate(
      p.value = pvalue_to_star(p.value),
      term = case_match(
        term,
        '(Intercept)' ~ '（定数）',
        .default = term
      )
    ) |>
    select(独立変数 = term, 偏回帰係数 = estimate, 標準誤差 = std.error, p.value) |>
    mutate(across(where(is.numeric), \(x) scales::number(x, accuracy = 0.001)))

  .fit_stats <-
    tibble(
      独立変数 = c('自由度調整済み決定係数', 'F値', 'N'),
      偏回帰係数 = c(
        .glance_tab$adj.r.squared |> scales::number(accuracy = 0.001),
        .glance_tab$statistic |> scales::number(accuracy = 0.001, big.mark = ''),
        .glance_tab$nobs
      ),
      標準誤差 = c(
        NA,
        pvalue_to_star(.glance_tab$p.value),
        NA
      ),
    )

  .wb <-
    wb_workbook() |>
    # シート追加
    wb_add_worksheet(sheet = '回帰分析') |>
    # データ追加
    wb_add_data(x = bind_rows(.coef_tab, .fit_stats), start_row = 1, na.strings = '') |>
    # p.valueを空欄に変更
    wb_add_data(x = '', dims = 'D1')

  # シートデータ取得
  .sheetdata <- wb_data(.wb)

  # 注を追加
  .wb$add_data(
    x = '注：*: p<0.05、**: p<0.01、***: p<0.001。',
    dims = wb_dims(from_row = 1 + nrow(.sheetdata) + 1)
  )

  # 注のセル結合
  .wb$merge_cells(dims = wb_dims(rows = 1 + nrow(.sheetdata) + 1, cols = 1:ncol(.sheetdata)))

  # 中央揃え
  .wb$add_cell_style(
    dims = wb_dims(x = .sheetdata, select = 'data', cols = '偏回帰係数'),
    apply_alignment = TRUE,
    horizontal = 'center',
    vertical = 'center'
  )$
    # 中央揃え
    add_cell_style(
      dims = wb_dims(
        x = .sheetdata, select = 'data', cols = '標準誤差', rows = 1:(nrow(.sheetdata) - 3)
      ),
      apply_alignment = TRUE,
      horizontal = 'center',
      vertical = 'center'
    )$
    # 中央揃え
    add_cell_style(
      dims = wb_dims(x = .sheetdata, select = 'col_names'),
      apply_alignment = TRUE,
      horizontal = 'center',
      vertical = 'center'
    )

  # 罫線
  .wb$add_border(
    dims = wb_dims(x = .sheetdata, select = 'x'),
    top_border = 'thin', bottom_border = 'thin', left_border = NULL, right_border = NULL
  )$
    add_border(
      dims = wb_dims(x = .sheetdata, select = 'data', rows = 1:(nrow(.sheetdata) - 3)),
      top_border = 'thin', bottom_border = 'thin', left_border = NULL, right_border = NULL
    )

  class(.wb) <- append(class(.wb), 'jisshu_reg.lm', after = 1)

  print(bind_rows(.coef_tab, .fit_stats))
  return(.wb)
}

