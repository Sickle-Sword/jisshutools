#'
#' @rdname jisshu_reg
#' @import dplyr
#' @importFrom broom tidy glance
#' @importFrom scales number
#' @importFrom tibble tibble
#' @importFrom rlang f_lhs as_name
#' @importFrom stringr str_glue
#' @importFrom purrr pluck_exists
#' @importFrom openxlsx2 wb_workbook wb_add_worksheet wb_add_data wb_dims wb_data
#'
#' @exportS3Method jisshutools::jisshu_reg
#'


jisshu_reg.multinom <- function(object) {

  # nnet::multinomで`model = TRUE`になっているかチェック
  if(!pluck_exists(object, 'model')) stop('nnet::multinom()で`model = TRUE`を設定してください。')

  # 従属変数の変数名と参照カテゴリ
  y_name <- as_name(f_lhs(object$terms))
  y_ref <- object$lab[1]

  # nullモデルを推定し、統計量を算出
  nullmodel <-
    update(object, ~1, data = object$model, trace = FALSE) |>
    glance() |>
    rename_with(\(x) paste0(x, '_null'))

  .glance_tab <-
    bind_cols(
      glance(object),
      nullmodel
    ) |>
    # devianceとdfの差分を計算
    mutate(
      diff.deviance = abs(deviance_null - deviance),
      diff.df = abs(edf_null - edf),
      p.value = pchisq(diff.deviance, diff.df, lower.tail = FALSE),
      Nagelkerke = (1 - exp((deviance - deviance_null)/nobs))/(1 - exp(-deviance_null/nobs))
    )

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
    select(!!y_name := y.level, 独立変数 = term, 偏回帰係数 = estimate, 標準誤差 = std.error, p.value) |>
    mutate(across(where(is.numeric), \(x) number(x, accuracy = 0.001)))

  .fit_stats <-
    tibble(
      独立変数 = c('Nagelkerke決定係数', 'モデルχ2乗値', 'N'),
      偏回帰係数 = c(
        number(.glance_tab$Nagelkerke, accuracy = 0.001),
        number(.glance_tab$diff.deviance, accuracy = 0.001, big.mark = ''),
        .glance_tab$nobs
      ),
      標準誤差 = c(
        NA,
        pvalue_to_star(.glance_tab$p.value),
        NA
      ),
    )

  # エクセルの表作成
  .wb <-
    wb_workbook() |>
    # シート追加
    wb_add_worksheet(sheet = '回帰分析') |>
    # データ追加
    wb_add_data(x = bind_rows(.coef_tab, .fit_stats), start_row = 1, na.strings = '')

  # シートデータ取得
  .sheetdata <- wb_data(.wb)

  # セルの書き換え
  # p.valueを空欄に変更
  .wb$add_data(x = '', dims = 'E1')$
    # 従属変数の名前と参照カテゴリを追加
    add_data(
      x = str_glue('{y_name}（参照カテゴリ：{y_ref}）') |>
        as.character(),
      dims = 'A1'
    )

  # 注を追加
  .wb$add_data(
    x = '注：+: p<0.1、*: p<0.05、**: p<0.01、***: p<0.001。',
    dims = wb_dims(from_row = 1 + nrow(.sheetdata) + 1)
  )
  # 注のセル結合
  .wb$merge_cells(dims = wb_dims(rows = 1 + nrow(.sheetdata) + 1, cols = 1:ncol(.sheetdata)))

  # 同じカテゴリの行を特定
  mergecells <-
    bind_rows(.coef_tab, .fit_stats) |>
    # データは2行目から
    mutate(id = row_number() + 1) |>
    summarise(
      variable_index = list(id),
      last_row = max(id),
      .by = !!y_name
    )
  # 同じカテゴリの行をセル結合
  walk(mergecells$variable_index, \(x) .wb$merge_cells(dims = wb_dims(x, 1)))
  # 結合したセルの値を上に揃える
  .wb$add_cell_style(
    dims = wb_dims(x = .sheetdata, select = 'data', cols = y_name),
    apply_alignment = TRUE,
    horizontal = 'left',
    vertical = 'top'
  )

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
  # 結合したセルの最後の行に罫線を追加
  walk(
    mergecells$last_row,
    \(x) .wb$add_border(
      dims = wb_dims(x, 1:ncol(.sheetdata)),
      top_border = NULL, bottom_border = 'thin', left_border = NULL, right_border = NULL
    )
  )

  class(.wb) <- append(class(.wb), 'jisshu_reg.multinom', after = 1)

  print(bind_rows(.coef_tab, .fit_stats))
  return(.wb)
}
