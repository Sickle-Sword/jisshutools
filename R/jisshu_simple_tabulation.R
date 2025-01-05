#' Simple Tabulation for Chosa Jisshu
#'
#' 調査実習用の単純集計表を作成する関数です。
#'
#' @param data a `tibble` or `data.frame`
#' @param path a file path to save the excel file
#' @param grade specify the grade (i.e. school year) variable. It must be a factor variable with the levels '1年', '2年', '3年'.
#' @param gender specify the gender variable. It must be a factor variable with the levels '男', '女'.
#' @param include variables to include in the table. Specify the variables in tidy-select style like `dplyr::select()`. Default is `everything()`, e.g. all variables.
#' @param cont_var continuous variables to compute mean instead of frequency. Specify the variables in tidy-select style like `dplyr::select()`. Default is `NULL`.
#'
#' @examples
#' library(dplyr)
#' library(haven)
#' data <-
#'   tibble(
#'     ID = 1:100,
#'     grade =
#'       sample(1:3, 100, replace = TRUE, prob = c(0.3, 0.3, 0.4)) |>
#'       labelled_spss(labels = c('1年' = 1, '2年' = 2, '3年' = 3), label = '学年'),
#'     gender =
#'       sample(c(1, 2, 9), 100, replace = TRUE, prob = c(0.45, 0.45, 0.1)) |>
#'       labelled_spss(labels = c('男' = 1, '女' = 2, '無回答' = 9), label = '性別'),
#'     Q1 = rnorm(100, mean = 0, sd = 1),
#'     Q1_na = sample(c('無回答', '非該当', '有効'), 100, replace = TRUE, prob = c(0.1, 0.1, 0.8)),
#'     Q2 =
#'       sample(c(1:4, 8, 9), 100, replace = TRUE, prob = c(0.2, 0.3, 0.3, 0.1, 0.1, 0.1)) |>
#'       labelled_spss(
#'         labels = c(
#'           'あてはまる' = 1,
#'           'まああてはまる' = 2,
#'           'あまりあてはまらない' = 3,
#'           'まったくあてはまらない' = 4,
#'           '無回答' = 9,
#'           '非該当' = 8
#'         ),
#'         label = 'あなたは猫好きである'
#'       ),
#'   ) |>
#'   mutate(
#'     Q1 = case_when(
#'       Q1_na == '無回答' ~ 9,
#'       Q1_na == '非該当' ~ 8,
#'       .default = Q1
#'     ) |>
#'       labelled_spss(labels = c('非該当' = 8, '無回答' = 9), label = 'あなたの猫度を答えてください')
#'   ) |>
#'   select(!Q1_na)
#'
#' data
#'
#' count(data, grade)
#' count(data, gender)
#' count(data, Q1, sort = TRUE)
#' count(data, Q2)
#'
#' \dontrun{
#' jisshu_simple_tabulation(
#'   data |> as_factor(),
#'   'neko.xlsx',
#'   grade = grade,
#'   gender = gender,
#'   include = !c(ID),
#'   cont_var = c(Q1)
#' )
#' }
#'
#' @import dplyr
#' @importFrom purrr map imap list_rbind walk
#' @importFrom tibble tibble as_tibble enframe
#' @importFrom tidyr pivot_wider
#' @importFrom readr parse_double
#' @importFrom stringr str_c
#' @importFrom scales number percent
#' @importFrom janitor tabyl
#' @importFrom tidyselect eval_select
#' @importFrom openxlsx2 wb_workbook wb_add_worksheet wb_add_data wb_dims wb_data
#' @importFrom kamaken haven_variable_label
#'
#' @export
#'

# エクセルに所定の形式で出力
jisshu_simple_tabulation <- function(data, path, grade, gender, include = everything(), cont_var = NULL) {
  # 連続変数の変数名を取得
  cont_var <- tidyselect::eval_select(enexpr(cont_var), data) |> names()

  combined_table <-
    # 全体・学年別・性別データの作成
    prep_data_list(data, include = {{ include }}, grade = {{ grade }}, gender = {{ gender }}) |>
    # 各データセットにおける単純集計表の作成
    map(\(x) simple_tabulation(x, cont_var = cont_var)) |>
    # 単純集計表の結合
    list_rbind(names_to = 'dataset') |>
    # ワイド形式に変換
    pivot_wider(names_from = dataset, values_from = percent, values_fill = '0.0%') |>
    # 変数名と変数ラベル結合
    left_join(
      haven_variable_label(data) |>
        mutate(ラベル = case_when(is.na(ラベル) ~ 変数, .default = ラベル)),
      by = join_by(variable == 変数)
    ) |>
    mutate(variable = str_c(variable, ' ', ラベル) |> coalesce(variable, label)) |>
    # いらない変数削除
    select(!c(位置, ラベル))

  # エクセルファイルに出力
  # ワークブック作成
  wb <-
    wb_workbook() |>
    # シート追加
    wb_add_worksheet(sheet = '単純集計表') |>
    # データ追加
    wb_add_data(x = combined_table, start_row = 2) |>
    # ヘッダーを追加
    wb_add_data(
      x = c('', '', '全体', '学年', '学年', '学年', '性別', '性別') |>
        enframe() |>
        pivot_wider(),
      start_row = 1, col_names = FALSE
    )

  # シートのデータを取得
  sheet_data <- wb_data(wb)

  # ヘッダーセル結合
  wb$merge_cells(dims = 'A1:B2')$
    merge_cells(dims = 'C1:C2')$
    merge_cells(dims = 'D1:F1')$
    merge_cells(dims = wb_dims(rows = 1, cols = 7:ncol(sheet_data)))


  # 変数名のセル結合をする列特定
  mergecell <-
    combined_table |>
    # データは3行目から（ヘッダーが2行分ある）
    mutate(id = row_number() + 2) |>
    # 連続変数と先頭の度数は結合しない
    filter(!label %in% c('度数', cont_var)) |>
    summarise(variable_index = list(id), .by = variable)

  # 変数名セル結合
  walk(mergecell$variable_index, \(x) wb$merge_cells(dims = wb_dims(x, 1)))

  # 連続変数として指定した変数のAB列を結合
  mergecell2 <-
    combined_table |>
    # データは3行目から（ヘッダーが2行分ある）
    mutate(id = row_number() + 2) |>
    # 連続変数のみ残す
    filter(label %in% c('度数', cont_var)) |>
    summarise(variable_index = list(id), .by = variable)

  # 連続変数のセル結合
  walk(mergecell2$variable_index, \(x) wb$merge_cells(dims = wb_dims(x, 1:2)))

  # 変数名を上揃え
  wb$add_cell_style(
    dims = wb_dims(x = sheet_data, cols = 1),
    apply_alignment = TRUE,
    vertical = 'top',
    # 折り返し
    wrap_text = TRUE
  )

  # ヘッダー2行を中央揃え
  wb$add_cell_style(
    dims = wb_dims(rows = 1:2, cols = 1:ncol(sheet_data)),
    apply_alignment = TRUE,
    horizontal = 'center'
  )

  # 罫線追加（グリッド）
  wb$add_border(
    dims = wb_dims(x = sheet_data, select = 'x'),
    inner_hgrid = 'thin',
    inner_vgrid = 'thin'
  )

  # セルの幅設定
  wb$set_col_widths(cols = 1:2, widths = 20)

  # 保存
  wb$save(file = path)
  return(wb)
}


# 全体・学年別・性別データの作成
prep_data_list <- function(data, include, grade, gender) {
  list(
    `全体` = data |> select({{ include }}),
    `1年生` = data |> filter({{ grade }} == '1年') |> select({{ include }}),
    `2年生` = data |> filter({{ grade }} == '2年') |> select({{ include }}),
    `3年生` = data |> filter({{ grade }} == '3年') |> select({{ include }}),
    `男` = data |> filter({{ gender }} == '男') |> select({{ include }}),
    `女` = data |> filter({{ gender }} == '女') |> select({{ include }})
  )
}

# 各データセットにおける単純集計表の作成
simple_tabulation <- function(data, cont_var) {
  .table <-
    imap(
      data,
      \(x, y)
      # 連続変数は平均値を算出
      if(y %in% cont_var) {
        tibble(
          label = y,
          percent = case_when(
            x == '非該当' ~ NA,
            x == '無回答' ~ NA,
            .default = as.character(x)
          ) |>
            parse_double() |>
            mean(na.rm = TRUE) |>
            number(accuracy = 0.1)
        )
      } else {
        x |>
          tabyl(show_na = FALSE) |>
          as_tibble() |>
          select(label = x, percent) |>
          mutate(
            label = as.character(label),
            percent = percent(percent, accuracy = 0.1)
          )
      }
    ) |>
    list_rbind(names_to = 'variable')
  # 度数を先頭に追加
  bind_rows(
    tibble(variable = '度数', label = '度数', percent = nrow(data) |> number(big.mark = ',')),
    .table
  )
}

