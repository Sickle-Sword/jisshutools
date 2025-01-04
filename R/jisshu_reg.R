#' Regression Table for Chosa Jisshu
#'
#' 調査実習用の回帰分析結果を作成する関数です。
#'
#' `lm`や`glm`、`nnet::multinom`の結果から、執筆要項に沿ったエクセルの表が作成できます。
#'
#' エクセルへの保存はexampleを参照してください。
#'
#' @param object an object for which a summary is desired. For example, a model object, such as those returned by `lm`, `glm`, and `multinom`.
#'
#' @examples
#' data <- tibble::tibble(
#'   x1 = rnorm(100, mean = 0, sd = 1),
#'   x2 = rnorm(100, mean = 0, sd = 1),
#'   y = 0.2 + 0.3*x1 + 0.5*x2 + rnorm(100, mean = 0, sd = 0.1),
#'   y_bin = rbinom(100, 1, plogis(y)),
#'   y_multinom = cut(
#'     y,
#'     breaks = quantile(y, probs = c(0, 0.25, 0.75, 1)),
#'     labels = c('Q1', 'Q2', 'Q3'),
#'     include.lowest = TRUE
#'   )
#' )
#'
#' data
#'
#' # Create a regression table
#' # Linear regression
#' model_lm <- lm(y ~ x1 + x2, data = data)
#' result_lm <- jisshu_reg(model_lm)
#'
#' # Logistic regression
#' model_glm <- glm(y_bin ~ x1 + x2, data = data, family = binomial(link = 'logit'))
#' result_glm <- jisshu_reg(model_glm)
#'
#' # Multinomial logistic regression
#' model_multinom <- nnet::multinom(y_multinom ~ x1 + x2, data = data, trace = FALSE)
#' result_multinom <- jisshu_reg(model_multinom)
#'
#' # Glance the regression table
#' \dontrun{
#'   result_lm$open()
#'   result_glm$open()
#'   result_multinom$open()
#' }
#'
#' # Save the regression table as an excel file
#' \dontrun{
#'   result_lm$save('hoge_lm.xlsx')
#'   result_glm$save('hoge_glm.xlsx')
#'   result_multinom$save('hoge_multinom.xlsx')
#' }
#'
#' @export
#'

jisshu_reg <- function(object) {
  UseMethod('jisshu_reg')
}

