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
#'   y = 0.2 + 0.3*x1 + 0.5*x2 + rnorm(100, mean = 0, sd = 0.1),
#'   y_bin = rbinom(100, 1, plogis(y))
#' )
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
#' # Save the regression table as an excel file
#' \dontrun{
#' result_lm$save('hoge_lm.xlsx')
#' result_glm$save('hoge_glm.xlsx')
#' }
#'
#' @export
#'

jisshu_reg <- function(object) {
  UseMethod('jisshu_reg')
}

