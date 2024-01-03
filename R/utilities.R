# utility functions for the project


pvalue_to_star <- function(.var) {
  case_when(
    {{.var}} < 0.001 ~ '***',
    {{.var}} < 0.01 & {{.var}} >= 0.001 ~ '**',
    {{.var}} < 0.05 & {{.var}} >= 0.01 ~ '*',
    .default = ''
  )
}
