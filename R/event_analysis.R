#' Load and Prepare Data
#'
#' This function loads event, firm, and market data from Excel files and prepares them for analysis.
#'
#' @param event_file Path to the event data file.
#' @param firm_file Path to the firm data file.
#' @param market_file Path to the market data file.
#' @return A list containing cleaned event and merged data frames.
#' @importFrom dplyr rename_all mutate left_join
#' @importFrom readxl read_excel
#' @importFrom magrittr %>%
#' @export
load_and_prepare_data <- function(event_file, firm_file, market_file) {
  # 自动加载数据，不指定列类型
  event_df <- readxl::read_excel(event_file) %>% dplyr::rename_all(tolower)
  firm_df <- readxl::read_excel(firm_file) %>% dplyr::rename_all(tolower)
  market_df <- readxl::read_excel(market_file) %>% dplyr::rename_all(tolower)

  # 确保 Symbol 列为字符类型
  if ("symbol" %in% colnames(event_df)) {
    event_df$symbol <- as.character(event_df$symbol)
  }
  if ("symbol" %in% colnames(firm_df)) {
    firm_df$symbol <- as.character(firm_df$symbol)
  }
  if ("symbol" %in% colnames(market_df)) {
    market_df$symbol <- as.character(market_df$symbol)
  }

  # 将日期列转换为日期类型
  event_df <- dplyr::mutate(event_df, date = as.Date(date))
  firm_df <- dplyr::mutate(firm_df, date = as.Date(date))
  market_df <- dplyr::mutate(market_df, date = as.Date(date))

  # 合并数据
  merged_df <- dplyr::left_join(firm_df, market_df, by = "date")
  list(event_df = event_df, merged_df = merged_df)
}


#' Calculate CAR and Statistics
#'
#' This function calculates cumulative abnormal returns (CAR) and performs a t-test.
#'
#' @param ar Numeric vector of abnormal returns.
#' @return A list with CAR, t-stat, and p-value.
#' @importFrom stats sd pt
#' @export
calculate_statistics <- function(ar) {
  CAR <- sum(ar, na.rm = TRUE)
  if (length(ar) > 1 && stats::sd(ar, na.rm = TRUE) > 0) {
    t_stat <- CAR / (stats::sd(ar, na.rm = TRUE) / sqrt(length(ar)))
    p_value <- 2 * stats::pt(-abs(t_stat), df = length(ar) - 1)
  } else {
    t_stat <- NA
    p_value <- NA
  }
  list(CAR = CAR, t_stat = t_stat, p_value = round(p_value, 5))
}

#' Non-Parametric Tests
#'
#' Performs non-parametric tests: sign test and Wilcoxon test.
#'
#' @param ar Numeric vector of abnormal returns.
#' @return A list with sign test and Wilcoxon test results.
#' @importFrom stats binom.test wilcox.test
#' @export
non_parametric_tests <- function(ar) {
  num_positive <- sum(ar > 0, na.rm = TRUE)
  total <- length(ar)
  sign_test_p_value <- if (total > 0) stats::binom.test(num_positive, total, p = 0.5)$p.value else NA
  wilcoxon <- tryCatch(stats::wilcox.test(ar, exact = FALSE), error = function(e) return(list(statistic = NA, p.value = NA)))
  list(sign_test_p_value = sign_test_p_value, wilcoxon_stat = wilcoxon$statistic, wilcoxon_p_value = round(wilcoxon$p.value, 5))
}

#' Calculate Event Model
#'
#' Calculates CAR for market model and market-adjusted model.
#'
#' @param firm_data Data frame of firm data.
#' @param event_date Date of the event.
#' @param estimation_window Number of days for the estimation window.
#' @param event_window Numeric vector indicating the event window (e.g., c(-10, 10)).
#' @return A list of CAR and test results.
#' @importFrom dplyr filter
#' @importFrom stats lm predict
#' @importFrom magrittr %>%
#' @export
calculate_event_model <- function(firm_data, event_date, estimation_window = 120, event_window = c(-10, 10)) {
  window_length <- diff(event_window) + 1

  estimation_data <- dplyr::filter(firm_data, date < event_date) %>% utils::tail(estimation_window)

  if (nrow(estimation_data) < 30) return(rep(NA, 12))

  model <- stats::lm(dretnd ~ retindex, data = estimation_data)
  event_window_data <- dplyr::filter(firm_data, date >= (event_date + event_window[1]) & date <= (event_date + event_window[2]))

  if (nrow(event_window_data) < window_length) return(rep(NA, 12))

  ar_market_model <- event_window_data$dretnd - stats::predict(model, event_window_data)
  ar_market_adjusted <- event_window_data$dretnd - event_window_data$retindex

  stats_market_model <- calculate_statistics(ar_market_model)
  stats_market_adjusted <- calculate_statistics(ar_market_adjusted)

  non_param_stats_market <- non_parametric_tests(ar_market_model)
  non_param_stats_adjusted <- non_parametric_tests(ar_market_adjusted)

  c(stats_market_model, non_param_stats_market, stats_market_adjusted, non_param_stats_adjusted)
}

#' Analyze Events
#'
#' This function computes CAR for multiple events and saves the result as a CSV file.
#'
#' @param event_df Data frame of event information.
#' @param merged_df Merged data frame of firm and market data.
#' @param estimation_window Number of days for the estimation window.
#' @param event_window Numeric vector indicating the event window (e.g., c(-10, 10)).
#' @return A data frame with CAR and test statistics.
#' @importFrom purrr map2_dfr
#' @importFrom utils write.csv
#' @export
analyze_events <- function(event_df, merged_df, estimation_window = 120, event_window = c(-10, 10)) {
  results <- purrr::map2_dfr(event_df$symbol, event_df$date, ~ {
    firm_data <- dplyr::filter(merged_df, stkcd == .x)

    if (nrow(firm_data) < estimation_window) {
      return(NULL)
    }

    event_stats <- calculate_event_model(firm_data, .y, estimation_window, event_window)

    if (all(is.na(event_stats))) {
      return(NULL)
    }

    shortname <- event_df$shortname[event_df$symbol == .x]

    c(Symbol = as.character(.x), ShortName = shortname, Event_Date = .y, event_stats)
  })

  utils::write.csv(results, "car_results_per_event.csv", row.names = FALSE)
  results
}


#' Apply Conditional Formatting
#'
#' Applies conditional formatting to an Excel file based on p-values.
#'
#' @param filename Path to the Excel file.
#' @param p_value_column Column letter for p-values to format.
#' @importFrom openxlsx loadWorkbook createStyle conditionalFormatting saveWorkbook
#' @export
apply_conditional_formatting <- function(filename, p_value_column) {
  # 检查文件是否存在
  if (!file.exists(filename)) {
    stop("File does not exist.")
  }

  # 加载 Excel 文件
  wb <- openxlsx::loadWorkbook(filename)

  # 获取工作表（假设是第一个工作表）
  ws <- wb[[1]]

  # 创建不同条件的颜色填充样式
  yellow_fill <- openxlsx::createStyle(fontColour = "#000000", bgFill = "#FFFF00")
  blue_fill <- openxlsx::createStyle(fontColour = "#000000", bgFill = "#ADD8E6")
  green_fill <- openxlsx::createStyle(fontColour = "#000000", bgFill = "#90EE90")

  # 应用条件格式
  openxlsx::conditionalFormatting(wb, sheet = 1, cols = which(names(ws) == p_value_column), rows = 2:(nrow(ws) + 1),
                                  rule = "<0.01", style = yellow_fill)
  openxlsx::conditionalFormatting(wb, sheet = 1, cols = which(names(ws) == p_value_column), rows = 2:(nrow(ws) + 1),
                                  rule = "AND(D2>=0.01,D2<0.05)", style = blue_fill)
  openxlsx::conditionalFormatting(wb, sheet = 1, cols = which(names(ws) == p_value_column), rows = 2:(nrow(ws) + 1),
                                  rule = "AND(D2>=0.05,D2<0.1)", style = green_fill)

  # 保存带有条件格式的 Excel 文件
  openxlsx::saveWorkbook(wb, filename, overwrite = TRUE)
}


#' Calculate Daily Abnormal Returns for Event Window
#'
#' This function calculates daily abnormal returns (AR) for each day in the event window.
#'
#' @param event_df Data frame of event information.
#' @param merged_df Merged data frame of firm and market data.
#' @param estimation_window Number of days for the estimation window.
#' @param event_window Numeric vector indicating the event window (e.g., c(-10, 10)).
#' @return A data frame of daily AR for each event.
#' @importFrom purrr map2_dfr
#' @export
calculate_daily_ar <- function(event_df, merged_df, estimation_window = 120, event_window = c(-10, 10)) {
  window_length <- diff(event_window) + 1

  all_ar_values <- purrr::map2_dfr(event_df$symbol, event_df$date, ~ {
    firm_data <- dplyr::filter(merged_df, stkcd == .x)

    estimation_data <- dplyr::filter(firm_data, date < .y) %>% utils::tail(estimation_window)
    if (nrow(estimation_data) < estimation_window) {
      return(data.frame(matrix(NA, nrow = 1, ncol = window_length)))
    }

    model <- stats::lm(dretnd ~ retindex, data = estimation_data)
    event_window_data <- dplyr::filter(firm_data, date >= (.y + event_window[1]) & date <= (.y + event_window[2]))

    if (nrow(event_window_data) < window_length) {
      return(data.frame(matrix(NA, nrow = 1, ncol = window_length)))
    }

    ar_market_model <- event_window_data$dretnd - stats::predict(model, event_window_data)
    ar_values <- rep(NA, window_length)
    ar_values[1:length(ar_market_model)] <- ar_market_model

    data.frame(t(ar_values))
  })

  colnames(all_ar_values) <- paste0("Day_", event_window[1]:event_window[2])
  all_ar_values
}

#' Perform Daily Mean CAR T-test
#'
#' This function performs a t-test on the mean CAR for each day in the event window.
#'
#' @param ar_df Data frame of daily AR for each event.
#' @param event_window Numeric vector indicating the event window (e.g., c(-10, 10)).
#' @return A data frame with mean CAR, t-statistics, and p-values for each day.
#' @importFrom stats t.test
#' @importFrom utils write.csv
#' @export
daily_mean_t_test <- function(ar_df, event_window = c(-10, 10)) {
  days <- event_window[1]:event_window[2]

  results <- purrr::map_dfr(days, function(day) {
    ar_day <- ar_df[[paste0("Day_", day)]]

    if (sum(!is.na(ar_day)) < 2) {
      return(data.frame(Day = day, Mean_CAR = NA, t_stat = NA, p_value = NA))
    }

    t_test <- stats::t.test(ar_day, mu = 0, na.rm = TRUE)

    data.frame(
      Day = day,
      Mean_CAR = mean(ar_day, na.rm = TRUE),
      t_stat = t_test$statistic,
      p_value = t_test$p.value
    )
  })

  utils::write.csv(results, "average_car_and_t_test_results.csv", row.names = FALSE)
  results
}
