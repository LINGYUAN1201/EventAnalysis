# test-event_analysis.R

library(testthat)
library(EventAnalysis)

# 创建测试数据
event_df <- data.frame(
  symbol = "A",
  shortname = "Test Firm",
  date = as.Date("2023-01-01")
)

firm_data <- data.frame(
  stkcd = "A",
  date = seq(as.Date("2022-07-01"), as.Date("2023-01-10"), by = "day"),
  dretnd = rnorm(194),
  retindex = rnorm(194)
)

market_data <- data.frame(
  date = seq(as.Date("2022-07-01"), as.Date("2023-01-10"), by = "day"),
  retindex = rnorm(194)
)

# 合并 firm_data 和 market_data 模拟 merged_df
merged_df <- merge(firm_data, market_data, by = "date", suffixes = c("", "_market"))

test_that("Data loading works with simulated data", {
  expect_true(is.data.frame(event_df))
  expect_true(is.data.frame(merged_df))
})

test_that("Event analysis produces expected output format", {
  results <- analyze_events(event_df, merged_df, estimation_window = 150, event_window = c(-5, 5))
  expect_true(is.data.frame(results))
  expect_true(all(c("Symbol", "Event_Date") %in% names(results)))
})

test_that("Daily AR calculation works", {
  ar_df <- calculate_daily_ar(event_df, merged_df, estimation_window = 120, event_window = c(-5, 5))
  expect_true(is.data.frame(ar_df))
})

test_that("Daily mean CAR t-test works", {
  ar_df <- calculate_daily_ar(event_df, merged_df, estimation_window = 120, event_window = c(-5, 5))
  daily_mean_results <- daily_mean_t_test(ar_df, event_window = c(-5, 5))
  expect_true(is.data.frame(daily_mean_results))
  expect_true(all(c("Day", "Mean_CAR", "t_stat", "p_value") %in% names(daily_mean_results)))
})
