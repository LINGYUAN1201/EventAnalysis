# EventAnalysis

`EventAnalysis` is an R package designed for event study analysis, helping researchers and analysts to assess the impact of events on stock returns, financial indicators, and other metrics. This package is especially useful for financial market research, allowing users to flexibly set event windows and benchmarks.

## Installation

```R
#Install from GitHub
install.packages("devtools")
devtools::install_github("LINGYUAN1201/EventAnalysis")
```

## Features
The EventAnalysis package offers a range of features to make event studies straightforward and efficient, including:

*   Event Window Setting: Supports customization of the pre- and post-event windows to evaluate event impact flexibly.
*   Return Calculation: Calculates cumulative returns, abnormal returns, and cumulative abnormal returns (CAR).
*   Statistical Testing: Includes built-in statistical tests to assess the significance of event impacts.
*   Industry Comparison: Supports comparative analysis between high- and low-emission industries.

## Quick Start
The following example demonstrates how to conduct a basic event study using the EventAnalysis package:

### Load the Package

```R
library(EventAnalysis)
```

### Data Preparation

Currently this package can run the CAR calculated by the market model and the market adjustment model, so you need to prepare three data files:

*   Event.xlsx: This file mainly contains the date of the Event for each company

*   Firm.xlsx: This file mainly includes the closing prices and earnings of each company over a period of time

*   Market.xlsx: This file mainly contains Market index data

### Load data
```R
data <- load_and_prepare_data("Event.xlsx", "Firm.xlsx", "Market.xlsx")
event_df <- data$event_df
merged_df <- data$merged_df
```
### Compute the CAR of multiple events and generate a CSV file
```R
car_results <- analyze_events(event_df, merged_df, estimation_window = 120, event_window = c(-10, 10))
```
### Calculate the daily AR of the event window
```R
ar_df <- calculate_daily_ar(event_df, merged_df, estimation_window = 120, event_window = c(-10, 10))
```
### Calculate daily average CAR and perform t test
```R
daily_mean_results <- daily_mean_t_test(ar_df, event_window = c(-10, 10))
```

## Others
This R package is based on Python code for the same functionality, designed with the ease of use of R in mind. If you are interested in this content, you can contact me directly.

My email: LingYUAN1201@outlook.com
