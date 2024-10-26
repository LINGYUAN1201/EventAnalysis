% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/event_analysis.R
\name{daily_mean_t_test}
\alias{daily_mean_t_test}
\title{Perform Daily Mean CAR T-test}
\usage{
daily_mean_t_test(ar_df, event_window = c(-10, 10))
}
\arguments{
\item{ar_df}{Data frame of daily AR for each event.}

\item{event_window}{Numeric vector indicating the event window (e.g., c(-10, 10)).}
}
\value{
A data frame with mean CAR, t-statistics, and p-values for each day.
}
\description{
This function performs a t-test on the mean CAR for each day in the event window.
}
