% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/event_analysis.R
\name{analyze_events}
\alias{analyze_events}
\title{Analyze Events}
\usage{
analyze_events(
  event_df,
  merged_df,
  estimation_window = 120,
  event_window = c(-10, 10)
)
}
\arguments{
\item{event_df}{Data frame of event information.}

\item{merged_df}{Merged data frame of firm and market data.}

\item{estimation_window}{Number of days for the estimation window.}

\item{event_window}{Numeric vector indicating the event window (e.g., c(-10, 10)).}
}
\value{
A data frame with CAR and test statistics.
}
\description{
This function computes CAR for multiple events and saves the result as a CSV file.
}
