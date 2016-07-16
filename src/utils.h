#ifndef tidyxl_utils_
#define tidyxl_utils_

#include <Rcpp.h>

inline void makeDataFrame(Rcpp::List& Lists, const Rcpp::CharacterVector& colnames) {
  // turn List of Vectors into a data frame without checking anything
  int n = Rf_length(Lists[0]);
  Lists.attr("class") = Rcpp::CharacterVector::create("tbl_df", "tbl", "data.frame");
  Lists.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, -n); // Dunno how this works
  Lists.attr("names") = colnames;
}

#endif
