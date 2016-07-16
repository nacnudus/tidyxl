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

// Based on hadley/readxl
// Simple parser: does not check that order of numbers and letters is correct
// row_ and column_ are one-based
inline void parseAddress(std::string& address, int& row, int& col) {
  col = 0;
  row = 0;

  // Iterate though std::string character by character
  for(std::string::const_iterator iter = address.begin();
      iter != address.end(); ++iter) {
    if (*iter >= '0' && *iter <= '9') { // If it's a number
      row = row * 10 + (*iter - '0'); // Then multiply existing row by 10 and add new number
    } else if (*iter >= 'A' && *iter <= 'Z') { // If it's a character
      col = 26 * col + (*iter - 'A' + 1); // Then do similarly  with columns
    } else {
      Rcpp::stop("Invalid character '%s' in cell ref '%s'", *iter, address);
    }
  }
}

inline void getChildValueString(Rcpp::String& container,
    const char* childname, rapidxml::xml_node<>* parent) {
  rapidxml::xml_node<>* child = parent->first_node(childname);
  if (child == NULL) {
    container = NA_STRING;
  } else {
    container = child->value();
  }
}

inline void getAttributeValueString(Rcpp::String& container,
    const char* attributename, rapidxml::xml_node<>* parent) {
  rapidxml::xml_attribute<>* attribute = parent->first_attribute(attributename);
  if (attribute == NULL) {
    container = NA_STRING;
  } else {
    container = attribute->value();
  }
}

#endif
