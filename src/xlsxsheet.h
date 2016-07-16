#ifndef XLSXSHEET_
#define XLSXSHEET_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxbook.h"

class xlsxsheet {

std::string name_;
rapidxml::xml_document<> xml_;
rapidxml::xml_node<>* worksheet_;
rapidxml::xml_node<>* sheetData_;
unsigned long long int cellcount_;

// The remaining variables go to R
Rcpp::CharacterVector address_;   // Value of node r
Rcpp::IntegerVector   row_;       // Parsed address_ (one-based)
Rcpp::IntegerVector   col_;       // Parsed address_ (one-based)
Rcpp::CharacterVector content_;   // Unparsed value of node v

Rcpp::List  value_;               // Parsed values wrapped in unnamed lists
Rcpp::CharacterVector type_;      // Type of the parsed value

Rcpp::LogicalVector   logical_;   // Parsed value
Rcpp::NumericVector   numeric_;   // Parsed value
Rcpp::NumericVector   date_;      // Parsed value
Rcpp::CharacterVector character_; // Parsed value
Rcpp::CharacterVector error_;     // Parsed value

// The following are always used.
Rcpp::NumericVector   height_;      // Provided to cell constructor
Rcpp::NumericVector   width_;       // Provided to cell constructor

public:

  xlsxsheet(const std::string& bookPath,
      const int& i, xlsxbook& book);
  Rcpp::List information();       // Cells contents and styles DF wrapped in list

private:

  void cacheCellcount();
  void initializeColumns();
  void parseSheetData();

};

#endif
