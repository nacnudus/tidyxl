#ifndef XLSXVALIDATION_
#define XLSXVALIDATION_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxbook.h"

class xlsxvalidation {

  public:

    Rcpp::CharacterVector sheet_;
    Rcpp::CharacterVector ref_;
    Rcpp::CharacterVector type_;
    Rcpp::CharacterVector operator_;
    Rcpp::CharacterVector formula1_;
    Rcpp::CharacterVector formula2_;
    Rcpp::LogicalVector   allow_blank_;
    Rcpp::LogicalVector   show_input_message_;
    Rcpp::CharacterVector prompt_title_;
    Rcpp::CharacterVector prompt_;
    Rcpp::LogicalVector   show_error_message_;
    Rcpp::CharacterVector error_title_;
    Rcpp::CharacterVector error_;
    Rcpp::CharacterVector error_style_;

    xlsxvalidation(
      std::string path,
      Rcpp::CharacterVector sheet_paths,
      Rcpp::CharacterVector sheet_names);

    Rcpp::List information();       // Validation rules DF wrapped in list

};

#endif
