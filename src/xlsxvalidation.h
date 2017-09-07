#ifndef XLSXVALIDATION_
#define XLSXVALIDATION_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxbook.h"

class xlsxvalidation {

  public:

    xlsxbook& book_; // To determine the date system (1900/1904)

    Rcpp::List information_; // Wrapper for variables returned to R

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
        const std::string& sheet_path,
        xlsxbook& book);

    Rcpp::List& information();       // Validation rules DF wrapped in list

};

#endif
