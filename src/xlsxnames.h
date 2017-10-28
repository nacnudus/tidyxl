#ifndef XLSXNAMES_
#define XLSXNAMES_

#include <Rcpp.h>
#include "rapidxml.h"

class xlsxnames {

  public:

    Rcpp::List information_;         // Wrapper for variables returned to R

    Rcpp::CharacterVector name_;
    Rcpp::IntegerVector   sheet_id_;
    Rcpp::CharacterVector formula_;
    Rcpp::CharacterVector comment_;
    Rcpp::LogicalVector   hidden_;

    xlsxnames(const std::string& path);

    Rcpp::List& information();       // Validation rules DF wrapped in list

};

#endif

