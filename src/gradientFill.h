#ifndef GRADIENTFILL_
#define GRADIENTFILL_

class xlsxstyles; // Forward declaration because font is included in xlsxstyles.

#include <Rcpp.h>
#include "rapidxml.h"
#include "color.h"

class gradientFill {

  public:

    Rcpp::String        type_;
    Rcpp::IntegerVector degree_;
    Rcpp::NumericVector left_;
    Rcpp::NumericVector right_;
    Rcpp::NumericVector top_;
    Rcpp::NumericVector bottom_;

    color               color1_;
    color               color2_;

    gradientFill(); // Default constructor
    gradientFill(rapidxml::xml_node<>* gradientFill, xlsxstyles* styles);
};

#endif
