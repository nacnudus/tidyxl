#ifndef GRADIENTFILL_
#define GRADIENTFILL_

class styles; // Forward declaration because font is included in styles.

#include <Rcpp.h>
#include "rapidxml.h"
#include "color.h"

class gradientFill {

  public:

    Rcpp::CharacterVector type_;
    Rcpp::IntegerVector   degree_;
    Rcpp::NumericVector   left_;
    Rcpp::NumericVector   right_;
    Rcpp::NumericVector   top_;
    Rcpp::NumericVector   bottom_;

    color                 color1_;
    color                 color2_;

    gradientFill(); // Default constructor
    gradientFill(rapidxml::xml_node<>* gradientFill, styles* styles);
};

#endif
