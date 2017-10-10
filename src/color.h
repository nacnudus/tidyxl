#ifndef COLOR_
#define COLOR_

class xlsxstyles; // Forward declaration because color is included in font, which is
              // included in xlsxstyles.

#include <Rcpp.h>
#include "rapidxml.h"

class color {

  public:

    Rcpp::String        rgb_;
    Rcpp::IntegerVector theme_;
    Rcpp::IntegerVector indexed_;
    Rcpp::NumericVector tint_;

    color(); // default constructor
    color(rapidxml::xml_node<>* color, xlsxstyles* styles);
};

#endif
