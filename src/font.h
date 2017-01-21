#ifndef FONT_
#define FONT_

class styles; // Forward declaration because font is included in styles.

#include <Rcpp.h>
#include "rapidxml.h"
#include "color.h"

class font {

  public:

    Rcpp::LogicalVector b_;         // bold
    Rcpp::LogicalVector i_;         // italic
    Rcpp::String        u_;         // underline (val attribute e.g. "none" or no attribute at all)
    Rcpp::LogicalVector strike_;    // strikethrough
    Rcpp::String        vertAlign_; // (val attribute)
    Rcpp::IntegerVector size_;      // or sz for googlesheets
    color               color_;
    Rcpp::String        name_;
    Rcpp::IntegerVector family_;
    Rcpp::String        scheme_;

    font(); // default constructor
    font(rapidxml::xml_node<>* font, styles* styles);
};

#endif
