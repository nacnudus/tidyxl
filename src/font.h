#ifndef FONT_
#define FONT_

class xlsxstyles; // Forward declaration because font is included in xlsxstyles.

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
    Rcpp::NumericVector size_;      // or sz for googlesheets
    color               color_;
    Rcpp::String        name_;
    Rcpp::IntegerVector family_;
    Rcpp::String        scheme_;

    font(); // default constructor
    font(rapidxml::xml_node<>* font, xlsxstyles* styles);
};

#endif
