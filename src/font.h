#ifndef FONT_
#define FONT_

class xlsxstyles; // Forward declaration because font is included in xlsxstyles.

#include <Rcpp.h>
#include "rapidxml.h"
#include "color.h"

class font {

  public:

    int          b_;         // bold
    int          i_;         // italic
    Rcpp::String u_;         // underline (val attribute e.g. "none" or no attribute at all)
    int          strike_;    // strikethrough
    Rcpp::String vertAlign_; // (val attribute)
    double       size_;      // or sz for googlesheets
    color        color_;
    Rcpp::String name_;
    int          family_;
    Rcpp::String scheme_;

    font(rapidxml::xml_node<>* font, xlsxstyles* styles);
};

#endif
