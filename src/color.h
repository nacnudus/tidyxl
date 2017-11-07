#ifndef COLOR_
#define COLOR_

class xlsxstyles; // Forward declaration because color is included in font, which is
              // included in xlsxstyles.

#include <Rcpp.h>
#include "rapidxml.h"

class color {

  public:

    Rcpp::String rgb_;
    Rcpp::String theme_;
    int          indexed_;
    double       tint_;

    color(); // default constructor
    color(rapidxml::xml_node<>* color, xlsxstyles* styles);
};

#endif
