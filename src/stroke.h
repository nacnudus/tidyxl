#ifndef STROKE_
#define STROKE_

class xlsxstyles; // Forward declaration because font is included in xlsxstyles.

#include <Rcpp.h>
#include "rapidxml.h"
#include "color.h"

class stroke {

  public:

    Rcpp::String style_;
    color        color_;

    stroke(rapidxml::xml_node<>* stroke, xlsxstyles* styles);
};

#endif
