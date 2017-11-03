#ifndef PATTERNFILL_
#define PATTERNFILL_

class xlsxstyles; // Forward declaration because font is included in xlsxstyles.

#include <Rcpp.h>
#include "rapidxml.h"
#include "color.h"

class patternFill {

  public:

    color        fgColor_;
    color        bgColor_;
    Rcpp::String patternType_;

    patternFill(rapidxml::xml_node<>* patternFill, xlsxstyles* styles);
};

#endif
