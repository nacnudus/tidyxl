#ifndef PATTERNFILL_
#define PATTERNFILL_

class styles; // Forward declaration because font is included in styles.

#include <Rcpp.h>
#include "rapidxml.h"
#include "color.h"

class patternFill {

  public:

    color        fgColor_;
    color        bgColor_;
    Rcpp::String patternType_;

    patternFill(); // Default constructor
    patternFill(rapidxml::xml_node<>* patternFill, styles* styles);
};

#endif
