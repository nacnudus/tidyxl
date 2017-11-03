#ifndef GRADIENTFILL_
#define GRADIENTFILL_

class xlsxstyles; // Forward declaration because font is included in xlsxstyles.

#include <Rcpp.h>
#include "rapidxml.h"
#include "color.h"
#include "gradientStop.h"

class gradientFill {

  public:

    Rcpp::String type_;
    int          degree_;
    double       left_;
    double       right_;
    double       top_;
    double       bottom_;
    gradientStop stop1_;
    gradientStop stop2_;

    gradientFill(rapidxml::xml_node<>* gradientFill, xlsxstyles* styles);
};

#endif
