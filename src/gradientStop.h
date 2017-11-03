#ifndef GRADIENTSTOP_
#define GRADIENTSTOP_

class xlsxstyles;

#include <Rcpp.h>
#include "rapidxml.h"
#include "color.h"

class gradientStop {

  public:

    double position_;
    color  color_;

    gradientStop(); // Default constructor
    gradientStop(rapidxml::xml_node<>* stop, xlsxstyles* styles);
};

#endif
