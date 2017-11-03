#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxstyles.h"
#include "gradientStop.h"

using namespace Rcpp;

gradientStop::gradientStop() { // Default constructor
  position_ = NA_REAL;
}

gradientStop::gradientStop(rapidxml::xml_node<>* stop, xlsxstyles* styles) {
  position_ = strtod(stop->first_attribute("position")->value(), NULL);
  color_ = color(stop->first_node("color"), styles);
}
