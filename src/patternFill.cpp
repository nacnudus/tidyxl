#include <Rcpp.h>
#include "rapidxml.h"
#include "patternFill.h"
#include "styles.h"
#include "color.h"

using namespace Rcpp;

patternFill::patternFill() {} // Default constructor

patternFill::patternFill(rapidxml::xml_node<>* patternFill,
    styles* styles
    ) {
  patternType_ = NA_STRING;

  if (patternFill != NULL) {
    fgColor_ = color(patternFill->first_node("fgColor"), styles);
    bgColor_ = color(patternFill->first_node("bgColor"), styles);
    patternType_ = patternFill->first_attribute("patternType")->value();
  }
}
