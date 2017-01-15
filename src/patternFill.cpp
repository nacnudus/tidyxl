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
  // Initialize variable
  patternType_ = CharacterVector::create(NA_STRING);

  if (patternFill != NULL) {
    fgColor_ = color(patternFill->first_node("fgColor"), styles);
    bgColor_ = color(patternFill->first_node("bgColor"), styles);
    patternType_[0] = patternFill->first_attribute("patternType")->value();
  }
}
