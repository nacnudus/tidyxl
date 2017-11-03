#include <Rcpp.h>
#include "rapidxml.h"
#include "patternFill.h"
#include "xlsxstyles.h"
#include "color.h"

using namespace Rcpp;

patternFill::patternFill(rapidxml::xml_node<>* patternFill,
    xlsxstyles* styles
    ) {
  patternType_ = NA_STRING;

  if (patternFill != NULL) {
    fgColor_ = color(patternFill->first_node("fgColor"), styles);
    bgColor_ = color(patternFill->first_node("bgColor"), styles);
    std::string patternType = patternFill->first_attribute("patternType")->value();
    if (patternType != "none") {
      patternType_ = patternFill->first_attribute("patternType")->value();
    }
  }
}
