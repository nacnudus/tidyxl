#include <Rcpp.h>
#include "rapidxml.h"
#include "stroke.h"
#include "xlsxstyles.h"
#include "color.h"

using namespace Rcpp;

stroke::stroke(
    rapidxml::xml_node<>* stroke,
    xlsxstyles* styles
    ) {
  style_ = NA_STRING;

  if (stroke != NULL) {
    rapidxml::xml_attribute<>* style = stroke->first_attribute("style");
    if (style != NULL) {
      style_ = style->value();
      color_ = color(stroke->first_node("color"), styles);
    }
  }
}
