#include <Rcpp.h>
#include "rapidxml.h"
#include "stroke.h"
#include "styles.h"
#include "color.h"

using namespace Rcpp;

stroke::stroke() { // Default constructor
  style_ = CharacterVector::create(NA_STRING);
}

stroke::stroke(bool blank) { // Construct empty vectors
}

stroke::stroke(
    rapidxml::xml_node<>* stroke,
    styles* styles
    ) {
  // Initialize variables
  style_ = CharacterVector::create(NA_STRING);

  if (stroke != NULL) {
    rapidxml::xml_attribute<>* style = stroke->first_attribute("style");
    if (style != NULL) {
      style_[0] = style->value();
      color_ = color(stroke->first_node("color"), styles);
    }
  }
}
