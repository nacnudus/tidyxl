#include <Rcpp.h>
#include "rapidxml.h"
#include "border.h"
#include "styles.h"
#include "stroke.h"

using namespace Rcpp;

border::border() {}; // default constructor

border::border(bool blank) { // Construct empty vectors
}

border::border(rapidxml::xml_node<>* border,
    styles* styles
    ) {
  // Initialize variables
  diagonalDown_ = IntegerVector::create(NA_INTEGER);
  diagonalUp_   = IntegerVector::create(NA_INTEGER);
  outline_      = IntegerVector::create(NA_INTEGER);

  rapidxml::xml_attribute<>* diagonalDown = border->first_attribute("diagonalDown");
  if (diagonalDown != NULL)
    diagonalDown_[0] = strtol(diagonalDown->value(), NULL, 10);

  rapidxml::xml_attribute<>* diagonalUp = border->first_attribute("diagonalUp");
  if (diagonalUp != NULL)
    diagonalUp_[0] = strtol(diagonalUp->value(), NULL, 10);

  rapidxml::xml_attribute<>* outline = border->first_attribute("outline");
  if (outline != NULL)
    outline_[0] = strtol(outline->value(), NULL, 10);

  left_ = stroke(border->first_node("left"), styles);
  right_ = stroke(border->first_node("right"), styles);
  start_ = stroke(border->first_node("start"), styles); // gnumeric
  end_ = stroke(border->first_node("end"), styles);     // gnumeric
  top_ = stroke(border->first_node("top"), styles);
  bottom_ = stroke(border->first_node("bottom"), styles);
  diagonal_ = stroke(border->first_node("diagonal"), styles);
  vertical_ = stroke(border->first_node("vertical"), styles);
  horizontal_ = stroke(border->first_node("horizontal"), styles);
}
