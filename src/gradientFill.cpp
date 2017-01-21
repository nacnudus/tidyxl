#include <Rcpp.h>
#include "rapidxml.h"
#include "gradientFill.h"
#include "styles.h"
#include "color.h"

using namespace Rcpp;

gradientFill::gradientFill() {} // Default constructor

gradientFill::gradientFill(
    rapidxml::xml_node<>* gradientFill,
    styles* styles
    ) {
  // Initialize variables
  type_   = NA_STRING;
  degree_ = IntegerVector::create(NA_INTEGER);
  left_   = NumericVector::create(NA_REAL);
  right_  = NumericVector::create(NA_REAL);
  top_    = NumericVector::create(NA_REAL);
  bottom_ = NumericVector::create(NA_REAL);

  if (gradientFill != NULL) {
    rapidxml::xml_attribute<>* type = gradientFill->first_attribute("type");
    if (type != NULL)
      type_ = type->value();

    rapidxml::xml_attribute<>* degree = gradientFill->first_attribute("degree");
    if (degree != NULL)
      degree_[0] = strtol(degree->value(), NULL, 10);

    rapidxml::xml_attribute<>* left = gradientFill->first_attribute("left");
    if (left != NULL)
      left_[0] = strtof(left->value(), NULL);

    rapidxml::xml_attribute<>* right = gradientFill->first_attribute("right");
    if (right != NULL)
      right_[0] = strtof(right->value(), NULL);

    rapidxml::xml_attribute<>* top = gradientFill->first_attribute("top");
    if (top != NULL)
      top_[0] = strtof(top->value(), NULL);

    rapidxml::xml_attribute<>* bottom = gradientFill->first_attribute("bottom");
    if (bottom != NULL)
      bottom_[0] = strtof(bottom->value(), NULL);

    rapidxml::xml_node<>* stop_node = gradientFill->first_node("stop");
    color1_ = color(stop_node, styles);
    stop_node = stop_node->next_sibling();
    color2_ = color(stop_node, styles);
  }
}
