#include <Rcpp.h>
#include "rapidxml.h"
#include "gradientFill.h"
#include "xlsxstyles.h"
#include "color.h"

using namespace Rcpp;

gradientFill::gradientFill() {} // Default constructor

gradientFill::gradientFill(
    rapidxml::xml_node<>* gradientFill,
    xlsxstyles* styles
    ) {
  // Initialize variables
  type_   = NA_STRING;
  degree_ = NA_INTEGER;
  left_   = NA_REAL;
  right_  = NA_REAL;
  top_    = NA_REAL;
  bottom_ = NA_REAL;

  if (gradientFill != NULL) {
    rapidxml::xml_attribute<>* type = gradientFill->first_attribute("type");
    if (type != NULL) {
      type_ = type->value();
      degree_ = NA_INTEGER;
      rapidxml::xml_attribute<>* left = gradientFill->first_attribute("left");
      rapidxml::xml_attribute<>* right = gradientFill->first_attribute("right");
      rapidxml::xml_attribute<>* top = gradientFill->first_attribute("top");
      rapidxml::xml_attribute<>* bottom = gradientFill->first_attribute("bottom");
      left_ = (left != NULL) ? strtod(left->value(), NULL) : 0;
      right_ = (right != NULL) ? strtod(right->value(), NULL) : 0;
      top_ = (top != NULL) ? strtod(top->value(), NULL) : 0;
      bottom_ = (bottom != NULL) ? strtod(bottom->value(), NULL) : 0;
    } else {
      type_ = NA_STRING;
      left_   = NA_REAL;
      right_  = NA_REAL;
      top_    = NA_REAL;
      bottom_ = NA_REAL;
      rapidxml::xml_attribute<>* degree = gradientFill->first_attribute("degree");
      if (degree != NULL) { degree_ = strtol(degree->value(), NULL, 10); } else { degree_ = 0; }
    }

    rapidxml::xml_node<>* stop1 = gradientFill->first_node("stop");
    color color1(stop1->first_node("color"), styles);
    std::string pos1 = stop1->first_attribute("position")->value();

    rapidxml::xml_node<>* stop2 = stop1->next_sibling();
    color color2(stop2->first_node("color"), styles);
    std::string pos2 = stop2->first_attribute("position")->value();

    if (pos1 == "0") {
        color0_ = color1;
    } else if (pos1 == "0.5") {
        color05_ = color1;
    } else if (pos1 == "1") {
        color1_ = color1;
    }

    if (pos2 == "0") {
        color0_ = color2;
    } else if (pos2 == "0.5") {
        color05_ = color2;
    } else if (pos2 == "1") {
        color1_ = color2;
    }
  }
}
