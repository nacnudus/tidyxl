#include <Rcpp.h>
#include "rapidxml.h"
#include "xf.h"

using namespace Rcpp;

xf::xf(rapidxml::xml_node<>* xf) {
  CharacterVector NA_STRING_VECTOR(1, NA_STRING);
  IntegerVector NA_INTEGER_VECTOR(1, NA_INTEGER);

  numFmtId_          =    int_value(xf, "numFmtId");
  fontId_            =    int_value(xf, "fontId");
  fillId_            =    int_value(xf, "fillId");
  borderId_          =    int_value(xf, "borderId");

  applyNumberFormat_ =    int_value(xf, "applyNumberFormat");
  applyFont_         =    int_value(xf, "applyFont");
  applyFill_         =    int_value(xf, "applyFill");
  applyBorder_       =    int_value(xf, "applyBorder");
  applyAlignment_    =    int_value(xf, "applyAlignment");
  applyProtection_   =    int_value(xf, "applyProtection");

  xfId_              =    int_value(xf, "xfId");

  rapidxml::xml_node<>* alignment = xf->first_node("alignment");
  if (alignment == NULL) {
    horizontal_      =    NA_STRING_VECTOR;
    vertical_        =    NA_STRING_VECTOR;
    wrapText_        =    NA_INTEGER_VECTOR;
    readingOrder_    =    NA_INTEGER_VECTOR;
    indent_          =    NA_INTEGER_VECTOR;
    justifyLastLine_ =    NA_INTEGER_VECTOR;
    shrinkToFit_     =    NA_INTEGER_VECTOR;
    textRotation_    =    NA_INTEGER_VECTOR;
  } else {
    horizontal_      = string_value(alignment, "horizontal");
    vertical_        = string_value(alignment, "vertical");
    wrapText_        =    int_value(alignment, "wrapText");
    readingOrder_    =    int_value(alignment, "readingOrder");
    indent_          =    int_value(alignment, "indent");
    justifyLastLine_ =    int_value(alignment, "justifyLastLine");
    shrinkToFit_     =    int_value(alignment, "shrinkToFit");
    textRotation_    =    int_value(alignment, "textRotation");
  }

  rapidxml::xml_node<>* protection = xf->first_node("protection");
  if (protection == NULL) {
    locked_          =    NA_INTEGER;
    hidden_          =    NA_INTEGER;
  } else {
    locked_          =    int_value(protection, "locked");
    hidden_          =    int_value(protection, "hidden");
  }
}

IntegerVector xf::int_value(rapidxml::xml_node<>* node, const char* name) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  IntegerVector out(1);
  if (attribute == NULL) {
    out[0] = NA_INTEGER;
  } else {
    out[0] = strtol(attribute->value(), NULL, 10);
  }
  return(out);
}

CharacterVector xf::string_value(rapidxml::xml_node<>* node, const char* name) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  CharacterVector out(1);
  if (attribute == NULL) {
    out[0] = NA_STRING;
  } else {
    out[0] = attribute->value();
  }
  return(out);
}
