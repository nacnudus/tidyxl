#include <Rcpp.h>
#include "rapidxml.h"
#include "xf.h"

using namespace Rcpp;

xf::xf() {}; // Default constructor

xf::xf(rapidxml::xml_node<>* xf) {
  numFmtId_          =    int_value(xf, "numFmtId");
  fontId_            =    int_value(xf, "fontId");
  fillId_            =    int_value(xf, "fillId");
  borderId_          =    int_value(xf, "borderId");

  applyNumberFormat_ =    bool_value(xf, "applyNumberFormat");
  applyFont_         =    bool_value(xf, "applyFont");
  applyFill_         =    bool_value(xf, "applyFill");
  applyBorder_       =    bool_value(xf, "applyBorder");
  applyAlignment_    =    bool_value(xf, "applyAlignment");
  applyProtection_   =    bool_value(xf, "applyProtection");

  xfId_              =    int_value(xf, "xfId");
  if (xfId_[0] == NA_INTEGER) xfId_[0] = 0;

  rapidxml::xml_node<>* alignment = xf->first_node("alignment");
  if (alignment == NULL) {
    horizontal_      =    CharacterVector::create(NA_STRING);
    vertical_        =    CharacterVector::create(NA_STRING);
    wrapText_        =    IntegerVector::create(NA_INTEGER);
    readingOrder_    =    CharacterVector::create(NA_STRING);
    indent_          =    IntegerVector::create(NA_INTEGER);
    justifyLastLine_ =    IntegerVector::create(NA_INTEGER);
    shrinkToFit_     =    IntegerVector::create(NA_INTEGER);
    textRotation_    =    IntegerVector::create(NA_INTEGER);
  } else {
    horizontal_      =    string_value(alignment, "horizontal");
    vertical_        =    string_value(alignment, "vertical");
    wrapText_        =    bool_value(alignment, "wrapText");
    readingOrder_    =    readingOrder(alignment);
    indent_          =    int_value(alignment, "indent");
    justifyLastLine_ =    bool_value(alignment, "justifyLastLine");
    shrinkToFit_     =    bool_value(alignment, "shrinkToFit");
    textRotation_    =    int_value(alignment, "textRotation");
  }

  rapidxml::xml_node<>* protection = xf->first_node("protection");
  if (protection == NULL) {
    locked_          =    NA_LOGICAL;
    hidden_          =    NA_LOGICAL;
  } else {
    locked_          =    bool_value(protection, "locked");
    hidden_          =    bool_value(protection, "hidden");
  }
}

LogicalVector xf::bool_value(rapidxml::xml_node<>* node, const char* name) {
  LogicalVector out(1, NA_LOGICAL);
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    out[0] = strtol(attribute->value(), NULL, 10);
  }
  return(out);
}

IntegerVector xf::int_value(rapidxml::xml_node<>* node, const char* name) {
  IntegerVector out(1, NA_INTEGER);
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    out[0] = strtol(attribute->value(), NULL, 10);
  }
  return(out);
}

CharacterVector xf::string_value(rapidxml::xml_node<>* node, const char* name) {
  CharacterVector out(1, NA_STRING);
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    out[0] = attribute->value();
  }
  return(out);
}

CharacterVector xf::readingOrder(rapidxml::xml_node<>* node) {
  CharacterVector out(1, NA_STRING);
  rapidxml::xml_attribute<>* attribute = node->first_attribute("readingOrder");
  if (attribute != NULL) {
    out[0] = readingOrderChr_[strtol(attribute->value(), NULL, 10)];
  }
  return(out);
}

