#include <Rcpp.h>
#include "rapidxml.h"
#include "xf.h"

using namespace Rcpp;

xf::xf() {} // Default constructor

xf::xf(rapidxml::xml_node<>* xf):
  readingOrderChr_{"context", "left-to-right", "right-to-left"}
{
  numFmtId_          = int_value(xf, "numFmtId");
  fontId_            = int_value(xf, "fontId");
  fillId_            = int_value(xf, "fillId");
  borderId_          = int_value(xf, "borderId");

  applyNumberFormat_ = bool_value(xf, "applyNumberFormat");
  applyFont_         = bool_value(xf, "applyFont");
  applyFill_         = bool_value(xf, "applyFill");
  applyBorder_       = bool_value(xf, "applyBorder");
  applyAlignment_    = bool_value(xf, "applyAlignment");
  applyProtection_   = bool_value(xf, "applyProtection");

  xfId_ = int_value(xf, "xfId");
  if (xfId_ == NA_INTEGER) xfId_ = 0;

  rapidxml::xml_node<>* alignment = xf->first_node("alignment");
  if (alignment == NULL) {
    horizontal_      = NA_STRING;
    vertical_        = NA_STRING;
    wrapText_        = NA_LOGICAL;
    readingOrder_    = NA_STRING;
    indent_          = NA_INTEGER;
    justifyLastLine_ = NA_LOGICAL;
    shrinkToFit_     = NA_LOGICAL;
    textRotation_    = NA_INTEGER;
  } else {
    horizontal_      = string_value(alignment, "horizontal");
    vertical_        = string_value(alignment, "vertical");
    wrapText_        = bool_value(alignment, "wrapText");
    readingOrder_    = readingOrder(alignment);
    indent_          = int_value(alignment, "indent");
    justifyLastLine_ = bool_value(alignment, "justifyLastLine");
    shrinkToFit_     = bool_value(alignment, "shrinkToFit");
    textRotation_    = int_value(alignment, "textRotation");
  }

  rapidxml::xml_node<>* protection = xf->first_node("protection");
  if (protection == NULL) {
    locked_          = NA_LOGICAL;
    hidden_          = NA_LOGICAL;
  } else {
    locked_          = bool_value(protection, "locked");
    hidden_          = bool_value(protection, "hidden");
  }
}

int xf::bool_value(rapidxml::xml_node<>* node, const char* name) {
  std::string value;
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute == NULL) {
    return(true);
  }
  value = attribute->value();
  if (value == "0" || value == "false") {
    return(false);
  }
  return(true);
}

int xf::int_value(rapidxml::xml_node<>* node, const char* name) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    return(strtol(attribute->value(), NULL, 10));
  }
  return(NA_INTEGER);
}

String xf::string_value(rapidxml::xml_node<>* node, const char* name) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    return(attribute->value());
  }
  return(NA_STRING);
}

String xf::readingOrder(rapidxml::xml_node<>* node) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute("readingOrder");
  if (attribute != NULL) {
    return(readingOrderChr_[strtol(attribute->value(), NULL, 10)]);
  }
  return(NA_STRING);
}
