#include <Rcpp.h>
#include "rapidxml.h"
#include "xf.h"

using namespace Rcpp;

xf::xf() {} // Default constructor

xf::xf(rapidxml::xml_node<>* xf):
  readingOrderChr_{"context", "left-to-right", "right-to-left"}
{
  numFmtId_          = int_value(xf, "numFmtId", NA_INTEGER);
  fontId_            = int_value(xf, "fontId", NA_INTEGER);
  fillId_            = int_value(xf, "fillId", NA_INTEGER);
  borderId_          = int_value(xf, "borderId", NA_INTEGER);

  applyNumberFormat_ = bool_value(xf, "applyNumberFormat", true);
  applyFont_         = bool_value(xf, "applyFont", true);
  applyFill_         = bool_value(xf, "applyFill", true);
  applyBorder_       = bool_value(xf, "applyBorder", true);
  applyAlignment_    = bool_value(xf, "applyAlignment", true);
  applyProtection_   = bool_value(xf, "applyProtection", true);

  xfId_ = int_value(xf, "xfId", NA_INTEGER);
  if (xfId_ == NA_INTEGER) xfId_ = 0;

  rapidxml::xml_node<>* alignment = xf->first_node("alignment");
  if (alignment == NULL) {
    horizontal_      = "general";
    vertical_        = "bottom";
    wrapText_        = false;
    readingOrder_    = "context";
    indent_          = 0;
    justifyLastLine_ = false;
    shrinkToFit_     = false;
    textRotation_    = 0;
  } else {
    horizontal_      = string_value(alignment, "horizontal", "general");
    vertical_        = string_value(alignment, "vertical", "bottom");
    wrapText_        = bool_value(alignment, "wrapText", false);
    readingOrder_    = readingOrder(alignment);
    indent_          = int_value(alignment, "indent", 0);
    justifyLastLine_ = bool_value(alignment, "justifyLastLine", false);
    shrinkToFit_     = bool_value(alignment, "shrinkToFit", false);
    textRotation_    = int_value(alignment, "textRotation", 0);
  }

  rapidxml::xml_node<>* protection = xf->first_node("protection");
  if (protection == NULL) {
    locked_          = NA_LOGICAL;
    hidden_          = NA_LOGICAL;
  } else {
    locked_          = bool_value(protection, "locked", true);
    hidden_          = bool_value(protection, "hidden", true);
  }
}

int xf::bool_value(rapidxml::xml_node<>* node, const char* name, int _default) {
  std::string value;
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute == NULL) {
    return(true);
  }
  value = attribute->value();
  if (value == "0" || value == "false") {
    return(false);
  } else {
    return(true);
  }
  return(_default);
}

int xf::int_value(rapidxml::xml_node<>* node, const char* name, int _default) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    return(strtol(attribute->value(), NULL, 10));
  }
  return(_default);
}

String xf::string_value(rapidxml::xml_node<>* node, const char* name,
    Rcpp::String _default) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    return(attribute->value());
  }
  return(_default);
}

String xf::readingOrder(rapidxml::xml_node<>* node) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute("readingOrder");
  if (attribute != NULL) {
    return(readingOrderChr_[strtol(attribute->value(), NULL, 10)]);
  }
  return("context");
}
