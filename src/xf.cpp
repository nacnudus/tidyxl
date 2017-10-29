#include <Rcpp.h>
#include "rapidxml.h"
#include "xf.h"
#include "xml_value.h"

using namespace Rcpp;

xf::xf() {} // Default constructor

xf::xf(rapidxml::xml_node<>* xf):
  readingOrderChr_{"context", "left-to-right", "right-to-left"}
{
  numFmtId_          = int_attr(xf, "numFmtId");
  fontId_            = int_attr(xf, "fontId");
  fillId_            = int_attr(xf, "fillId");
  borderId_          = int_attr(xf, "borderId");

  applyNumberFormat_ = bool_attr(xf, "applyNumberFormat");
  applyFont_         = bool_attr(xf, "applyFont");
  applyFill_         = bool_attr(xf, "applyFill");
  applyBorder_       = bool_attr(xf, "applyBorder");
  applyAlignment_    = bool_attr(xf, "applyAlignment");
  applyProtection_   = bool_attr(xf, "applyProtection");

  xfId_ = int_attr(xf, "xfId");
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
    horizontal_      = string_attr(alignment, "horizontal");
    vertical_        = string_attr(alignment, "vertical");
    wrapText_        = bool_attr(alignment, "wrapText");
    readingOrder_    = readingOrder(alignment);
    indent_          = int_attr(alignment, "indent");
    justifyLastLine_ = bool_attr(alignment, "justifyLastLine");
    shrinkToFit_     = bool_attr(alignment, "shrinkToFit");
    textRotation_    = int_attr(alignment, "textRotation");
  }

  rapidxml::xml_node<>* protection = xf->first_node("protection");
  if (protection == NULL) {
    locked_          = NA_LOGICAL;
    hidden_          = NA_LOGICAL;
  } else {
    locked_          = bool_attr(protection, "locked");
    hidden_          = bool_attr(protection, "hidden");
  }
}

String xf::readingOrder(rapidxml::xml_node<>* node) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute("readingOrder");
  if (attribute != NULL) {
    return(readingOrderChr_[std::stoi(attribute->value())]);
  }
  return(NA_STRING);
}
