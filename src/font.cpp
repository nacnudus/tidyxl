#include <Rcpp.h>
#include "rapidxml.h"
#include "font.h"
#include "xlsxstyles.h"
#include "color.h"

using namespace Rcpp;

font::font() {} // Default constructor

font::font(rapidxml::xml_node<>* font,
    xlsxstyles* styles
    ): color_(font->first_node("color"), styles)
{
  // Initialize variables
  b_         = LogicalVector::create(false);
  i_         = LogicalVector::create(false);
  u_         = NA_STRING;

  strike_    = LogicalVector::create(false);
  vertAlign_ = NA_STRING;
  size_      = NumericVector::create(NA_REAL);

  name_      = NA_STRING;
  family_    = IntegerVector::create(NA_INTEGER);
  scheme_    = NA_STRING;

  rapidxml::xml_node<>* b = font->first_node("b");
  if (b != NULL) {
    b_[0] = true;
  }

  rapidxml::xml_node<>* i = font->first_node("i");
  if (i != NULL) {
    i_[0] = true;
  }

  rapidxml::xml_node<>* u = font->first_node("u");
  if (u != NULL) {
    rapidxml::xml_attribute<>* val = u->first_attribute("val");
    if (val != NULL) {
      u_ = val->value();
    } else {
      u_ = "single";
    }
  }

  rapidxml::xml_node<>* strike = font->first_node("strike");
  if (strike != NULL) {
    strike_[0] = true;
  }

  rapidxml::xml_node<>* vertAlign = font->first_node("vertAlign");
  if (vertAlign != NULL) {
    rapidxml::xml_attribute<>* val = vertAlign->first_attribute("val");
    if (val != NULL) {
      vertAlign_ = val->value();
    }
  }

  // Excel seems to use 'sz', while googlesheets seems to use "size"
  rapidxml::xml_node<>* size = font->first_node("size");
  rapidxml::xml_node<>* sz = font->first_node("sz");
  if (size != NULL) {
    size_[0] = strtod(size->first_attribute("val")->value(), NULL);
  } else {
    if (sz != NULL) {
      size_[0] = strtod(sz->first_attribute("val")->value(), NULL);
    }
  }

  rapidxml::xml_node<>* name = font->first_node("name");
  if (name != NULL) {
    name_ = name->first_attribute("val")->value();
  }

  rapidxml::xml_node<>* family = font->first_node("family");
  if (family != NULL) {
    family_[0] = strtol(family->first_attribute("val")->value(), NULL, 10);
  }

  rapidxml::xml_node<>* scheme = font->first_node("scheme");
  if (scheme != NULL) {
    scheme_ = scheme->first_attribute("val")->value();
  }
}
