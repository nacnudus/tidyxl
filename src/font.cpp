#include <Rcpp.h>
#include "rapidxml.h"
#include "font.h"
#include "styles.h"
#include "color.h"

using namespace Rcpp;

font::font() {} // Default constructor

font::font(rapidxml::xml_node<>* font,
    styles* styles
    ): color_(font->first_node("color"), styles)
{
  // Initialize variables
  b_         = LogicalVector::create(NA_LOGICAL);
  i_         = LogicalVector::create(NA_LOGICAL);
  u_         = CharacterVector::create(NA_STRING);

  strike_    = LogicalVector::create(NA_LOGICAL);
  vertAlign_ = CharacterVector::create(NA_STRING);
  size_      = IntegerVector::create(NA_INTEGER);

  name_      = CharacterVector::create(NA_STRING);
  family_    = IntegerVector::create(NA_INTEGER);
  scheme_    = CharacterVector::create(NA_STRING);

  rapidxml::xml_node<>* b = font->first_node("b");
  if (b != NULL) {
    b_[0] = true;
  } else {
    b_[0] = false;
  }

  rapidxml::xml_node<>* i = font->first_node("i");
  if (i != NULL) {
    i_[0] = true;
  } else {
    i_[0] = false;
  }

  rapidxml::xml_node<>* u = font->first_node("u");
  if (u != NULL) {
    rapidxml::xml_attribute<>* val = u->first_attribute("val");
    if (val != NULL) {
      u_[0] = val->value();
    } else {
      u_[0] = "single";
    }
  } else {
    u_[0] = NA_STRING;
  }

  rapidxml::xml_node<>* strike = font->first_node("strike");
  if (strike != NULL) {
    strike_[0] = true;
  } else {
    strike_[0] = false;
  }

  rapidxml::xml_node<>* vertAlign = font->first_node("vertAlign");
  if (vertAlign != NULL) {
    rapidxml::xml_attribute<>* val = vertAlign->first_attribute("val");
    if (val != NULL) {
      vertAlign_[0] = val->value();
    } else {
      vertAlign_[0] = "baseline";
    }
  } else {
    vertAlign_[0] = NA_STRING;
  }

  // Excel seems to use 'sz', while googlesheets seems to use "size"
  rapidxml::xml_node<>* size = font->first_node("size");
  rapidxml::xml_node<>* sz = font->first_node("sz");
  if (size == NULL) {
    if (sz == NULL) {
      size_ = NA_INTEGER;
    } else {
      size_ = strtol(sz->first_attribute("val")->value(), NULL, 10);
    }
  } else {
    size_ = strtol(size->first_attribute("val")->value(), NULL, 10);
  }

  rapidxml::xml_node<>* name = font->first_node("name");
  if (name != NULL) {
    rapidxml::xml_attribute<>* val = name->first_attribute("val");
    name_[0] = name->first_attribute("val")->value();
  } else {
    name_[0] = NA_STRING;
  }

  rapidxml::xml_node<>* family = font->first_node("family");
  if (family != NULL) {
    rapidxml::xml_attribute<>* val = family->first_attribute("val");
    family_[0] = strtol(family->first_attribute("val")->value(), NULL, 10);
  } else {
    family_[0] = NA_INTEGER;
  }

  rapidxml::xml_node<>* scheme = font->first_node("scheme");
  if (scheme != NULL) {
    rapidxml::xml_attribute<>* val = scheme->first_attribute("val");
    scheme_[0] = scheme->first_attribute("val")->value();
  } else {
    scheme_[0] = NA_STRING;
  }
}
