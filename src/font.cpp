#include <Rcpp.h>
#include "rapidxml.h"
#include "font.h"
#include "xlsxstyles.h"
#include "color.h"

using namespace Rcpp;

font::font(rapidxml::xml_node<>* font,
    xlsxstyles* styles
    ): color_(font->first_node("color"), styles)
{
  rapidxml::xml_node<>* b = font->first_node("b");
  b_ = b != NULL;

  rapidxml::xml_node<>* i = font->first_node("i");
  i_ = i != NULL;

  rapidxml::xml_node<>* u = font->first_node("u");
  if (u != NULL) {
    rapidxml::xml_attribute<>* val = u->first_attribute("val");
    if (val != NULL) {
      u_ = val->value();
    } else {
      u_ = "single";
    }
  } else {
    u_ = NA_STRING;
  }

  rapidxml::xml_node<>* strike = font->first_node("strike");
  strike_ = strike != NULL;

  rapidxml::xml_node<>* vertAlign = font->first_node("vertAlign");
  if (vertAlign != NULL) {
    rapidxml::xml_attribute<>* val = vertAlign->first_attribute("val");
    if (val != NULL) {
      vertAlign_ = val->value();
    } else {
      vertAlign_ = NA_STRING;
    }
  } else {
    vertAlign_ = NA_STRING;
  }

  // Excel seems to use 'sz', while googlesheets seems to use "size"
  rapidxml::xml_node<>* size = font->first_node("size");
  rapidxml::xml_node<>* sz = font->first_node("sz");
  if (size != NULL) {
    size_ = strtod(size->first_attribute("val")->value(), NULL);
  } else {
    if (sz != NULL) {
      size_ = strtod(sz->first_attribute("val")->value(), NULL);
    } else {
      size_ = NA_REAL;
    }
  }

  rapidxml::xml_node<>* name = font->first_node("name");
  if (name != NULL) {
    name_ = name->first_attribute("val")->value();
  } else {
    name_ = NA_STRING;
  }

  rapidxml::xml_node<>* family = font->first_node("family");
  if (family != NULL) {
    family_ = strtol(family->first_attribute("val")->value(), NULL, 10);
  } else {
    family_ = NA_INTEGER;
  }

  rapidxml::xml_node<>* scheme = font->first_node("scheme");
  if (scheme != NULL) {
    scheme_ = scheme->first_attribute("val")->value();
  } else {
    scheme_ = NA_STRING;
  }
}
