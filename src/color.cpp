#include <Rcpp.h>
#include "rapidxml.h"
#include "color.h"

using namespace Rcpp;

color::color(rapidxml::xml_node<>* color, styles& styles) {
  rapidxml::xml_attribute<>* rgb = color->first_attribute("rgb");
  if (rgb != NULL) {
    rgb_[0] = rgb->value();
  } else {
    rgb_[0] = NA_STRING;
  }

  rapidxml::xml_attribute<>* theme = color->first_attribute("theme");
  if (theme != NULL) {
    int theme_int = strtol(theme->value(), NULL, 10);
    theme_[0] = theme_int;
    rgb_[0] = styles.theme_[theme_int - 1];
  } else {
    theme_[0] = NA_INTEGER;
  }

  rapidxml::xml_attribute<>* index = color->first_attribute("index");
  if (index != NULL) {
    int index_int = strtol(index->value(), NULL, 10);
    index_[0] = index_int;
    rgb_[0] = styles.index_[index_int - 1];
  } else {
    index_[0] = NA_INTEGER;
  }
}
