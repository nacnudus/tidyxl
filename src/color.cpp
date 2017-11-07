#include <Rcpp.h>
#include "rapidxml.h"
#include "color.h"
#include "xlsxstyles.h"

using namespace Rcpp;

color::color() { // Default constructor
  // Initialize variables
  rgb_     = NA_STRING;
  theme_   = NA_STRING;
  indexed_ = NA_INTEGER;
  tint_    = NA_REAL;
}

color::color(rapidxml::xml_node<>* color, xlsxstyles* styles) {
  // Initialize variables
  rgb_     = NA_STRING;
  theme_   = NA_STRING;
  indexed_ = NA_INTEGER;
  tint_    = NA_REAL;

  if (color != NULL) {
    rapidxml::xml_attribute<>* _auto = color->first_attribute("auto");
    if (_auto != NULL) {
      // Colour is system-dependent, which tidyxl assumes means black.
      rgb_ = "FF000000";
    } else {
      rapidxml::xml_attribute<>* rgb = color->first_attribute("rgb");
      if (rgb != NULL) {
        rgb_ = rgb->value();
      }

      rapidxml::xml_attribute<>* theme = color->first_attribute("theme");
      if (theme != NULL) {
        int theme_int = strtol(theme->value(), NULL, 10) ;
        theme_ = styles->theme_name_[theme_int];
        rgb_ = styles->theme_[theme_int];
      }

      rapidxml::xml_attribute<>* indexed = color->first_attribute("indexed");
      if (indexed != NULL) {
        int indexed_int = strtol(indexed->value(), NULL, 10) + 1;
        indexed_ = indexed_int;
        rgb_ = styles->indexed_[indexed_int - 1];
      }

      rapidxml::xml_attribute<>* tint = color->first_attribute("tint");
      if (tint != NULL) {
        tint_ = strtod(tint->value(), NULL);
      }
    }
  }
}
