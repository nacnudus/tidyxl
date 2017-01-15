#include <Rcpp.h>
#include "rapidxml.h"
#include "color.h"
#include "styles.h"

using namespace Rcpp;

color::color() { // Default constructor
  // Initialize variables
  rgb_     = CharacterVector::create(NA_STRING);
  theme_   = IntegerVector::create(NA_INTEGER);
  indexed_ = IntegerVector::create(NA_INTEGER);
  tint_    = NumericVector::create(NA_REAL);
}

color::color(bool blank) { // Construct empty vectors
}

color::color(rapidxml::xml_node<>* color, styles* styles) {
  // Initialize variables
  rgb_     = CharacterVector::create(NA_STRING);
  theme_   = IntegerVector::create(NA_INTEGER);
  indexed_ = IntegerVector::create(NA_INTEGER);
  tint_    = NumericVector::create(NA_REAL);

  if (color != NULL) {
    rapidxml::xml_attribute<>* _auto = color->first_attribute("auto");
    if (_auto != NULL) {
      // Colour is system-dependent, which tidyxl assumes means black.
      rgb_[0] = "FF000000";
    } else {
      rapidxml::xml_attribute<>* rgb = color->first_attribute("rgb");
      if (rgb != NULL) {
        rgb_[0] = rgb->value();
      }

      rapidxml::xml_attribute<>* theme = color->first_attribute("theme");
      if (theme != NULL) {
        int theme_int = strtol(theme->value(), NULL, 10) + 1;
        theme_[0] = theme_int;
        rgb_[0] = styles->theme_[theme_int - 1];
      }

      rapidxml::xml_attribute<>* indexed = color->first_attribute("indexed");
      if (indexed != NULL) {
        int indexed_int = strtol(indexed->value(), NULL, 10) + 1;
        indexed_[0] = indexed_int;
        rgb_[0] = styles->indexed_[indexed_int - 1];
      }

      rapidxml::xml_attribute<>* tint = color->first_attribute("tint");
      if (tint != NULL) {
        tint_[0] = strtof(tint->value(), NULL);
      }
    }
  }
}
