#ifndef COLOR_
#define COLOR_

class styles; // Forward declaration because color is included in font, which is
              // included in styles.

#include <Rcpp.h>
#include "rapidxml.h"

class color {

  public:

    Rcpp::CharacterVector rgb_;
    Rcpp::IntegerVector   theme_;
    Rcpp::IntegerVector   indexed_;
    Rcpp::NumericVector   tint_;

    color(); // default constructor
    color(bool blank); // construct empty vectors
    color(rapidxml::xml_node<>* color, styles* styles);
};

#endif
