#ifndef COLOR_
#define COLOR_

#include <Rcpp.h>
#include "rapidxml.h"
#include "styles.h"

class color {

  public:

    Rcpp::CharacterVector rgb_;
    Rcpp::IntegerVector   theme_;
    Rcpp::IntegerVector   index_;

    color(rapidxml::xml_node<>* color, styles& styles);
};

#endif
