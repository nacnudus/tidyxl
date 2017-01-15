#ifndef COLOR_
#define COLOR_

#include <Rcpp.h>
#include "rapidxml.h"

class clrScheme {

  public:

    Rcpp::CharacterVector rgb_;

    color(rapidxml::xml_node<>* color);
};

#endif
