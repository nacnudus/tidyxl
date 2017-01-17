#ifndef BORDER_
#define BORDER_

class styles; // Forward declaration because fill is included in styles.

#include <Rcpp.h>
#include "rapidxml.h"
#include "stroke.h"

class border {

  public:

    Rcpp::IntegerVector diagonalDown_;
    Rcpp::IntegerVector diagonalUp_;
    Rcpp::IntegerVector outline_;
    stroke  left_;
    stroke  right_;
    stroke  start_; // gnumeric
    stroke  end_;   // gnumeric
    stroke  top_;
    stroke  bottom_;
    stroke  diagonal_;
    stroke  vertical_;
    stroke  horizontal_;

    border(); // Default constructor
    border(bool blank); // construct empty vectors
    border(rapidxml::xml_node<>* border, styles* styles);
};

#endif
