#ifndef BORDER_
#define BORDER_

class xlsxstyles; // Forward declaration because fill is included in xlsxstyles.

#include <Rcpp.h>
#include "rapidxml.h"
#include "stroke.h"

class border {

  public:

    int    diagonalDown_;
    int    diagonalUp_;
    int    outline_;
    stroke left_;
    stroke right_;
    stroke start_; // gnumeric
    stroke end_;   // gnumeric
    stroke top_;
    stroke bottom_;
    stroke diagonal_;
    stroke vertical_;
    stroke horizontal_;

    border(rapidxml::xml_node<>* border, xlsxstyles* styles);
};

#endif
