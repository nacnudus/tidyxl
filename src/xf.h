#ifndef XF_
#define XF_

#include <Rcpp.h>
#include "rapidxml.h"

class xf {
  // ECMA part 1 page 1753

  public:

    // basics
    int numFmtId_;
    int fontId_;
    int fillId_;
    int borderId_;


    // alignment
    std::vector<std::string>  readingOrderChr_; // lookup values of readingOrder
    Rcpp::String horizontal_;
    Rcpp::String vertical_;
    int          wrapText_;
    Rcpp::String readingOrder_; // 0=context, 1=left-to-right, 2=right-to-left
    int          indent_;
    int          justifyLastLine_;
    int          shrinkToFit_;
    int          textRotation_;

    // protection
    int locked_;
    int hidden_;

    // index into styles
    int xfId_;

    // whether to apply the format at this level
    int applyNumberFormat_;
    int applyFont_;
    int applyFill_;
    int applyBorder_;
    int applyAlignment_;
    int applyProtection_;

    xf(); // Default constructor
    xf(rapidxml::xml_node<>* xf);

    // looked-up value of readingOrder attribute
    Rcpp::String readingOrder(rapidxml::xml_node<>* xf);
};

#endif
