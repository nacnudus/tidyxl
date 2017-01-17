#ifndef XF_
#define XF_

#include <Rcpp.h>
#include "rapidxml.h"

class xf {
  // ECMA part 1 page 1753

  public:

  // basics
  Rcpp::IntegerVector   numFmtId_;
  Rcpp::IntegerVector   fontId_;
  Rcpp::IntegerVector   fillId_;
  Rcpp::IntegerVector   borderId_;

  // alignment
  Rcpp::CharacterVector horizontal_;
  Rcpp::CharacterVector vertical_;
  Rcpp::IntegerVector   wrapText_;
  Rcpp::IntegerVector   readingOrder_;
  Rcpp::IntegerVector   indent_;
  Rcpp::IntegerVector   justifyLastLine_;
  Rcpp::IntegerVector   shrinkToFit_;
  Rcpp::IntegerVector   textRotation_;

  // protection
  Rcpp::IntegerVector   locked_;
  Rcpp::IntegerVector   hidden_;

  // index into styles
  Rcpp::IntegerVector   xfId_;

  // whether to apply the format at this level
  Rcpp::IntegerVector   applyNumberFormat_;
  Rcpp::IntegerVector   applyFont_;
  Rcpp::IntegerVector   applyFill_;
  Rcpp::IntegerVector   applyBorder_;
  Rcpp::IntegerVector   applyAlignment_;
  Rcpp::IntegerVector   applyProtection_;

    xf(); // Default constructor
    xf(rapidxml::xml_node<>* xf);

    // integer value of an attribute
    Rcpp::IntegerVector int_value(rapidxml::xml_node<>* xf, const char* name);

    // string value of an attribute
    Rcpp::CharacterVector string_value(rapidxml::xml_node<>* xf, const char* name);
};

#endif

