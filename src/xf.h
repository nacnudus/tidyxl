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
  static const std::vector<std::string> readingOrderChr_; // defined in tidyxl.cpp
  Rcpp::CharacterVector horizontal_;
  Rcpp::CharacterVector vertical_;
  Rcpp::LogicalVector   wrapText_;
  Rcpp::CharacterVector readingOrder_; // 0=context, 1=left-to-right, 2=right-to-left
  Rcpp::IntegerVector   indent_;
  Rcpp::LogicalVector   justifyLastLine_;
  Rcpp::LogicalVector   shrinkToFit_;
  Rcpp::IntegerVector   textRotation_;

  // protection
  Rcpp::LogicalVector   locked_;
  Rcpp::LogicalVector   hidden_;

  // index into styles
  Rcpp::IntegerVector   xfId_;

  // whether to apply the format at this level
  Rcpp::LogicalVector   applyNumberFormat_;
  Rcpp::LogicalVector   applyFont_;
  Rcpp::LogicalVector   applyFill_;
  Rcpp::LogicalVector   applyBorder_;
  Rcpp::LogicalVector   applyAlignment_;
  Rcpp::LogicalVector   applyProtection_;

    xf(); // Default constructor
    xf(rapidxml::xml_node<>* xf);

    // boolean value of an attribute
    Rcpp::LogicalVector bool_value(rapidxml::xml_node<>* xf, const char* name);

    // integer value of an attribute
    Rcpp::IntegerVector int_value(rapidxml::xml_node<>* xf, const char* name);

    // string value of an attribute
    Rcpp::CharacterVector string_value(rapidxml::xml_node<>* xf, const char* name);

    // looked-up value of readingOrder attribute
    Rcpp::CharacterVector readingOrder(rapidxml::xml_node<>* xf);
};

#endif
