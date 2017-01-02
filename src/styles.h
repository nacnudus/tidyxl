#ifndef STYLES_
#define STYLES_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xf.h"

class styles {

  public:

    Rcpp::CharacterVector theme_;   // rgb equivalent of theme no.
    Rcpp::CharacterVector index_;   // rgb equivalent of index no.

    std::vector<xf> cellXfs_;

    std::vector<xf> cellStyleXfs_;
    Rcpp::CharacterVector cellStyles_; // names of cell styles, if available

    Rcpp::CharacterVector numFmts_;
    Rcpp::LogicalVector isDate_;

    styles(const std::string& path);

    void cacheThemeRgb(const std::string& path);
    void cacheIndexRgb();

    void cacheCellXfs(rapidxml::xml_node<>* styleSheet);
    void cacheCellStyleXfs(rapidxml::xml_node<>* styleSheet); // also cellStyles, if available

    void cacheNumFmts(rapidxml::xml_node<>* styleSheet);
    bool isDateFormat(std::string formatCode);

};

#endif
