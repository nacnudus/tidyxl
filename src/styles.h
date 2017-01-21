#ifndef STYLES_
#define STYLES_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xf.h"
#include "font.h"
#include "fill.h"
#include "border.h"

struct colors {
  Rcpp::CharacterVector rgb;
  Rcpp::IntegerVector   theme;
  Rcpp::IntegerVector   indexed;
  Rcpp::NumericVector   tint;
};

class styles {

  public:

    Rcpp::CharacterVector theme_;   // rgb equivalent of theme no.
    Rcpp::CharacterVector indexed_;   // rgb equivalent of index no.

    std::vector<xf> cellXfs_;

    std::vector<xf> cellStyleXfs_;
    Rcpp::CharacterVector cellStyles_; // names of cell styles, if available

    Rcpp::CharacterVector numFmts_;
    Rcpp::LogicalVector isDate_;

    std::vector<font> fonts_;
    std::vector<fill> fills_;
    std::vector<border> borders_;

    xf style_formats_; // built up by applyFormats() from xf definitions
    xf local_formats_; // built up by applyFormats() from xf definitions

    Rcpp::List style_; // inside-out List version of style_formats_
    Rcpp::List local_; // inside-out List version of local_formats_

    styles(const std::string& path);

    void cacheThemeRgb(const std::string& path);
    void cacheIndexedRgb();

    void cacheCellXfs(rapidxml::xml_node<>* styleSheet);
    void cacheCellStyleXfs(rapidxml::xml_node<>* styleSheet); // also cellStyles, if available

    void cacheNumFmts(rapidxml::xml_node<>* styleSheet);
    bool isDateFormat(std::string formatCode);

    void cacheFonts(rapidxml::xml_node<>* styleSheet);
    void cacheFills(rapidxml::xml_node<>* styleSheet);
    void cacheBorders(rapidxml::xml_node<>* styleSheet);

    void clone_color(color& from, colors& to); // Append values in one color to vectors in another
    Rcpp::List list_color(colors& original); // Convert struct to list

    void applyFormats(); // Build each style on top of the normal style
    Rcpp::List zipFormats(xf styles); // Turn the formats inside-out to return to R

};

#endif
