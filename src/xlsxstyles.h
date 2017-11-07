#ifndef XLSXSTYLES_
#define XLSXSTYLES_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xf.h"
#include "font.h"
#include "fill.h"
#include "border.h"

struct colors {
  Rcpp::CharacterVector rgb;
  Rcpp::CharacterVector theme;
  Rcpp::IntegerVector   indexed;
  Rcpp::NumericVector   tint;
  colors() {}
  colors(int n):
    rgb     (n, NA_STRING),
    theme   (n, NA_STRING),
    indexed (n, NA_INTEGER),
    tint    (n, NA_REAL)
  {}
};

class xlsxstyles {

  public:

    Rcpp::CharacterVector theme_name_; // name of theme, e.g. "accent1"
    Rcpp::CharacterVector theme_;      // rgb equivalent of theme no.
    Rcpp::CharacterVector indexed_;    // rgb equivalent of index no.

    std::vector<xf> cellXfs_;

    std::vector<xf> cellStyleXfs_;
    Rcpp::CharacterVector cellStyles_; // names of cell styles, ordered by xfId
    std::map<int, std::string> cellStyles_map_; // map of cell style names, for lookup by xfId

    Rcpp::CharacterVector numFmts_;
    Rcpp::LogicalVector isDate_;

    std::vector<font> fonts_;
    std::vector<fill> fills_;
    std::vector<border> borders_;

    std::vector<xf> style_formats_; // built up by applyFormats() from xf definitions
    std::vector<xf> local_formats_; // built up by applyFormats() from xf definitions

    Rcpp::List style_; // inside-out List version of style_formats_
    Rcpp::List local_; // inside-out List version of local_formats_

    xlsxstyles(const std::string& path);

    void cacheThemeRgb(const std::string& path);
    void cacheIndexedRgb();

    void cacheCellXfs(rapidxml::xml_node<>* styleSheet);
    void cacheCellStyleXfs(rapidxml::xml_node<>* styleSheet); // also cellStyles, if available
    void cacheNumFmts(rapidxml::xml_node<>* styleSheet);
    void cacheFonts(rapidxml::xml_node<>* styleSheet);
    void cacheFills(rapidxml::xml_node<>* styleSheet);
    void cacheBorders(rapidxml::xml_node<>* styleSheet);

    // Insert values of one color object into vectors in another
    void clone_color(color& from, colors& to, int& i);
    Rcpp::List list_color(colors& original, bool is_style); // Convert struct to list

    void applyFormats(); // Build each style on top of the normal style
    Rcpp::List zipFormats(std::vector<xf> styles, bool is_style); // Turn the formats inside-out to return to R

};

#endif
