#ifndef TIDYXL_STRING_
#define TIDYXL_STRING_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxstyles.h"
#include <R_ext/GraphicsDevice.h> // Rf_ucstoutf8 is exported in R_ext/GraphicsDevice.h

// Based on tidyverse/readxl
// unescape an ST_Xstring. See 22.9.2.19 [p3786]
inline std::string unescape(const std::string& s) {
  std::string out;
  out.reserve(s.size());

  for (size_t i = 0; i < s.size(); i++) {
    if (i+6 < s.size() && s[i] == '_' && s[i+1] == 'x'
        && isxdigit(s[i+2]) && isxdigit(s[i+3])
        && isxdigit(s[i+4]) && isxdigit(s[i+5]) && s[i+6] == '_') {
      // extract character
      unsigned int ch = strtoul(&s[i+2], NULL, 16);
      char utf8[16]; // 16 from definition of Rf_ucstoutf8
      Rf_ucstoutf8(utf8, ch);
      out += utf8;
      i += 6; // skip to the final '_'
    } else {
      out.push_back(s[i]);
    }
  }

  return out;
}

// Based on tidyverse/readxl
// Parser for <si> and <is> inlineStr tags CT_Rst [p3893]
// returns true if a string is found, false if missing.
inline void parseString(const rapidxml::xml_node<>* string, std::string& out) {
  out.clear();
  const rapidxml::xml_node<>* t = string->first_node("t");
  if (t != NULL) {
    // According to the spec (CT_Rst, p3893) a single <t> element
    // may coexist with zero or more <r> elements.
    //
    // However, software that read these files do not appear to exclusively
    // follow this spec.
    //
    // MacOSX preview, considers only the <t> element if found and ignores any
    // additional r elements.
    //
    //
    // Excel 2010 appears to produce only elements with r elements or with a
    // single t element and no mixtures. It will, however, consider an element
    // with a single t element followed by one or more r elements as valid,
    // concatenating the results. Any other combination of r and t elements
    // is considered invalid.
    //
    // We read the <t> tag, if present, first, then concatenate any <r> tags.
    // All Excel 2010 sheets will read correctly under this regime.
    out = unescape(std::string(t->value(), t->value_size()));
  }
  // iterate over all r elements
  for (const rapidxml::xml_node<>* r = string->first_node("r"); r != NULL;
      r = r->next_sibling("r")) {
    // a unique t element should be present (CT_RElt [p3893])
    // but MacOSX preview just ignores chunks with no t element present
    const rapidxml::xml_node<>* t = r->first_node("t");
    if (t != NULL) {
      out += unescape(std::string(t->value(), t->value_size()));
    }
  }
}

// Return a dataframe, one row per substring, columns for formatting
inline Rcpp::List parseFormattedString(
    const rapidxml::xml_node<>* string,
    xlsxstyles& styles) {

  int n(0); // number of substrings
  for (rapidxml::xml_node<>* node = string->first_node();
       node != NULL;
       node = node->next_sibling()) {
    ++n;
  }

  Rcpp::CharacterVector character(n, NA_STRING);
  Rcpp::LogicalVector bold(n, NA_LOGICAL);
  Rcpp::LogicalVector italic(n, NA_LOGICAL);
  Rcpp::CharacterVector underline(n, NA_STRING);
  Rcpp::LogicalVector strike(n, NA_LOGICAL);
  Rcpp::CharacterVector vertAlign(n, NA_STRING);
  Rcpp::NumericVector size(n, NA_REAL);
  Rcpp::CharacterVector color_rgb(n, NA_STRING);
  Rcpp::IntegerVector color_theme(n, NA_INTEGER);
  Rcpp::NumericVector color_tint(n, NA_REAL);
  Rcpp::IntegerVector color_indexed(n, NA_INTEGER);
  Rcpp::CharacterVector font(n, NA_STRING);
  Rcpp::IntegerVector family(n, NA_INTEGER);
  Rcpp::CharacterVector scheme(n, NA_STRING);

  int i(0); // ith substring
  for (rapidxml::xml_node<>* node = string->first_node();
       node != NULL;
       node = node->next_sibling()) {

    std::string node_name = node->name();
    if (node_name == "t") {

      SET_STRING_ELT(character, i, Rf_mkCharCE(node->value(), CE_UTF8));

    } else if (node_name == "r") {

      SET_STRING_ELT(character, i, Rf_mkCharCE(node->first_node("t")->value(), CE_UTF8));
      rapidxml::xml_node<>* rPr = node->first_node("rPr");

      if (rPr != NULL) {

        bold[i] = rPr->first_node("b") != NULL;
        italic[i] = rPr->first_node("i") != NULL;

        rapidxml::xml_node<>* u = rPr->first_node("u");
        if (u != NULL) {
          rapidxml::xml_attribute<>* u_val = u->first_attribute("val");
          if (u_val != NULL) {
            underline[i] = u_val->value();
          } else {
            underline[i] = "single";
          }
        }

        strike[i] = rPr->first_node("strike") != NULL;

        rapidxml::xml_node<>* vertAlign_node = rPr->first_node("vertAlign");
        if (vertAlign_node != NULL) {
          vertAlign[i] = vertAlign_node->first_attribute("val")->value();
        }

        rapidxml::xml_node<>* sz = rPr->first_node("sz");
        if (sz != NULL) {
          size[i] = strtod(sz->value(), NULL);
        }

        rapidxml::xml_node<>* color = rPr->first_node("color");
        if (color != NULL) {

          rapidxml::xml_attribute<>* rgb_attr = color->first_attribute("rgb");
          if (rgb_attr != NULL) {

            color_rgb[i] = rgb_attr->value();

          } else {

            rapidxml::xml_attribute<>* theme_attr = color->first_attribute("theme");
            if (theme_attr != NULL) {

              int theme_int = strtol(theme_attr->value(), NULL, 10) + 1;
              color_rgb[i] = styles.theme_[theme_int - 1];
              color_theme[i] = theme_int;

              rapidxml::xml_attribute<>* tint_attr = color->first_attribute("tint");
              if (tint_attr != NULL) {
                color_tint[i] = strtod(tint_attr->value(), NULL);
              }

            } else {

              // no known case
              // # nocov start
              rapidxml::xml_attribute<>* indexed_attr = color->first_attribute("indexed");
              if (indexed_attr != NULL) {
                int indexed_int = strtol(indexed_attr->value(), NULL, 10) + 1;
                color_rgb[i] = styles.indexed_[indexed_int - 1];
                color_indexed[i] = indexed_int;
              }
              // # nocov end

            }
          }
        }

        rapidxml::xml_node<>* rFont = rPr->first_node("rFont");
        if (rFont != NULL) {
          font[i] = rFont->first_attribute("val")->value();
        }

        rapidxml::xml_node<>* family_node = rPr->first_node("family");
        if (family_node != NULL) {
          family[i] = strtol(family_node->first_attribute("val")->value(), NULL, 10);
        }

        rapidxml::xml_node<>* scheme_node = rPr->first_node("scheme");
        if (scheme_node != NULL) {
          scheme[i] = scheme_node->first_attribute("val")->value();
        }

      }
    }
    ++i;
  }

  Rcpp::List out = Rcpp::List::create(
      Rcpp::_["character"] = character,
      Rcpp::_["bold"] = bold,
      Rcpp::_["italic"] = italic,
      Rcpp::_["underline"] = underline,
      Rcpp::_["strike"] = strike,
      Rcpp::_["vertAlign"] = vertAlign,
      Rcpp::_["size"] = size,
      Rcpp::_["color_rgb"] = color_rgb,
      Rcpp::_["color_theme"] = color_theme,
      Rcpp::_["color_indexed"] = color_indexed,
      Rcpp::_["color_tint"] = color_tint,
      Rcpp::_["font"] = font,
      Rcpp::_["family"] = family,
      Rcpp::_["scheme"] = scheme);

  // Turn list of vectors into a data frame without checking anything
  out.attr("class") = Rcpp::CharacterVector::create("tbl_df", "tbl", "data.frame");
  out.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, -n); // Dunno how this works (the -n part)

  return out;
}

#endif
