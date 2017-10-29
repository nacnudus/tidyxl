#ifndef tidyxl_utils_
#define tidyxl_utils_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxstyles.h"
#include <R_ext/GraphicsDevice.h> // Rf_ucstoutf8 is exported in R_ext/GraphicsDevice.h

// Follow tidyverse/readxl
// https://github.com/tidyverse/readxl/commit/63ef215f57322dd5d7a27799a2a3fe463bd39fc7
// in correcting subsecond rounding by converting serial date to
// decimilliseconds, using a well-known hack to round to nearest integer w/o
// C++11 or boost, e.g.
// http://stackoverflow.com/questions/485525/round-for-float-in-c, and
// converting back to serial date.
inline double dateRound(double date) {
  double ms = date * 10000;
  ms = (ms >= 0.0 ? std::floor(ms + 0.5) : std::ceil(ms - 0.5));
  return ms / 10000;
}

inline std::string ref(std::string& sheet, std::string& cell) {
  return "'" + sheet + "'!" + cell;
}

// Follow tidyverse/readxl
// https://github.com/tidyverse/readxl/commit/c9a54ae9ce0394808f6d22e8ef1a7a647b2d92bb
// by correcting for Excel's faithful re-implementation of the Lotus 1-2-3 bug,
// in which February 29, 1900 is included in the date system, even though 1900
// was not actually a leap year
// https://support.microsoft.com/en-us/help/214326/excel-incorrectly-assumes-that-the-year-1900-is-a-leap-year
// How we address this: If date is *prior* to the non-existent leap day: add a
// day If date is on the non-existent leap day: make negative and, in due
// course, NA Otherwise: do nothing
inline double checkDate(double& date, int& dateSystem, int& dateOffset, std::string ref) {
  if (dateSystem == 1900 && date < 61) {
    date = (date < 60) ? date + 1 : -1;
  }
  if (date < 0) {
    Rcpp::warning("NA inserted for impossible 1900-02-29 datetime: " + ref);
    return NA_REAL;
  } else {
    return dateRound((date - dateOffset) * 86400);
  }
}

// Convert datetime doubles to strings "%Y-%m-%d %H:%M:%S"
// TODO: Support subseconds
inline std::string formatDate(double& date, int& dateSystem, int& dateOffset) {
  const char* format;
  if (date < 1) {                      // time only (in Excel's formuat)
    format = "%H:%M:%S";
  } else {                             // date and maybe time
    /* double intpart; */
    /* if (modf(date, &intpart) == 0.0) { // date only, not time */
    /*   format = "%Y-%m-%d"; */
    /* } else {                           // date and time */
      format = "%Y-%m-%d %H:%M:%S";
    /* } */
  }
  /* // This doesn't work on Windows before 1970, so use R functions instead */
  /* char out[charsize]; */
  /* date = checkDate(date, dateSystem, dateOffset, ""); // Convert to POSIX */
  /* std::time_t date_time_t(date); */
  /* std::tm date_tm = *std::gmtime(&date_time_t); */
  /* std::strftime(out, sizeof(out), format, &date_tm); */
  // Use R functions instead of C++ strftime, which doesn't work before 1970 on
  // Windows
  date = checkDate(date, dateSystem, dateOffset, ""); // Convert to POSIX
  std::string out;
  Rcpp::Function r_as_POSIXlt("as.POSIXlt");
  Rcpp::Function r_format_POSIXlt("format.POSIXlt");
  out =  Rcpp::as<std::string>(r_format_POSIXlt(r_as_POSIXlt(date,
                                                             "",
                                                             "1970-01-01"),
                                                format,
                                                false));
  return out;
}

// Based on hadley/readxl
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

// Based on hadley/readxl
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

    } else {

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

              rapidxml::xml_attribute<>* indexed_attr = color->first_attribute("indexed");
              if (indexed_attr != NULL) {
                int indexed_int = strtol(indexed_attr->value(), NULL, 10) + 1;
                color_rgb[i] = styles.indexed_[indexed_int - 1];
                color_indexed[i] = indexed_int;
              }

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
