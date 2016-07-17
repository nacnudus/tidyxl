#ifndef tidyxl_utils_
#define tidyxl_utils_

#include <Rcpp.h>
#include "rapidxml.h"
#include <R_ext/GraphicsDevice.h> // Rf_ucstoutf8 is exported in R_ext/GraphicsDevice.h


inline void makeDataFrame(Rcpp::List& Lists, const Rcpp::CharacterVector& colnames) {
  // turn List of Vectors into a data frame without checking anything
  int n = Rf_length(Lists[0]);
  Lists.attr("class") = Rcpp::CharacterVector::create("tbl_df", "tbl", "data.frame");
  Lists.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, -n); // Dunno how this works
  Lists.attr("names") = colnames;
}

// Based on hadley/readxl
// Simple parser: does not check that order of numbers and letters is correct
// row_ and column_ are one-based
inline void parseAddress(std::string& address, int& row, int& col) {
  col = 0;
  row = 0;

  // Iterate though std::string character by character
  for(std::string::const_iterator iter = address.begin();
      iter != address.end(); ++iter) {
    if (*iter >= '0' && *iter <= '9') { // If it's a number
      row = row * 10 + (*iter - '0'); // Then multiply existing row by 10 and add new number
    } else if (*iter >= 'A' && *iter <= 'Z') { // If it's a character
      col = 26 * col + (*iter - 'A' + 1); // Then do similarly  with columns
    } else {
      Rcpp::stop("Invalid character '%s' in cell ref '%s'", *iter, address);
    }
  }
}

inline void getChildValueString(Rcpp::String& container,
    const char* childname, rapidxml::xml_node<>* parent) {
  rapidxml::xml_node<>* child = parent->first_node(childname);
  if (child == NULL) {
    container = NA_STRING;
  } else {
    container = child->value();
  }
}

inline void getAttributeValueString(Rcpp::String& container,
    const char* attributename, rapidxml::xml_node<>* parent) {
  rapidxml::xml_attribute<>* attribute = parent->first_attribute(attributename);
  if (attribute == NULL) {
    container = NA_STRING;
  } else {
    container = attribute->value();
  }
}

inline void getAttributeValueDouble(Rcpp::String& container,
    const char* attributename, rapidxml::xml_node<>* parent) {
  rapidxml::xml_attribute<>* attribute = parent->first_attribute(attributename);
  if (attribute == NULL) {
    container = NA_REAL;
  } else {
    container = atof(attribute->value());
  }
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
inline bool parseString(const rapidxml::xml_node<>* string, std::string& out) {
  bool found = false;
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
    out = unescape(t->value());
    found = true;
  }
  // iterate over all r elements
  for (const rapidxml::xml_node<>* r = string->first_node("r"); r != NULL;
       r = r->next_sibling("r")) {
    // a unique t element should be present (CT_RElt [p3893])
    // but MacOSX preview just ignores chunks with no t element present
    const rapidxml::xml_node<>* t = r->first_node("t");
    if (t != NULL) {
      out += unescape(t->value());
      found = true;
    }
  }
  return found;
}


#endif
