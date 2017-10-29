#ifndef tidyxl_xml_value_
#define tidyxl_xml_value_

#include <Rcpp.h>

// bool attribute value, defaults to true
inline int bool_attr(rapidxml::xml_node<>* node, const char* name) {
  std::string value;
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute == NULL) {
    return(true);
  }
  value = attribute->value();
  if (value == "0" || value == "false") {
    return(false);
  }
  return(true);
}

// bool attribute value, defaults to no change of given Rcpp::IntegerVector
inline void bool_attr(Rcpp::IntegerVector& out, unsigned long long int i,
    rapidxml::xml_node<>* node, const char* name) {
  std::string value;
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    value = attribute->value();
    if (value == "0" || value == "false") {
      out[i] = false;
    } else {
      out[i] = true;
    }
  }
}

// int attribute value, defaults to NA_INTEGER
inline int int_attr(rapidxml::xml_node<>* node, const char* name) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    return(std::stoi(attribute->value()));
  }
  return(NA_INTEGER);
}

// int attribute value, defaults to no change of given Rcpp::IntegerVector
inline void int_attr(Rcpp::IntegerVector& out, unsigned long long int i,
    rapidxml::xml_node<>* node, const char* name) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    out[i] = std::stoi(attribute->value());
  }
}

// double attribute value, defaults to NA_REAL
inline double double_attr(rapidxml::xml_node<>* node, const char* name) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    return(std::stod(attribute->value()));
  }
  return(NA_REAL);
}

// double attribute value, defaults to no change of given Rcpp::NumericVector
inline void double_attr(Rcpp::NumericVector& out, unsigned long long int i,
    rapidxml::xml_node<>* node, const char* name) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    out[i] = std::stod(attribute->value());
  }
}

// Rcpp::String attribute value, defaults to NA_STRING
inline Rcpp::String string_attr(rapidxml::xml_node<>* node, const char* name) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    return(attribute->value());
  }
  return(NA_STRING);
}

// Rcpp::String attribute value, defaults to no change of given
// Rcpp::CharacterVector
inline void string_attr(Rcpp::CharacterVector out, unsigned long long int i,
    rapidxml::xml_node<>* node, const char* name) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    out[i] = attribute->value();
  }
}

#endif
