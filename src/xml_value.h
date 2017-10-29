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

// bool attribute value, defaults to no change of given int
inline int bool_attr(int& out, rapidxml::xml_node<>* node, const char* name) {
  std::string value;
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    value = attribute->value();
    if (value == "0" || value == "false") {
      out = false;
    } else {
      out = true;
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

// int attribute value, defaults to no change of given int
inline int int_attr(int& out, rapidxml::xml_node<>* node, const char* name) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    out = std::stoi(attribute->value());
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

// int attribute value, defaults to no change of given int
inline double double_attr(double& out, rapidxml::xml_node<>* node, const char* name) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    out = std::stod(attribute->value());
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

// Rcpp::String attribute value, defaults to no change of givin Rcpp::String
inline Rcpp::String string_attr(Rcpp::String& out, rapidxml::xml_node<>* node, const char* name) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    return(attribute->value());
  }
}

#endif
