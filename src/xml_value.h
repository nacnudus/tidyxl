#ifndef tidyxl_xml_value_
#define tidyxl_xml_value_

#include <Rcpp.h>

// bool attribute value, with default
inline int bool_attr(rapidxml::xml_node<>* node, const char* name,
    int default_ = NA_LOGICAL) {
  std::string value;
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute == NULL) {
    return(default_);
  }
  value = attribute->value();
  if (value == "0" || value == "false") {
    return(false);
  }
  return(true);
}

// bool attribute value, defaults to no change of given bool
inline int bool_attr(bool& out, rapidxml::xml_node<>* node, const char* name,
    int default_ = NA_LOGICAL) {
  std::string value;
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    value = attribute->value();
    if (value == "0" || value == "false") {
      out = false;
    } else {
      out == true;
    }
  }
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

// int attribute value, with default
inline int int_attr(rapidxml::xml_node<>* node, const char* name,
    int default_ = NA_INTEGER) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    return(std::stoi(attribute->value()));
  }
  return(default_);
}

// int attribute value, defaults to no change of given int
inline void int_attr(int& out, unsigned long long int i,
    rapidxml::xml_node<>* node, const char* name) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    out = std::stoi(attribute->value());
  }
}

// int attribute value, defaults to no change of given Rcpp::IntegerVector
inline void int_attr(Rcpp::IntegerVector& out, unsigned long long int i,
    rapidxml::xml_node<>* node, const char* name) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    out[i] = std::stoi(attribute->value());
  }
}

// double attribute value, with default
inline double double_attr(int& out, rapidxml::xml_node<>* node, const char* name,
    double default_ = NA_REAL) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    out = std::stod(attribute->value());
  }
}

// double attribute value, defaults to no change of given double
inline double double_attr(double& out, rapidxml::xml_node<>* node, const char* name,
    double default_ = NA_REAL) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    out = std::stod(attribute->value());
  }
}

// double attribute value, defaults to no change of given Rcpp::NumericVector
inline void double_attr(Rcpp::NumericVector& out, unsigned long long int i,
    rapidxml::xml_node<>* node, const char* name) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    out[i] = std::stod(attribute->value());
  }
}

// Rcpp::String attribute value, with default
inline Rcpp::String string_attr(rapidxml::xml_node<>* node, const char* name,
    Rcpp::String default_ = NA_STRING) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    return(attribute->value());
  }
  return(default_);
}

// Rcpp::String attribute value, defaults to no change of given Rcpp::String
inline Rcpp::String string_attr(Rcpp::String& out, rapidxml::xml_node<>* node,
    const char* name, Rcpp::String default_ = NA_STRING) {
  rapidxml::xml_attribute<>* attribute = node->first_attribute(name);
  if (attribute != NULL) {
    out = attribute->value();
  }
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
