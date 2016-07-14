#ifndef XLSXBOOK_
#define XLSXBOOK_

#include <Rcpp.h>
#include "rapidxml.h"

class xlsxbook {

rapidxml::xml_document<> bookXml_;
rapidxml::xml_node<>* rootNode_;
rapidxml::xml_node<>* sheets_;

public:

  xlsxbook(const std::string& bookPath);
  Rcpp::List information();

};

#endif
