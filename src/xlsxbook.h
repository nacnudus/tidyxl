#ifndef XLSXBOOK_
#define XLSXBOOK_

#include <Rcpp.h>
#include "rapidxml.h"

class xlsxbook {

rapidxml::xml_document<> bookXml_;
rapidxml::xml_node<>* rootNode_;
std::vector<std::string> sheets_;

public:

  xlsxbook(const std::string& path); // constructor
  Rcpp::List information();
  const std::vector<std::string>& sheets(); // sheet names

private:

  void cacheSheets(rapidxml::xml_node<>* sheets);

};

#endif
