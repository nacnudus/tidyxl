#ifndef XLSXSHEET_
#define XLSXSHEET_

#include <Rcpp.h>
#include "rapidxml.h"

class xlsxsheet {

rapidxml::xml_document<> sheetXml_;
rapidxml::xml_node<>* rootNode_;
rapidxml::xml_node<>* sheetData_;

public:

  xlsxsheet(const std::string& bookPath, const int& i);
  Rcpp::List information();

};

#endif
