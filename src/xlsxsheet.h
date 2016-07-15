#ifndef XLSXSHEET_
#define XLSXSHEET_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxbook.h"

class xlsxsheet {

std::string name_;
rapidxml::xml_document<> sheetXml_;
rapidxml::xml_node<>* rootNode_;
rapidxml::xml_node<>* sheetData_;

public:

  xlsxsheet(const std::string& bookPath,
      const int& i, xlsxbook& book);
  Rcpp::List information();

};

#endif
