#ifndef XLSXBOOK_
#define XLSXBOOK_

#include <Rcpp.h>
#include "rapidxml.h"

class xlsxbook {

rapidxml::xml_document<> xml_;
rapidxml::xml_node<>* workbook_;
std::vector<std::string> sheets_;

public:

  xlsxbook(const std::string& path); // constructor
  std::vector<std::string> sheets(); // sheet names

private:

  void cacheSheets(rapidxml::xml_node<>* sheets); // sheet names

};

#endif
