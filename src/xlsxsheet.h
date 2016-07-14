#ifndef XLSXSHEET_
#define XLSXSHEET_

#include <Rcpp.h>
#include "rapidxml.h"

class xlsxsheet {

public:

  xlsxsheet(const std::string& bookPath, const int& i);
  Rcpp::List information();

};

#endif


