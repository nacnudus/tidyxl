#ifndef XLSXCELL_
#define XLSXCELL_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxbook.h"
#include "xlsxsheet.h"

class xlsxcell {

  std::string address_;
  int col_;
  int row_;

  public:

    xlsxcell(
        rapidxml::xml_node<>* cell, // the cell node,
        xlsxsheet* sheet,           // the parent worksheet
        xlsxbook& book,             // the parent workbook
        unsigned long long int& i   // the index of the cell
        );

    void parseAddress(
        rapidxml::xml_node<>* cell,
        xlsxsheet* sheet,
        xlsxbook& book,
        unsigned long long int& i
        );

    void cacheValue(
        rapidxml::xml_node<>* cell,
        xlsxsheet* sheet,
        xlsxbook& book,
        unsigned long long int& i
        );

    void cacheFormula(
        rapidxml::xml_node<>* cell,
        xlsxsheet* sheet,
        xlsxbook& book,
        unsigned long long int& i
        );

};

#endif
