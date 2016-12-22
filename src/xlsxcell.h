#ifndef XLSXCELL_
#define XLSXCELL_

#include <Rcpp.h>
#include "rapidxml.h"
#include "xlsxbook.h"
#include "xlsxsheet.h"

class xlsxcell {

  public:

    /* celltype celltype;     // TODO: cell type enumeration */

    /* // TODO: Parse the value into specific types */
    /* // The following are NA unless relevant to the cell type, but if NA then */
    /* // they should never be sought, because celltype should flag which is relevant. */
    /* bool         logical_;    // Parsed value */
    /* double       numeric_;    // Parsed value */
    /* double       date_;       // Parsed value */
    /* Rcpp::String character_;  // Parsed value */
    /* Rcpp::String error_;      // Parsed value */

    xlsxcell(
        rapidxml::xml_node<>* cell, // the cell node,
        xlsxsheet* sheet,           // the parent worksheet
        xlsxbook& book,             // the parent workbook
        unsigned long long int& i   // the index of the cell
        );

    void parseAddress( rapidxml::xml_node<>* cell, xlsxsheet* sheet, unsigned long long int& i);
    void cacheValue( rapidxml::xml_node<>* cell, xlsxsheet* sheet, xlsxbook& book, unsigned long long int& i);

    // Not the whole format, but only the index of the local and style 
    // formats (cellXfs and cellStyleXfs)
    void cacheFormat( rapidxml::xml_node<>* cell, xlsxsheet* sheet, xlsxbook& book, unsigned long long int& i);

    void cacheFormula( rapidxml::xml_node<>* cell, xlsxsheet* sheet, unsigned long long int& i);
};

#endif
