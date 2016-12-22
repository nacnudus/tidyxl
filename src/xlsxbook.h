#ifndef XLSXBOOK_
#define XLSXBOOK_

#include <Rcpp.h>
#include "rapidxml.h"

class xlsxbook {

  public:

    const std::string& path_;                     // workbook path
    std::vector<std::string> sheets_;             // sheet names
    std::vector<std::string> strings_;            // strings table
    std::vector<unsigned long int> cellXfs_xfId_; // link from cellXfs from cellStyleXfs

    xlsxbook(const std::string& path);    // constructor

    void cacheSheets(rapidxml::xml_node<>* sheets); // sheet names, not data
    void cacheStrings();
    void cacheCellXfsXfId(); // 'xfId' of 'cellXfs' in 'styles.xml' to link cells with themes

};

#endif
