#ifndef XLSXBOOK_
#define XLSXBOOK_

#include <Rcpp.h>
#include "rapidxml.h"

class xlsxbook {

public:

const std::string& path_;
std::vector<std::string> sheets_;
std::vector<std::string> strings_;
unsigned long int strings_size_;
std::vector<unsigned long int> cellXfs_xfId_;

  xlsxbook(const std::string& path);    // constructor
  std::string path();                   // workbook path
  std::vector<std::string>& sheets();   // sheet names
  std::vector<std::string>& strings();  // strings table
  unsigned long int strings_size();     // length of strings table
  std::vector<unsigned long int>& cellXfs_xfId(); // link from cellXfs from cellStyleXfs

  void cacheSheets(rapidxml::xml_node<>* sheets); // sheet names
  void cacheStrings();
  void cacheCellXfsXfId(); // 'xfId' of 'cellXfs' in 'styles.xml' to link cells with themes

};

#endif
