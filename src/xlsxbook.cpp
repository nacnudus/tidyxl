#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "xlsxbook.h"
#include "xlsxsheet.h"
#include "styles.h"
#include "utils.h"

using namespace Rcpp;

xlsxbook::xlsxbook(const std::string& path): path_(path), styles_(path_) {
  std::string book = zip_buffer(path_, "xl/workbook.xml");

  rapidxml::xml_document<> xml;
  xml.parse<0>(&book[0]);

  rapidxml::xml_node<>* workbook = xml.first_node("workbook");

  cacheDateOffset(workbook); // Must come before cacheSheets
  cacheStrings();
}

xlsxbook::xlsxbook(
    const std::string& path,
    CharacterVector& sheet_paths,
    CharacterVector& sheet_names,
    CharacterVector& comments_paths):
  path_(path),
  sheet_paths_(sheet_paths),
  sheet_names_(sheet_names),
  comments_paths_(comments_paths),
  styles_(path_) {
  std::string book = zip_buffer(path_, "xl/workbook.xml");

  rapidxml::xml_document<> xml;
  xml.parse<0>(&book[0]);

  rapidxml::xml_node<>* workbook = xml.first_node("workbook");

  cacheDateOffset(workbook); // Must come before cacheSheets
  cacheStrings();
  cacheSheetXml();
  cacheInformation();
}

// Based on hadley/readxl
void xlsxbook::cacheStrings() {
  if (!zip_has_file(path_, "xl/sharedStrings.xml"))
    return;

  std::string xml = zip_buffer(path_, "xl/sharedStrings.xml");
  rapidxml::xml_document<> sharedStrings;
  sharedStrings.parse<0>(&xml[0]);

  rapidxml::xml_node<>* sst = sharedStrings.first_node("sst");
  rapidxml::xml_attribute<>* uniqueCount = sst->first_attribute("uniqueCount");
  if (uniqueCount != NULL) {
    unsigned long int n = strtol(uniqueCount->value(), NULL, 10);
    strings_.reserve(n);
  }

  // 18.4.8 si (String Item) [p1725]
  for (rapidxml::xml_node<>* string = sst->first_node();
      string; string = string->next_sibling()) {
    std::string out;
    parseString(string, out);    // missing strings are treated as empty ""
    strings_.push_back(out);
  }
}

void xlsxbook::cacheDateOffset(rapidxml::xml_node<>* workbook) {
  rapidxml::xml_node<>* workbookPr = workbook->first_node("workbookPr");
  if (workbookPr != NULL) {
    rapidxml::xml_attribute<>* date1904 = workbookPr->first_attribute("date1904");
    if (date1904 != NULL) {
      std::string is1904 = date1904->value();
      if ((is1904 == "1") || (is1904 == "true")) {
        dateSystem_ = 1904;
        dateOffset_ = 24107;
        return;
      }
    }
  }

  dateSystem_ = 1900;
  dateOffset_ = 25569;
}

void xlsxbook::cacheSheetXml() {
  // Loop through sheets, reading the xml into memory
  List sheet_list(sheet_paths_.size());

  CharacterVector::iterator in_it;
  List::iterator sheet_list_it;

  int i = 0;
  for(in_it = sheet_paths_.begin(); in_it != sheet_paths_.end(); ++in_it) {
    std::string xml(*in_it);
    sheet_xml_.push_back(zip_buffer(path_, xml));
  }
}

void xlsxbook::cacheInformation() {
  // Loop through sheets
  List sheet_list(sheet_paths_.size());

  std::vector<std::string>::iterator in_it;
  List::iterator sheet_list_it;

  int i = 0;
  for(in_it = sheet_xml_.begin(), sheet_list_it = sheet_list.begin();
      in_it != sheet_xml_.end();
      ++in_it, ++sheet_list_it) {
    String sheet_name(sheet_names_[i]);
    String comments_path(comments_paths_[i]);
    xlsxsheet sheet(sheet_name, *in_it, *this, comments_path);
    rapidxml::xml_document<> doc;
    doc.parse<0>(&(*in_it)[0]);
    rapidxml::xml_node<>* workbook = doc.first_node("worksheet");
    rapidxml::xml_node<>* sheetData = workbook->first_node("sheetData");
    sheet.parseSheetData(sheetData);
    sheet.appendComments();
    *sheet_list_it = sheet.information();
    ++i;
  }

  sheet_list.attr("names") = sheet_names_;

  information_ = List::create(
      _["data"] = sheet_list,
      _["formats"] = List::create(
        _["local"] = styles_.local_,
        _["style"] = styles_.style_));
}
