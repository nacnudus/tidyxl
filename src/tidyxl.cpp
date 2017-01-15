#include <algorithm>
#include <Rcpp.h>
#include "zip.h"
#include "xlsxsheet.h"
#include "xlsxbook.h"

using namespace Rcpp;

// This is the entry-point from R into Rcpp.  The main function, xlsx_read_,
// instantiates the xlsxbook class, and then passes that by reference to one or
// more xlsxsheet instances, iterating through them and returning a list of
// 'sheets'.

// [[Rcpp::export]]
List xlsx_read_(
    std::string path,
    IntegerVector sheets,
    CharacterVector names,
    CharacterVector comments_paths
    ) {
  // Parse book-level information (e.g. styles, themes, strings, date system)
  xlsxbook book(path);

  // Loop through sheets
  List sheet_list(sheets.size());

  IntegerVector::iterator in_it;
  List::iterator sheet_list_it;

  int i = 0;
  for(in_it = sheets.begin(), sheet_list_it = sheet_list.begin(); in_it != sheets.end();
      ++in_it, ++sheet_list_it) {
    Rcpp::String comments_path(comments_paths[i]);
    *sheet_list_it = xlsxsheet(*in_it, book, comments_path).information();
    ++i;
  }

  sheet_list.attr("names") = names;

  List out = List::create(
      _["data"] = sheet_list,
      _["formats"] = List::create(
        _["local"] = book.styles_.local_,
        _["style"] = book.styles_.style_));

  return out;
}

inline String comments_path_(std::string path, std::string id_string) {
  // Given a sheet id, return the path to the comments file, should one exist
  std::string sheet_rels = "xl/worksheets/_rels/sheet" + id_string + ".xml.rels";
  if (zip_has_file(path, sheet_rels)) {
    std::string targets_text = zip_buffer(path, "xl/worksheets/_rels/sheet" + id_string + ".xml.rels");
    rapidxml::xml_document<> targets_xml;
    targets_xml.parse<0>(&targets_text[0]);
    rapidxml::xml_node<>* relationships = targets_xml.first_node("Relationships");
    for (rapidxml::xml_node<>* relationship = relationships->first_node("Relationship");
        relationship; relationship = relationship->next_sibling()) {
      std::string target_attr = relationship->first_attribute("Target")->value();
      if (target_attr.substr(0, 11) == "../comments") {
        // Return the comments file path
        return(target_attr.replace(0, 2, "xl"));
      }
    }
  }
  return NA_STRING;
}

// [[Rcpp::export]]
DataFrame xlsx_sheets_(std::string path) {
  // Return a list of worksheets,  their index numbers, names, and comments
  // paths.

  // Get the ids of worksheets (rather than chart sheets, etc.)
  std::string targets_text = zip_buffer(path, "xl/_rels/workbook.xml.rels");
  rapidxml::xml_document<> targets_xml;
  targets_xml.parse<0>(&targets_text[0]);
  std::vector<std::string> id_targets;
  rapidxml::xml_node<>* relationships = targets_xml.first_node("Relationships");
  for (rapidxml::xml_node<>* relationship = relationships->first_node("Relationship");
      relationship; relationship = relationship->next_sibling()) {
    std::string target_attr = relationship->first_attribute("Target")->value();
    if (target_attr.substr(0, 10) == "worksheets") {
      // Only store worksheets
      id_targets.push_back(relationship->first_attribute("Id")->value());
    }
  }

  // Get the ids and names of all sheets/charts/etc. to be looked up later
  std::string names_text = zip_buffer(path, "xl/workbook.xml");
  rapidxml::xml_document<> names_xml;
  names_xml.parse<0>(&names_text[0]);
  std::map<std::string, std::string> id_names;
  rapidxml::xml_node<>* workbook = names_xml.first_node("workbook");
  rapidxml::xml_node<>* sheets = workbook->first_node("sheets");
  for (rapidxml::xml_node<>* sheet = sheets->first_node("sheet");
      sheet; sheet = sheet->next_sibling()) {
    id_names[sheet->first_attribute("r:id")->value()] = sheet->first_attribute("name")->value();
  }

  // Look up the names of worksheets only, and get the paths to any comments
  // files
  IntegerVector ids;
  CharacterVector names;
  CharacterVector comments_paths;
  for(std::vector<std::string>::iterator it = id_targets.begin(); it != id_targets.end(); ++it) {
    std::string id_string(*it);
    // Look up the name by the id
    names.push_back(id_names[id_string]);
    // Extract the integer index from the id
    id_string.replace(0, 3, "");
    ids.push_back(strtol(id_string.c_str(), NULL, 10));
    // Fetch the path to the comments file, should one exist
    comments_paths.push_back(comments_path_(path, id_string));
  }

  // Return a data frame
  DataFrame out = DataFrame::create(
      _["index"] = ids,
      _["name"] = names,
      _["comments_path"] = comments_paths,
      _["stringsAsFactors"] = false);

  return(out);
}
