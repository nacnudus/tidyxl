// [[Rcpp::plugins("cpp11")]]

#include <algorithm>
#include <Rcpp.h>
#include "zip.h"
#include "xlsxnames.h"
#include "xlsxvalidation.h"
#include "xlsxbook.h"
#include "xlsxstyles.h"
#include "date.h"

using namespace Rcpp;

// This is the entry-point from R into Rcpp.  The main function, xlsx_read_,
// instantiates the xlsxbook class, and then passes that by reference to one or
// more xlsxsheet instances, iterating through them and returning a list of
// 'sheets'.

// [[Rcpp::export]]
List xlsx_cells_(
    std::string path,
    CharacterVector sheet_paths,
    CharacterVector sheet_names,
    CharacterVector comments_paths
    ) {
  xlsxbook book(path, sheet_paths, sheet_names, comments_paths);
  return book.information_;
}

// [[Rcpp::export]]
List xlsx_formats_(std::string path) {
  xlsxstyles styles(path);
  return List::create(
    _["local"] = styles.local_,
    _["style"] = styles.style_);
}

inline String comments_path_(std::string path, std::string sheet_target) {
  // Given a sheet id, return the path to the comments file, should one exist
  std::string sheet_rels = "xl/worksheets/_rels/" + sheet_target.replace(0, 11, "") + ".rels";
  if (zip_has_file(path, sheet_rels)) {
    std::string targets_text = zip_buffer(path, sheet_rels);
    rapidxml::xml_document<> targets_xml;
    targets_xml.parse<rapidxml::parse_strip_xml_namespaces>(&targets_text[0]);
    rapidxml::xml_node<>* relationships = targets_xml.first_node("Relationships");
    for (rapidxml::xml_node<>* relationship = relationships->first_node("Relationship");
        relationship; relationship = relationship->next_sibling()) {
      std::string target = relationship->first_attribute("Target")->value();
      if (target.substr(0, 11) == "../comments") {
        // Return the comments file path
        return(target.replace(0, 2, "xl"));
      }
    }
  }
  return NA_STRING;
}

// [[Rcpp::export]]
List xlsx_sheet_files_(std::string path) {
  // Return a list of worksheets,  their index numbers, names, and comments
  // paths.

  // primary key of two 'tables' of worksheets and other objects
  std::string id;
  std::vector<std::string> ids;

  // Get the id and target of worksheets, and look up any comments file
  std::map<std::string, std::string> targets;
  std::map<std::string, String> comments;
  std::string rels_text = zip_buffer(path, "xl/_rels/workbook.xml.rels");
  rapidxml::xml_document<> rels_xml;
  rels_xml.parse<rapidxml::parse_strip_xml_namespaces>(&rels_text[0]);
  rapidxml::xml_node<>* relationships = rels_xml.first_node("Relationships");
  for (rapidxml::xml_node<>* relationship = relationships->first_node("Relationship");
      relationship; relationship = relationship->next_sibling()) {
    std::string target = relationship->first_attribute("Target")->value();
    std::string target_type = target.substr(0, 10);
    if (target_type == "worksheets" || target_type == "chartsheet") {
       // Only store worksheets and chartsheets -- requests for chartsheets are
       // handled in the R wrapper
      id = relationship->first_attribute("Id")->value();
      ids.push_back(id);
      targets.insert({id, "xl/" + target}) ;
      comments.insert({id, comments_path_(path, target)});
    }
  }

  // Get the id, name and sheetId (display order) of all sheets/charts/etc.
  std::map<std::string, std::string> names;
  std::map<std::string, int> sheetIds;
  std::string workbook_text = zip_buffer(path, "xl/workbook.xml");
  rapidxml::xml_document<> workbook_xml;
  workbook_xml.parse<rapidxml::parse_strip_xml_namespaces>(&workbook_text[0]);
  rapidxml::xml_node<>* workbook = workbook_xml.first_node("workbook");
  rapidxml::xml_node<>* sheets = workbook->first_node("sheets");
  for (rapidxml::xml_node<>* sheet = sheets->first_node("sheet");
      sheet; sheet = sheet->next_sibling()) {
    rapidxml::xml_attribute<>* r_id = sheet->first_attribute("id");
    if (r_id != NULL) {
      id = r_id->value();
    } else {
      stop("Invalid xl/workbook.xml: sheet element lacks id attribute"); // # nocov
    }
    names[id] = sheet->first_attribute("name")->value();
    sheetIds[id] = strtol(sheet->first_attribute("sheetId")->value(), NULL, 10);
  }

  // Join by id
  std::vector<std::string> out_name;
  std::vector<int> out_rId;
  std::vector<int> out_sheetId;
  std::vector<std::string> out_target;
  CharacterVector out_comments_path;
  for(std::vector<std::string>::iterator it = ids.begin(); it != ids.end(); ++it) {
    std::string key(*it);
    out_name.push_back(names[key]);
    out_rId.push_back(std::strtol(key.substr(3, std::string::npos).c_str(), NULL, 10));
    out_sheetId.push_back(sheetIds[key]);
    out_target.push_back(targets[key]);
    out_comments_path.push_back(comments[key]);
  }

  // Return a data frame
  List out = List::create(
      _["name"] = out_name,
      _["id"] = out_rId,
      _["index"] = out_sheetId,
      _["sheet_path"] = out_target,
      _["comments_path"] = out_comments_path);

  // Turn list of vectors into a data frame without checking anything
  int n = Rf_length(out[0]);
  out.attr("class") = CharacterVector::create("tbl_df", "tbl", "data.frame");
  out.attr("row.names") = IntegerVector::create(NA_INTEGER, -n); // Dunno how this works (the -n part)

  return out;
}

// [[Rcpp::export]]
List xlsx_validation_(
    std::string path,
    CharacterVector sheet_paths,
    CharacterVector sheet_names
    ) {
  return xlsxvalidation(path, sheet_paths, sheet_names).information();
}

// [[Rcpp::export]]
List xlsx_names_(std::string path) {
  return xlsxnames(path).information();
}

// [[Rcpp::export]]
LogicalVector is_date_format_(CharacterVector formats) {
  CharacterVector::iterator it;
  std::vector<bool> out(formats.size());
  size_t i(0);
  for(it = formats.begin(); it != formats.end(); ++it) {
    std::string format(*it);
    out[i] = isDateFormat(format);
    ++i;
  }
  return wrap(out);
}

// [[Rcpp::export]]
List xlsx_color_theme_(std::string path) {
  CharacterVector theme_name =
    CharacterVector::create("background1",
                            "text1",
                            "background2",
                            "text2",
                            "accent1",
                            "accent2",
                            "accent3",
                            "accent4",
                            "accent5",
                            "accent6",
                            "hyperlink",
                            "followed-hyperlink");
  CharacterVector theme_rgb(12, NA_STRING);
  std::string FF = "FF";
  if (zip_has_file(path, "xl/theme/theme1.xml")) {
    std::string theme1 = zip_buffer(path, "xl/theme/theme1.xml");
    rapidxml::xml_document<> theme1_xml;
    theme1_xml.parse<0>(&theme1[0]);
    rapidxml::xml_node<>* theme = theme1_xml.first_node("a:theme");
    rapidxml::xml_node<>* themeElements = theme->first_node("a:themeElements");
    rapidxml::xml_node<>* clrScheme = themeElements->first_node("a:clrScheme");

    // First, two sysClr nodes in the wrong order
    rapidxml::xml_node<>* color = clrScheme->first_node();
    theme_rgb[1] = FF + color->first_node()->first_attribute("lastClr")->value();
    color = color->next_sibling();
    theme_rgb[0] = FF + color->first_node()->first_attribute("lastClr")->value();

    // Then, two srgbClr nodes in the wrong order
    color = color->next_sibling();
    theme_rgb[3] = FF + color->first_node()->first_attribute("val")->value();
    color = color->next_sibling();
    theme_rgb[2] = FF + color->first_node()->first_attribute("val")->value();

    // Finally, eight more srgbClr nodes in the correct order
    // Can't reuse 'color' here, so use 'nextcolor'
    int i = 4;
    for (rapidxml::xml_node<>* nextcolor = color->next_sibling();
        nextcolor; nextcolor = nextcolor->next_sibling()) {
      theme_rgb[i] = FF + nextcolor->first_node()->first_attribute("val")->value();
      i++;
    }
  }

  List out = List::create(_["theme"] = theme_name,
                          _["rgb"] = theme_rgb);

  // Turn list of vectors into a data frame without checking anything
  int n = Rf_length(out[0]);
  out.attr("class") = CharacterVector::create("tbl_df", "tbl", "data.frame");
  out.attr("row.names") = IntegerVector::create(NA_INTEGER, -n); // Dunno how this works (the -n part)

  return out;
}
