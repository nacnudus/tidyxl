#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "xlsxvalidation.h"
#include "xlsxbook.h"
#include "date.h"

using namespace Rcpp;

int count_validations(std::string& xml) {
  // Count the number of validation rules in a file
  int n(0);
  rapidxml::xml_document<> doc;
  doc.parse<rapidxml::parse_non_destructive>(&xml[0]);
  rapidxml::xml_node<>* worksheet = doc.first_node("worksheet");
  rapidxml::xml_node<>* dataValidations =
    worksheet->first_node("dataValidations");
  if (dataValidations != NULL) {
    // Initialize the vectors to the number of rules, using the count attribute
    // if it exists, otherwise count the rules.
    rapidxml::xml_attribute<>* count =
      dataValidations->first_attribute("count");
    if (count != NULL) {
      n = strtol(count->value(), NULL, 10);
    } else {
      // No known case, so # nocov start
      for (rapidxml::xml_node<>* dataValidation =
             dataValidations->first_node("dataValidation");
           dataValidation;
           dataValidation = dataValidation->next_sibling("dataValidation")) {
          ++n;
      // # nocov end
      }
    }
  }
  return n;
}

void parseValidations(
    xlsxvalidation& validation,
    xlsxbook& book,
    std::string& sheet_name,
    std::string& xml,
    int& i) {
  rapidxml::xml_document<> doc;
  doc.parse<0>(&xml[0]);
  rapidxml::xml_node<>* worksheet = doc.first_node("worksheet");
  rapidxml::xml_node<>* dataValidations =
    worksheet->first_node("dataValidations");

  for (rapidxml::xml_node<>* dataValidation =
         dataValidations->first_node("dataValidation");
        dataValidation;
        dataValidation = dataValidation->next_sibling("dataValidation")) {

    validation.sheet_[i] = sheet_name;

    rapidxml::xml_attribute<>* sqref = dataValidation->first_attribute("sqref");
    if (sqref != NULL) {
      // Replace spaces with commas.  For space-delimited 'sqref' attributes,
      // which ought to be comma-delimited for consistency with the union
      // operator ',' in formulas.  Modifies in-place.
      std::string ref(sqref->value());
      std::replace(ref.begin(), ref.end(), ' ', ',');
      validation.ref_[i] = ref;
    }

    std::string type_string;
    rapidxml::xml_attribute<>* type = dataValidation->first_attribute("type");
    if (type != NULL) {
      // operator defaults to 'between', but when type is one of list or
      // operator then it should be NA
      type_string = type->value();
      validation.type_[i] = type_string;
      if (type_string == "list" || type_string == "custom")
        validation.operator_[i] = NA_STRING;
    }

    rapidxml::xml_attribute<>* Operator =
      dataValidation->first_attribute("operator");
    if (Operator != NULL)
      validation.operator_[i] = Operator->value();

    rapidxml::xml_node<>* formula1 = dataValidation->first_node("formula1");
    if (formula1 != NULL) {
      if (type_string == "date" || type_string == "time") {
        double date_double = strtod(formula1->value(), NULL);
        validation.formula1_[i] =
          formatDate(date_double, book.dateSystem_, book.dateOffset_);
      } else {
        validation.formula1_[i] = formula1->value();
      }
    }

    rapidxml::xml_node<>* formula2 = dataValidation->first_node("formula2");
    if (formula2 != NULL) {
      if (type_string == "date" || type_string == "time") {
        double date_double = strtod(formula2->value(), NULL);
        validation.formula2_[i] =
          formatDate(date_double, book.dateSystem_, book.dateOffset_);
      } else {
        validation.formula2_[i] = formula2->value();
      }
    }

    rapidxml::xml_attribute<>* allowBlank =
      dataValidation->first_attribute("allowBlank");
    if (allowBlank != NULL)
      validation.allow_blank_[i] = strtod(allowBlank->value(), NULL);

    rapidxml::xml_attribute<>* showInputMessage =
      dataValidation->first_attribute("showInputMessage");
    if (showInputMessage != NULL)
      validation.show_input_message_[i] =
        strtod(showInputMessage->value(), NULL);

    rapidxml::xml_attribute<>* promptTitle =
      dataValidation->first_attribute("promptTitle");
    if (promptTitle != NULL)
      validation.prompt_title_[i] = promptTitle->value();

    rapidxml::xml_attribute<>* prompt =
      dataValidation->first_attribute("prompt");
    if (prompt != NULL)
      validation.prompt_[i] = prompt->value();

    rapidxml::xml_attribute<>* showErrorMessage =
      dataValidation->first_attribute("showErrorMessage");
    if (showErrorMessage != NULL)
      validation.show_error_message_[i] =
        strtod(showErrorMessage->value(), NULL);

    rapidxml::xml_attribute<>* errorTitle =
      dataValidation->first_attribute("errorTitle");
    if (errorTitle != NULL)
      validation.error_title_[i] = errorTitle->value();

    rapidxml::xml_attribute<>* error = dataValidation->first_attribute("error");
    if (error != NULL)
      validation.error_[i] = error->value();

    rapidxml::xml_attribute<>* errorStyle =
      dataValidation->first_attribute("errorStyle");
    if (errorStyle != NULL)
      validation.error_style_[i] = errorStyle->value();

    ++i;
  }
}

xlsxvalidation::xlsxvalidation(
    std::string path,
    CharacterVector sheet_paths,
    CharacterVector sheet_names) {
  // Parse book-level information for the date system (1900/1904)
  xlsxbook book(path);

  // Loop through sheets
  List out(sheet_paths.size());

  std::vector<std::string> sheets_xml; // xml of each sheet
  std::vector<int> rules_count; // number of rules in individual sheets
  int n(0); // total number of rules in all sheets
  int i(0); // ith validation rule

  std::vector<std::string>::iterator sheet_xml;
  CharacterVector::iterator sheet_path;
  CharacterVector::iterator sheet_name;
  std::vector<int>::iterator rule_count;

  // Count the total number of rules in all sheets
  for(sheet_path = sheet_paths.begin();
      sheet_path != sheet_paths.end();
      ++sheet_path) {
    std::string path(*sheet_path);
    std::string xml = zip_buffer(book.path_, path);
    sheets_xml.push_back(xml);
    int count = count_validations(xml);
    rules_count.push_back(count);
    n += count;
  }

  // initialise output vectors
  sheet_              = CharacterVector(n, NA_STRING);
  ref_                = CharacterVector(n, NA_STRING);
  type_               = CharacterVector(n, NA_STRING);
  operator_           = CharacterVector(n, "between");
  formula1_           = CharacterVector(n, NA_STRING);
  formula2_           = CharacterVector(n, NA_STRING);
  allow_blank_        = LogicalVector(n,   false);
  show_input_message_ = LogicalVector(n,   false);
  prompt_title_       = CharacterVector(n, NA_STRING);
  prompt_             = CharacterVector(n, NA_STRING);
  show_error_message_ = LogicalVector(n,   false);
  error_title_        = CharacterVector(n, NA_STRING);
  error_              = CharacterVector(n, NA_STRING);
  error_style_        = CharacterVector(n, "stop");

  // Parse all the rules
  for(sheet_xml = sheets_xml.begin(),
      rule_count = rules_count.begin(),
      sheet_name = sheet_names.begin();
      sheet_xml != sheets_xml.end();
      ++sheet_xml, ++rule_count, ++sheet_name) {
    if (*rule_count != 0) {
      std::string name(*sheet_name);
      parseValidations(*this, book, name, *sheet_xml, i);
    }
  }
}

List xlsxvalidation::information() {
  // Returns a nested data frame of everything, the data frame itself wrapped in
  // a list.
  List out = List::create(
      _["sheet"] = sheet_,
      _["ref"] = ref_,
      _["type"] = type_,
      _["operator"] = operator_,
      _["formula1"] = formula1_,
      _["formula2"] = formula2_,
      _["allow_blank"] = allow_blank_,
      _["show_input_message"] = show_input_message_,
      _["prompt_title"] = prompt_title_,
      _["prompt_body"] = prompt_,
      _["show_error_message"] = show_error_message_,
      _["error_title"] = error_title_,
      _["error_body"] = error_,
      _["error_symbol"] = error_style_);

  // Turn list of vectors into a data frame without checking anything
  int n = Rf_length(out[0]);
  out.attr("class") = CharacterVector::create("tbl_df", "tbl", "data.frame");
  out.attr("row.names") = IntegerVector::create(NA_INTEGER, -n); // Dunno how this works (the -n part)

  return out;
}
