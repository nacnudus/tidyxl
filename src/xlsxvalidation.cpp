#include <Rcpp.h>
#include "zip.h"
#include "rapidxml.h"
#include "xlsxvalidation.h"
#include "xlsxbook.h"
#include "utils.h"

using namespace Rcpp;

xlsxvalidation::xlsxvalidation(
    const std::string& sheet_path,
    xlsxbook& book):
  book_(book) {
  std::string sheet_ = zip_buffer(book_.path_, sheet_path);

  rapidxml::xml_document<> xml;
  xml.parse<0>(&sheet_[0]);

  rapidxml::xml_node<>* worksheet = xml.first_node("worksheet");
  rapidxml::xml_node<>* dataValidations = worksheet->first_node("dataValidations");


  if (dataValidations != NULL) {
    // Initialize the vectors to the number of rules, using the count attribute
    // if it exists, otherwise count the rules.
    int n(0);
    rapidxml::xml_attribute<>* count = dataValidations->first_attribute("count");
    if (count != NULL) {
      n = strtol(count->value(), NULL, 10);
    } else {
      for (rapidxml::xml_node<>* dataValidation = dataValidations->first_node("dataValidation");
          dataValidation; dataValidation = dataValidation->next_sibling("dataValidation")) {
          ++n;
      }
    }

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

    int i(0);
    for (rapidxml::xml_node<>* dataValidation = dataValidations->first_node("dataValidation");
        dataValidation; dataValidation = dataValidation->next_sibling("dataValidation")) {

      rapidxml::xml_attribute<>* sqref = dataValidation->first_attribute("sqref");
      if (sqref != NULL) {
        // Replace spaces with commas.  For space-delimited 'sqref' attributes,
        // which ought to be comma-delimited for consistency with the union operator
        // ',' in formulas.  Modifies in-place.
        std::string ref(sqref->value());
        std::replace(ref.begin(), ref.end(), ' ', ',');
        ref_[i] = ref;
      }

      std::string type_string;
      rapidxml::xml_attribute<>* type = dataValidation->first_attribute("type");
      if (type != NULL) {
        // operator defaults to 'between', but when type is one of list or
        // operator then it should be NA
        type_string = type->value();
        type_[i] = type_string;
        if (type_string == "list" || type_string == "custom")
          operator_[i] = NA_STRING;
      }

      rapidxml::xml_attribute<>* Operator = dataValidation->first_attribute("operator");
      if (Operator != NULL)
        operator_[i] = Operator->value();

      rapidxml::xml_node<>* formula1 = dataValidation->first_node("formula1");
      if (formula1 != NULL) {
        if (type_string == "date" || type_string == "time") {
          double date_double = strtod(formula1->value(), NULL);
          formula1_[i] = formatDate(date_double, book_.dateSystem_, book_.dateOffset_);
        } else {
          formula1_[i] = formula1->value();
        }
      }

      rapidxml::xml_node<>* formula2 = dataValidation->first_node("formula2");
      if (formula2 != NULL) {
        if (type_string == "date" || type_string == "time") {
          double date_double = strtod(formula2->value(), NULL);
          formula2_[i] = formatDate(date_double, book_.dateSystem_, book_.dateOffset_);
        } else {
          formula2_[i] = formula2->value();
        }
      }

      rapidxml::xml_attribute<>* allowBlank = dataValidation->first_attribute("allowBlank");
      if (allowBlank != NULL)
        allow_blank_[i] = strtod(allowBlank->value(), NULL);

      rapidxml::xml_attribute<>* showInputMessage = dataValidation->first_attribute("showInputMessage");
      if (showInputMessage != NULL)
        show_input_message_[i] = strtod(showInputMessage->value(), NULL);

      rapidxml::xml_attribute<>* promptTitle = dataValidation->first_attribute("promptTitle");
      if (promptTitle != NULL)
        prompt_title_[i] = promptTitle->value();

      rapidxml::xml_attribute<>* prompt = dataValidation->first_attribute("prompt");
      if (prompt != NULL)
        prompt_[i] = prompt->value();

      rapidxml::xml_attribute<>* showErrorMessage = dataValidation->first_attribute("showErrorMessage");
      if (showErrorMessage != NULL)
        show_error_message_[i] = strtod(showErrorMessage->value(), NULL);

      rapidxml::xml_attribute<>* errorTitle = dataValidation->first_attribute("errorTitle");
      if (errorTitle != NULL)
        error_title_[i] = errorTitle->value();

      rapidxml::xml_attribute<>* error = dataValidation->first_attribute("error");
      if (error != NULL)
        error_[i] = error->value();

      rapidxml::xml_attribute<>* errorStyle = dataValidation->first_attribute("errorStyle");
      if (errorStyle != NULL)
        error_style_[i] = errorStyle->value();

      ++i;
    }
  }
}

List& xlsxvalidation::information() {
  // Returns a nested data frame of everything, the data frame itself wrapped in
  // a list.
  information_ = List::create(
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
  int n = Rf_length(information_[0]);
  information_.attr("class") = Rcpp::CharacterVector::create("tbl_df", "tbl", "data.frame");
  information_.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, -n); // Dunno how this works (the -n part)

  return information_;
}
