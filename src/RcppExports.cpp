// Generated by using Rcpp::compileAttributes() -> do not edit by hand
// Generator token: 10BE3573-1514-4C36-9D1C-5A225CD40393

#include <Rcpp.h>

using namespace Rcpp;

// xlsx_cells_
List xlsx_cells_(std::string path, CharacterVector sheet_paths, CharacterVector sheet_names, CharacterVector comments_paths);
RcppExport SEXP _tidyxl_xlsx_cells_(SEXP pathSEXP, SEXP sheet_pathsSEXP, SEXP sheet_namesSEXP, SEXP comments_pathsSEXP) {
BEGIN_RCPP
    Rcpp::RObject rcpp_result_gen;
    Rcpp::RNGScope rcpp_rngScope_gen;
    Rcpp::traits::input_parameter< std::string >::type path(pathSEXP);
    Rcpp::traits::input_parameter< CharacterVector >::type sheet_paths(sheet_pathsSEXP);
    Rcpp::traits::input_parameter< CharacterVector >::type sheet_names(sheet_namesSEXP);
    Rcpp::traits::input_parameter< CharacterVector >::type comments_paths(comments_pathsSEXP);
    rcpp_result_gen = Rcpp::wrap(xlsx_cells_(path, sheet_paths, sheet_names, comments_paths));
    return rcpp_result_gen;
END_RCPP
}
// xlsx_formats_
List xlsx_formats_(std::string path);
RcppExport SEXP _tidyxl_xlsx_formats_(SEXP pathSEXP) {
BEGIN_RCPP
    Rcpp::RObject rcpp_result_gen;
    Rcpp::RNGScope rcpp_rngScope_gen;
    Rcpp::traits::input_parameter< std::string >::type path(pathSEXP);
    rcpp_result_gen = Rcpp::wrap(xlsx_formats_(path));
    return rcpp_result_gen;
END_RCPP
}
// xlsx_sheet_files_
DataFrame xlsx_sheet_files_(std::string path);
RcppExport SEXP _tidyxl_xlsx_sheet_files_(SEXP pathSEXP) {
BEGIN_RCPP
    Rcpp::RObject rcpp_result_gen;
    Rcpp::RNGScope rcpp_rngScope_gen;
    Rcpp::traits::input_parameter< std::string >::type path(pathSEXP);
    rcpp_result_gen = Rcpp::wrap(xlsx_sheet_files_(path));
    return rcpp_result_gen;
END_RCPP
}
// xlsx_validation_
List xlsx_validation_(std::string path, CharacterVector sheet_paths, CharacterVector sheet_names);
RcppExport SEXP _tidyxl_xlsx_validation_(SEXP pathSEXP, SEXP sheet_pathsSEXP, SEXP sheet_namesSEXP) {
BEGIN_RCPP
    Rcpp::RObject rcpp_result_gen;
    Rcpp::RNGScope rcpp_rngScope_gen;
    Rcpp::traits::input_parameter< std::string >::type path(pathSEXP);
    Rcpp::traits::input_parameter< CharacterVector >::type sheet_paths(sheet_pathsSEXP);
    Rcpp::traits::input_parameter< CharacterVector >::type sheet_names(sheet_namesSEXP);
    rcpp_result_gen = Rcpp::wrap(xlsx_validation_(path, sheet_paths, sheet_names));
    return rcpp_result_gen;
END_RCPP
}
// xlsx_names_
List xlsx_names_(std::string path);
RcppExport SEXP _tidyxl_xlsx_names_(SEXP pathSEXP) {
BEGIN_RCPP
    Rcpp::RObject rcpp_result_gen;
    Rcpp::RNGScope rcpp_rngScope_gen;
    Rcpp::traits::input_parameter< std::string >::type path(pathSEXP);
    rcpp_result_gen = Rcpp::wrap(xlsx_names_(path));
    return rcpp_result_gen;
END_RCPP
}
// is_date_format_
LogicalVector is_date_format_(CharacterVector formats);
RcppExport SEXP _tidyxl_is_date_format_(SEXP formatsSEXP) {
BEGIN_RCPP
    Rcpp::RObject rcpp_result_gen;
    Rcpp::RNGScope rcpp_rngScope_gen;
    Rcpp::traits::input_parameter< CharacterVector >::type formats(formatsSEXP);
    rcpp_result_gen = Rcpp::wrap(is_date_format_(formats));
    return rcpp_result_gen;
END_RCPP
}
// xlex_
List xlex_(CharacterVector x);
RcppExport SEXP _tidyxl_xlex_(SEXP xSEXP) {
BEGIN_RCPP
    Rcpp::RObject rcpp_result_gen;
    Rcpp::RNGScope rcpp_rngScope_gen;
    Rcpp::traits::input_parameter< CharacterVector >::type x(xSEXP);
    rcpp_result_gen = Rcpp::wrap(xlex_(x));
    return rcpp_result_gen;
END_RCPP
}

static const R_CallMethodDef CallEntries[] = {
    {"_tidyxl_xlsx_cells_", (DL_FUNC) &_tidyxl_xlsx_cells_, 4},
    {"_tidyxl_xlsx_formats_", (DL_FUNC) &_tidyxl_xlsx_formats_, 1},
    {"_tidyxl_xlsx_sheet_files_", (DL_FUNC) &_tidyxl_xlsx_sheet_files_, 1},
    {"_tidyxl_xlsx_validation_", (DL_FUNC) &_tidyxl_xlsx_validation_, 3},
    {"_tidyxl_xlsx_names_", (DL_FUNC) &_tidyxl_xlsx_names_, 1},
    {"_tidyxl_is_date_format_", (DL_FUNC) &_tidyxl_is_date_format_, 1},
    {"_tidyxl_xlex_", (DL_FUNC) &_tidyxl_xlex_, 1},
    {NULL, NULL, 0}
};

RcppExport void R_init_tidyxl(DllInfo *dll) {
    R_registerRoutines(dll, NULL, CallEntries, NULL, NULL);
    R_useDynamicSymbols(dll, FALSE);
}
