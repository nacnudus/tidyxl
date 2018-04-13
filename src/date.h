#ifndef TIDYXL_DATE_
#define TIDYXL_DATE_

#include <Rcpp.h>

// Follow tidyverse/readxl
// https://github.com/tidyverse/readxl/commit/63ef215f57322dd5d7a27799a2a3fe463bd39fc7
// in correcting subsecond rounding by converting serial date to
// decimilliseconds, using a well-known hack to round to nearest integer w/o
// C++11 or boost, e.g.
// http://stackoverflow.com/questions/485525/round-for-float-in-c, and
// converting back to serial date.
inline double dateRound(double date) {
  double ms = date * 10000;
  ms = (ms >= 0.0 ? std::floor(ms + 0.5) : std::ceil(ms - 0.5));
  return ms / 10000;
}

// Follow tidyverse/readxl
// https://github.com/tidyverse/readxl/commit/c9a54ae9ce0394808f6d22e8ef1a7a647b2d92bb
// by correcting for Excel's faithful re-implementation of the Lotus 1-2-3 bug,
// in which February 29, 1900 is included in the date system, even though 1900
// was not actually a leap year
// https://support.microsoft.com/en-us/help/214326/excel-incorrectly-assumes-that-the-year-1900-is-a-leap-year
// How we address this: If date is *prior* to the non-existent leap day: add a
// day If date is on the non-existent leap day: make negative and, in due
// course, NA Otherwise: do nothing
inline double checkDate(double& date, int& dateSystem, int& dateOffset, std::string ref) {
  if (dateSystem == 1900 && date < 61) {
    date = (date < 60) ? date + 1 : -1;
  }
  if (date < 0) {
    Rcpp::warning("NA inserted for impossible 1900-02-29 datetime: " + ref);
    return NA_REAL;
  } else {
    return dateRound((date - dateOffset) * 86400);
  }
}

// Convert datetime doubles to strings "%Y-%m-%d %H:%M:%S"
// TODO: Support subseconds
inline std::string formatDate(double& date, int& dateSystem, int& dateOffset) {
  const char* format;
  if (date < 1) {                      // time only (in Excel's formuat)
    format = "%H:%M:%S";
  } else {                             // date and maybe time
    /* double intpart; */
    /* if (modf(date, &intpart) == 0.0) { // date only, not time */
    /*   format = "%Y-%m-%d"; */
    /* } else {                           // date and time */
      format = "%Y-%m-%d %H:%M:%S";
    /* } */
  }
  /* // This doesn't work on Windows before 1970, so use R functions instead */
  /* char out[charsize]; */
  /* date = checkDate(date, dateSystem, dateOffset, ""); // Convert to POSIX */
  /* std::time_t date_time_t(date); */
  /* std::tm date_tm = *std::gmtime(&date_time_t); */
  /* std::strftime(out, sizeof(out), format, &date_tm); */
  // Use R functions instead of C++ strftime, which doesn't work before 1970 on
  // Windows
  date = checkDate(date, dateSystem, dateOffset, ""); // Convert to POSIX
  std::string out;
  Rcpp::Function r_as_POSIXlt("as.POSIXlt");
  Rcpp::Function r_format_POSIXlt("format.POSIXlt");
  out =  Rcpp::as<std::string>(r_format_POSIXlt(r_as_POSIXlt(date,
                                                             "GMT",
                                                             "1970-01-01"),
                                                format,
                                                false));
  return out;
}

// Adapted from @reviewher https://github.com/tidyverse/readxl/issues/388
// See also ECMA Part 1 page 1785 (actual page 1795) section 18.8.31 "numFmts
// (Number Formats)"
#define CASEI(c) case c: case (c | 0x20)
#define CMPLC(j,n) if(x[i+j] | (0x20 == n))
inline bool isDateFormat(std::string x) {
  char escaped = 0;
  char bracket = 0;
  for (size_t i = 0; i < x.size(); ++i) switch (x[i]) {
    CASEI('D'):
    CASEI('E'):
    CASEI('H'):
    CASEI('M'):
    CASEI('S'):
    CASEI('Y'):
      if(!escaped && !bracket) return true;
      break;
    case '"':
      escaped = 1 - escaped; break;
    case '\\':
    case '_':
      ++i;
      break;
    case '[': if(!escaped) bracket = 1; break;
    case ']': if(!escaped) bracket = 0; break;
    CASEI('G'):
      if(i + 6 < x.size())
      CMPLC(1,'e')
      CMPLC(2,'n')
      CMPLC(3,'e')
      CMPLC(4,'r')
      CMPLC(5,'a')
      CMPLC(6,'l')
        return false;
  }
  return false;
}
#undef CMPLC
#undef CASEI

#endif
