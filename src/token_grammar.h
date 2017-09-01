//[[Rcpp::depends(piton)]] // for pegtl

#ifndef TOKEN_GRAMMAR_
#define TOKEN_GRAMMAR_

#include <pegtl.hpp>
#include <Rcpp.h>
#include "paren_type.h"

using namespace tao::TAOCPP_PEGTL_NAMESPACE;

namespace xltoken
{
  // Generic rules
  struct OpenSquareParen : one< '[' > {};
  struct CloseSquareParen : one< ']' > {};
  struct colon : one< ':' > {};
  struct exclamation : one< '!' > {};
  struct dot : one< '.' > {};
  struct underscore : one< '_' > {};
  struct pipe : one< '|' > {};
  struct backslash : one< '\\' > {};
  struct question : one< '?' > {};
  struct QuoteD : one< '"' > {};
  struct QuoteS : one< '\'' > {};

  // Separators
  struct dollar: one< '$' > {};
  struct comma: one< ',' > {};
  struct semicolon : one< ';' > {};
  struct openparen: one< '(' > {};
  struct closeparen: one< ')' > {};
  struct OpenCurlyParen : one< '{' > {};
  struct CloseCurlyParen : one< '}' > {};

  // Operators
  struct plusop : one< '+' > {};
  struct minusop : one< '-' > {};
  struct mulop : one< '*' > {};
  struct divop : one< '/' > {};
  struct expop : one< '^' > {};
  struct concatop : one< '&' > {};
  struct intersectop : disable < space > {};
  struct rangeop : one< ':' > {};
  struct percentop : one< '%' > {};
  struct gtop : one< '>' > {};
  struct eqop : one< '=' > {};
  struct ltop : one< '<' > {};
  struct neqop : string< '<', '>' > {};
  struct gteop : string< '>', '=' > {};
  struct lteop : string< '<', '=' > {};

  // The union operator ',' is omitted because it is ambiguous.  Usually a comma
  // represents a function separator, rather than a union operator.  XLParser
  // assumes that it is a function separator unless the whole expression is
  // within parentheses, e.g. "SMALL((A1,B2),1)", which seems to be how Excel
  // behaves.

  // The range operator ':' is usually handled inside the ref parser, where it
  // is used to interpret whole-row and whole-column ranges like AB:AC and
  // 12:13, but sometimes it is necessary to handle it separately when
  // range-returning functions like INDEX() are then ranged with something else,
  // e.g. INDEX(blahblahblah):D2

  // Text matches two QuoteD (") and anything between, i.e. character and
  // the surrounding pair of double-quotes.
  struct NotQuoteD : not_one< '"' > {};
  struct DoubleQuotedString : star< sor< seq< QuoteD, QuoteD >,
                                         NotQuoteD > >
  {};
  struct Text : seq< QuoteD, DoubleQuotedString, QuoteD > {};

  // Separators are characters that mark the end or beginning of more important
  // tokens.
  struct sep: sor< dollar,
                   comma,
                   semicolon,
                   openparen, closeparen,
                   OpenCurlyParen, CloseCurlyParen > {};
  struct notsep: if_then_else< at< disable< sep > >, failure, any > {};
  struct notseps: plus< notsep > {};

  // Bool is either TRUE or FALSE, followed by a separator
  struct Bool
    : seq< sor< string< 'T', 'R', 'U', 'E' >,
                string< 'F', 'A', 'L', 'S', 'E' > >,
           not_at< sor< alpha, disable< openparen > > > >
  {};

  // Error literal "#NULL!|#DIV/0!|#VALUE!|#NAME?|#NUM!|#N/A|#REF!"
  struct Error
    : sor< string< '#', 'D', 'I', 'V', '/', '0', '!' >,
           string< '#', 'N', '/', 'A' >,
           string< '#', 'N', 'A', 'M', 'E', '?' >,
           string< '#', 'N', 'U', 'L', 'L', '!' >,
           string< '#', 'N', 'U', 'M', '!' >,
           string< '#', 'R', 'E', 'F', '!' >,
           string< '#', 'V', 'A', 'L', 'U', 'E', '!' > >
  {};


  // Number matches a straightforward decimal number, including exponents.
  struct plusminus : sor< plusop, minusop >  {};
  template< typename D >
    struct decimal : if_then_else< dot,
                                   plus< D >,
                                   seq< plus< D >, opt< dot, plus< D > > > > {};
  struct e : one< 'e', 'E' > {};
  struct exponent : seq< opt< plusminus >, plus< digit > > {};
  struct Number
    : seq< opt< plusminus >,
           decimal< digit >,
           opt< e, exponent > >
  {};

  struct Function : seq< opt< string< '_', 'x', 'l', 'l', '.' > >,
                         plus< sor< alnum, underscore, dot > >,
                         disable< openparen > >
  {};

  struct Operator : sor< plusop,
                         minusop,
                         mulop,
                         divop,
                         expop,
                         concatop,
                         intersectop,
                         rangeop,
                         percentop,
                         eqop,
                         neqop, // Must precede lteop and ltop
                         gteop, // Must precede gtop
                         lteop, // Must precede ltop
                         gtop,
                         ltop >
  {};

  // Name as in named formula, as well as worksheet names
  // Start with a letter or underscore, continue with word character (letters,
  // numbers and underscore), dot or question mark
  // * first character: [\p{L}\\_]
  // * subsequent characters: [\w\\_\.\?]
  struct NameStartCharacter : sor< alpha, underscore, backslash > {};
  struct NameValidCharacter
    : sor< NameStartCharacter, digit, dot, question >
  {};
  struct Name : seq< NameStartCharacter, star< NameValidCharacter > > {};

  // Sheets and files
  // This doesn't check that a sheet is followed by a ref or a name.
  // Nor does it separate multiple sheets (ones separated by colons
  // Sheet1:Sheet2).  Finally, it includes any file indexes and doesn't
  // dereference them.  Single-quoted sheet names match two QuoteS (') and
  // anything between, i.e.  character and the surrounding pair of
  // single-quotes.  Any name, quoted or unquoted, must be suffixed by an
  // exclamation mark.
  struct UnquotedName : disable< Name > {};
  struct NotQuoteS : not_one< '\'' > {};
  struct SingleQuotedString : seq< QuoteS,
                                   star< sor< seq< QuoteS, QuoteS >,
                                              NotQuoteS > >,
                                   QuoteS >
  {};
  // Files are normalized to indexes like [0] which appear inside the quotes of
  // single-quoted sheet names, or outside non-quoted sheet names.
  struct File : seq< OpenSquareParen, plus< digit >, CloseSquareParen > {};
  struct Sheet : seq< sor< seq< opt< File >,
                                list< UnquotedName, colon > >,
                           SingleQuotedString >,
                      exclamation > {};

  // DDE (Dynamic Data Exchange) -- differs from XLParser, which does not demand
  // a pipe.
  struct DDEName : sor< UnquotedName, SingleQuotedString > {};
  struct DynamicDataExchange :
    sor< // Un-normalized form
         seq< DDEName,     // Program
              pipe,
              DDEName,     // Command
              exclamation,
              DDEName >,   // Parameters
         // Normalized form
         seq< File,        // Index of program and command
              exclamation,
              DDEName > >    // Parameters
  {};

  // Structured references
  struct EnclosedInBrackets
    : seq< OpenSquareParen,
           plus< not_one< '[', ']' > >,
           CloseSquareParen >
  {};
  struct StructuredReferenceExpression
    : seq< EnclosedInBrackets,
           opt< seq< sor< colon, disable< comma > >, EnclosedInBrackets > >,
           opt< seq< sor< colon, disable< comma > >, EnclosedInBrackets > >,
           opt< seq< colon, EnclosedInBrackets > > >
  {};
  struct StructuredReference
    : seq< sor< seq< OpenSquareParen, EnclosedInBrackets, CloseSquareParen >, // Not seen in the wild
                EnclosedInBrackets,
                seq< disable< Name >,
                     sor< EnclosedInBrackets,
                          seq< OpenSquareParen,
                               opt< StructuredReferenceExpression >,
                               CloseSquareParen > > > >,
            not_at< exclamation > >
  {};

  // After attempting a Ref, attempt anything else up to the next dollar, comma
  // or parentheses, which are characters that separate other tokens.
  struct Other: sor< sep, notseps > {};
  struct NotRef : sor< Text,
                       Sheet,
                       StructuredReference,
                       DynamicDataExchange,
                       Bool,
                       Error,
                       Function,
                       Name, // must come after Function, because it doesn't demand an exclamation
                       // To handle -1+=1, where some +- are operators, others are part of numbers
                       if_must< at< disable< seq< sor< bof,
                                                       sep,
                                                       rep< 2, plusminus > >,
                                                  Number > > >,
                                seq< sor< bof,
                                          sep,
                                          Operator >,
                                     Number > >,
                       Operator, // must precede Number otherwise 1-1 is interpreted as 1 and -1, not as an operation.
                       Number,
                       Other > {};

  // Refs
  // Anything above 1048576 is not a valid column
  struct MaybeRowToken : rep_min_max< 1, 7, digit > {};
  struct BadRowToken : seq< range< '1', '9' >,
                            range< '0', '9' >,
                            range< '4', '9' >,
                            range< '8', '9' >,
                            range< '5', '9' >,
                            range< '7', '9' >,
                            range< '7', '9' > >
  {};
  struct RowToken : seq< not_at< BadRowToken >, MaybeRowToken > {};
  // Anything above XFD is not a valid column
  struct MaybeColToken : rep_min_max< 1, 3, upper > {};
  struct BadColToken : seq< range< 'X', 'Z' >,
                            range< 'F', 'Z' >,
                            range< 'E', 'Z' > >
  {};
  struct ColToken : seq< not_at< BadColToken >, MaybeColToken > {};
  struct OptDollar : opt< disable< dollar > > {};
  struct OptRowToken : seq< OptDollar, RowToken > {};
  struct OptColToken : seq< OptDollar, ColToken > {};
  // Attempt to match addresses in this order
  // A:A
  // A1
  // A1:A2
  // 1:1
  struct Ref :
    seq< OptDollar,
         sor< seq< ColToken,
                   if_then_else< colon,
                                 OptColToken,                    // A:A
                                 seq< OptRowToken,               // A1
                                      opt< colon,
                                           OptColToken,
                                           OptRowToken  > > > >, // A1:A2
              seq< RowToken,
                   colon,
                   OptRowToken > >,
         not_at< sor< NameValidCharacter,        // not e.g. A1A or E09904.2!A1
                      disable< openparen > > > > // not e.g. LOG10()
  {};

  // Overall parsing rule.  The star< space > is to ignore meaningless spaces.
  struct root : seq< star< space >,
                     opt< Ref >,
                     star< space >,
                     star< seq< NotRef, star< space > >,
                           opt< seq< Ref, star< space > > > >,
                     star< space > > {};

  // Class template for user-defined actions that does
  // nothing by default.
  template<typename Rule>
    struct tokenize : nothing<Rule> {};

  // Specialisations of the user-defined action to do something when a rule
  // succeeds; is called with the portion of the input that matched the rule.

  template<> struct tokenize< Ref >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        levels.push_back(level);
        types.push_back("ref");
        tokens.push_back(in.string());
      }
  };

  template<> struct tokenize< Sheet >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        levels.push_back(level);
        types.push_back("sheet");
        tokens.push_back(in.string());
      }
  };

  template<> struct tokenize< Name >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        levels.push_back(level);
        types.push_back("name");
        tokens.push_back(in.string());
      }
  };

  template<> struct tokenize< DynamicDataExchange >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        levels.push_back(level);
        types.push_back("DDE");
        tokens.push_back(in.string());
      }
  };

  template<> struct tokenize< StructuredReference >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        levels.push_back(level);
        types.push_back("structured_ref");
        tokens.push_back(in.string());
      }
  };


  template<> struct tokenize< Text >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        levels.push_back(level);
        types.push_back("text");
        tokens.push_back(in.string());
      }
  };

  template<> struct tokenize< notseps >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        levels.push_back(level);
        types.push_back("other");
        tokens.push_back(in.string());
      }
  };

  template<> struct tokenize< comma >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        levels.push_back(level);
        tokens.push_back(in.string());
        switch (fun_paren.back()) {
          case paren_type::PARENTHESES:
            types.push_back("operator");
            break;

          case paren_type::FUNCTION:
            types.push_back("separator");
            break;
        }
      }
  };

  template<> struct tokenize< semicolon >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        levels.push_back(level);
        types.push_back("separator");
        tokens.push_back(in.string());
      }
  };

  template<> struct tokenize< openparen >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        levels.push_back(level);
        types.push_back("paren_open");
        tokens.push_back(in.string());
        ++level;
        fun_paren.push_back(paren_type::PARENTHESES);
      }
  };

  template<> struct tokenize< closeparen >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        level--;
        levels.push_back(level);
        tokens.push_back(in.string());
        switch (fun_paren.back()) {
          case paren_type::PARENTHESES:
            types.push_back("paren_close");
            break;

          case paren_type::FUNCTION:
            types.push_back("fun_close");
            break;
        }
        fun_paren.pop_back();
      }
  };

  template<> struct tokenize< OpenCurlyParen >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        levels.push_back(level);
        types.push_back("open_array");
        tokens.push_back(in.string());
        level++;
        fun_paren.push_back(paren_type::FUNCTION);
      }
  };

  template<> struct tokenize< CloseCurlyParen >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        level--;
        levels.push_back(level);
        types.push_back("close_array");
        tokens.push_back(in.string());
        fun_paren.pop_back();
      }
  };

  template<> struct tokenize< Bool >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        levels.push_back(level);
        types.push_back("bool");
        tokens.push_back(in.string());
      }
  };

  template<> struct tokenize< Error >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        levels.push_back(level);
        types.push_back("error");
        tokens.push_back(in.string());
      }
  };

  template<> struct tokenize< Number >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        levels.push_back(level);
        types.push_back("number");
        tokens.push_back(in.string());
      }
  };

  template<> struct tokenize< Function >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        // Handle the function name
        levels.push_back(level);
        types.push_back("function");
        std::string in_string(in.string());
        in_string.pop_back();               // remove the terminal '('
        tokens.push_back(in_string);

        // Handle the open parenthesis
        levels.push_back(level);
        types.push_back("fun_open");
        tokens.push_back("(");

        // Handle the new level and context
        ++level;
        fun_paren.push_back(paren_type::FUNCTION);
      }
  };

  template<> struct tokenize< Operator >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        levels.push_back(level);
        types.push_back("operator");
        tokens.push_back(in.string());
      }
  };

  template<> struct tokenize< space >
  {
    template< typename Input >
      static void apply( const Input & in,
                         int & level,
                         std::vector<int> & levels,
                         std::vector<paren_type> & fun_paren,
                         std::vector<std::string> & types,
                         std::vector<std::string> & tokens)
      {
        levels.push_back(level);
        types.push_back("space");
        tokens.push_back(in.string());
      }
  };

} // xltoken

#endif
