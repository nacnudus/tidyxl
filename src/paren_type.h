#ifndef PAREN_TYPE_
#define PAREN_TYPE_

// Track whether we're inside function parentheses MAX(in here) or normal
// parentheses MAX(1+(in here)+{or in here}).
enum class paren_type {FUNCTION, PARENTHESES};

#endif
