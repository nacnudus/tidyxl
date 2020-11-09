## Dependence on {piton}

This update is for compatibility with the newest version of the package "piton",
which has been submitted to CRAN, but has not yet been accepted because it
breaks this package.

The required change in tidyxl is trivial (the name of a C++
namespace), so I believe that tidyxl will pass on all platforms.  I have
tested it locally on Linux and Mac.  Until piton is accepted, I can't test
tidyxl with winbuilder.

The update to piton fixes the apparent warning in tidyxl on some platforms:

> Version: 1.0.6
> Check: compiled code
> Result: WARN
>     File 'tidyxl/libs/tidyxl.so':
>      Found 'abort', possibly from 'abort' (C)
>      Object: 'xlex.o'

## Test environments

### Local
* Arch Linux 5.4.72-1-lts            R-release 4.0.3 (2020-10-10)
* MacOSX Darwin                      R-release 4.0.3 (2020-10-10)

## R CMD check results
0 errors | 0 warnings | 1 note

> ❯ checking compilation flags used ... NOTE
>   Compilation used the following non-portable flag(s):
>    ‘-march=x86-64’

I think this only occurs in the local development environment on Linux.  No such
flag is used in `Makevars` or `Makevars.win`.

## Downstream dependencies

There are downstream dependencies, unpivotr and unheadr, which pass.
