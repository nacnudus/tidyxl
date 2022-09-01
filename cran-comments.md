## Resubmission

This is a resubmission. In this version I have:

* Updated a URL that was being redirected.

## Test environments

### Local
* Arch Linux 5.15.63-1-lts, R-release 4.2.1 (2022-06-23)

### Rhub
* Windows Server 2022, R-devel, 64 bit
* Ubuntu Linux 20.04.1 LTS, R-release, GCC
* Fedora Linux, R-devel, clang, gfortran
* Debian Linux, R-devel, GCC ASAN/UBSAN

### Winbuilder
* R Under development (unstable) (2022-08-31 r82783 ucrt)

## R CMD check results
0 errors | 0 warnings | 1 note

> * checking installed package size ... NOTE
>  installed size is 23.5Mb
>  sub-directories of 1Mb or more:
>    libs  22.5Mb

This is compiled C++ code.

## Downstream dependencies

There are downstream dependencies, unpivotr and unheadr, which pass.
