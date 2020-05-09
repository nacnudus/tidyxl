## Test environments

### Local
* Arch Linux 4.19.8-arch1-1-ARCH     R-release 4.0.0 (2020-04-24)

### Win-builder
* Windows Server 2008 (64-bit)       R-devel   2020-05-08 r78391

### Travis
* Ubuntu Linux 16.04.6 LTS x64       R-devel   2020-05-08 r78391
* Ubuntu Linux 16.04.6 LTS x64       R-release 4.0.0 (2020-04-24)
* Ubuntu Linux 16.04.6 LTS x64       R-oldrel  3.6.3 (2017-01-27)

### AppVeyor
* Windows Server 2012 R2 (x64) x64   R-release 4.0.0 (2020-04-24)
* Windows Server 2012 R2 (x64) i386  R-stable  4.0.0 (2020-04-24)
* Windows Server 2012 R2 (x64) i386  R-patched 2020-05-08 r78391
* Windows Server 2012 R2 (x64) i386  R-oldrel  3.6.3 (2020-02-29)

### Rhub
* Ubuntu Linux 16.04 LTS GCC         R-release 3.6.1 (2019-07-05)
* Windows Server 2008 R2 SP1         R-devel   2020-04-22 r78281
* Fedora Linux, clang gfotran        R-devel   2020-05-07 r78381
* Debian Linux, GCC ASAN/UBSAN       R-devel   2018-06-20 r74924

## R CMD check results
0 errors | 2 warnings | 1 note

> * checking compiled code ... WARNING
> File ‘tidyxl/libs/tidyxl.so’:
>   Found ‘abort’, possibly from ‘abort’ (C)
>     Object: ‘xlex.o’

This is a false positive. There is no `abort`.  This warning only appears on
Fedora Linux with R-devel, clang and gfortran.

> Found the following (possibly) invalid URLs:
>   URL: https://duncan-garmonsway.shinyapps.io/xlex/

From win-builder. The URL is valid.

> * checking installed package size ... NOTE
>   installed size is 10.2Mb
>   sub-directories of 1Mb or more:
>     libs   9.2Mb

This is compiled C++ code.

## Downstream dependencies

There are downstream dependencies, unpivotr and unheadr, which pass.
