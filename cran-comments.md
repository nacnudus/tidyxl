## Test environments

### Local
* Arch Linux 4.19.8-arch1-1-ARCH     R-release 3.5.1

### Win-builder
* Windows Server 2008 (64-bit)       R-devel   2018-12-28 r75917

### Travis
* Ubuntu Linux 14.04.5 LTS x64       R-devel   2018-12-30 r75926
* Ubuntu Linux 14.04.5 LTS x64       R-release 3.5.0
* Ubuntu Linux 14.04.5 LTS x64       R-oldrel  3.4.4

### AppVeyor
* Windows Server 2012 R2 (x64) x64   R-release 3.5.2
* Windows Server 2012 R2 (x64) i386  R-stable  3.5.2
* Windows Server 2012 R2 (x64) i386  R-patched 2018-12-29 r75924
* Windows Server 2012 R2 (x64) i386  R-oldrel  3.4.4

### Rhub
* Ubuntu Linux 16.04 LTS GCC         R-release 3.4.4
* Fedora Linux, clang, gfortran      R-devel   2018-12-22 r75884
* Windows Server 2008 R2 SP1         R-devel   2018-12-26 r75909
* Debian Linux, GCC ASAN/UBSAN       R-devel   2018-06-20 r74924

## R CMD check results
0 errors | 1 warnings | 2 notes

> * checking compiled code ... WARNING
> File ‘tidyxl/libs/tidyxl.so’:
>   Found ‘abort’, possibly from ‘abort’ (C)
>     Object: ‘xlex.o’

This is a false positive. There is no `abort`.  This warning only appears on
Fedora Linux with R-devel, clang and gfortran.

> Possibly mis-spelled words in DESCRIPTION:
>   Tokenizes (13:5)

Tokenizes is not a mis-spelled word.

> * checking installed package size ... NOTE
>   installed size is 11.8Mb
>   sub-directories of 1Mb or more:
>     libs  10.8Mb

The `libs` directory is compiled code.

## Downstream dependencies

There is one downstream dependency, unpivotr, which passes.
