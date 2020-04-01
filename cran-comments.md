## Resubmission

This is a resubmission. In this version I have:

* Fixed broken URIs in the documentation.

## Test environments

### Local
* Arch Linux 4.19.8-arch1-1-ARCH     R-release 3.6.2

### Win-builder
* Windows Server 2008 (64-bit)       R-devel   2020-03-11 r77925

### Travis
* Ubuntu Linux 16.04.6 LTS x64       R-devel   2020-03-13 r77948
* Ubuntu Linux 16.04.6 LTS x64       R-release 3.6.2
* Ubuntu Linux 16.04.6 LTS x64       R-oldrel  3.5.3

### AppVeyor
* Windows Server 2012 R2 (x64) x64   R-release 3.6.3
* Windows Server 2012 R2 (x64) i386  R-stable  3.6.3
* Windows Server 2012 R2 (x64) i386  R-patched 2020-03-11 r78011
* Windows Server 2012 R2 (x64) i386  R-oldrel  3.5.3

### Rhub
* Ubuntu Linux 16.04 LTS GCC         R-release 3.6.1
* Windows Server 2008 R2 SP1         R-devel   2020-03-08 r77917
* Debian Linux, GCC ASAN/UBSAN       R-devel   2018-06-20 r74924

## R CMD check results
0 errors | 0 warnings | 2 notes

> * checking compiled code ... WARNING
> File ‘tidyxl/libs/tidyxl.so’:
>   Found ‘abort’, possibly from ‘abort’ (C)
>     Object: ‘xlex.o’

This is a false positive. There is no `abort`.  This warning only appears on
Fedora Linux with R-devel, clang and gfortran.

> * checking compilation flags used ... NOTE
>   Compilation used the following non-portable flag(s):
>     ‘-march=x86-64’

I believe this is a local issue on Arch Linux 4.19.8-arch1-1-ARCH R-release
3.6.2

> Found the following (possibly) invalid URLs:
>   URL: https://duncan-garmonsway.shinyapps.io/xlex/

From winbuilder. The URL is valid.

## Downstream dependencies

There is one downstream dependency, unpivotr, which passes.
