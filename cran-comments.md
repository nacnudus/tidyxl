## Test environments

### Local
* Arch Linux 4.11.7-1-ARCH R-release 3.4.1

### Win-builder
* Windows Server 2008 (64-bit) R-devel 3.4.0 2017-07-05 r72891

### Travis
* Ubuntu Linux 12.04 LTS R-devel   2017-07-05 r72892
* Ubuntu Linux 12.04 LTS R-release 3.4.0
* Ubuntu Linux 12.04 LTS R-oldrel  3.3.3

### AppVeyor
* Windows Server 2012 R2 (x64) mingw_32     R-devel         2017-07-03 r72882
* Windows Server 2012 R2 (x64) x64 mingw_64 R-devel         2017-07-03 r72882
* Windows Server 2012 R2 (x64) x64 GCC      R-release 3.4.0
* Windows Server 2012 R2 (x64) i386 GCC     R-stable  3.4.0
* Windows Server 2012 R2 (x64) i386 GCC     R-patched 3.4.0 2017-06-17 r72808
* Windows Server 2012 R2 (x64) i386 GCC     R-oldrel  3.3.3

### Rhub
* Fedora Linux R-devel 2017-06-24 r72853 clang gfortran
* Debian Linux R-devel 2017-06-24 r72853 GCC ASAN/UBSAN

## R CMD check results
There were no ERRORs, WARNINGs, or NOTEs.

## Downstream dependencies
I have also run R CMD check on downstream dependencies of httr
(https://github.com/wch/checkresults/blob/master/httr/r-release).
All packages passed.
