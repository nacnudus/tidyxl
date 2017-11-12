## Test environments

### Local
* Arch Linux 4.13.11-1-ARCH              R-release 3.4.2

### Docker
* x86_64-pc-linux-gnu                    R-devel   2017-11-06 r73673
* x86_64-pc-linux-gnu                    R-release 3.4.2
* x86_64-pc-linux-gnu                    R-release 3.3.3

### Win-builder
* Windows Server 2008 (64-bit)           R-devel   2017-09-12 r73242

### Travis
* Ubuntu Linux 12.04 LTS                 R-release 3.4.2
* Ubuntu Linux 12.04 LTS                 R-oldrel  3.3.3

### AppVeyor
* Windows Server 2012 R2 (x64) x64 GCC   R-release 3.4.2
* Windows Server 2012 R2 (x64) i386 GCC  R-stable  3.4.2
* Windows Server 2012 R2 (x64) i386 GCC  R-patched 3.4.2 2017-11-09 r73698
* Windows Server 2012 R2 (x64) i386 GCC  R-oldrel  3.3.3

### Rhub
* Debian Linux R-devel 2017-07-26 r72972 GCC ASAN/UBSAN

## R CMD check results
0 errors | 0 warnings | 1 note

> checking installed package size ... NOTE
>   installed size is 10.0Mb
>   sub-directories of 1Mb or more:
>     libs   9.1Mb

That is all compiled code in the libs/ directory.

> Possibly mis-spelled words in DESCRIPTION:
>   Tokenizes (12:5)

'Tokenizes' is correct.
