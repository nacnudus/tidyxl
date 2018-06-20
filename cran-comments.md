## Test environments

### Local
* Arch Linux 4.16.9-1-ARCH               R-release 3.5.0

### Win-builder
* Windows Server 2008 (64-bit)           R-devel   2018-06-15 r74904
* Windows Server 2008 (64-bit)           R-release 3.5.0

### Travis
* Ubuntu Linux 12.04 LTS                 R-devel   2018-06-18 r74911
* Ubuntu Linux 12.04 LTS                 R-release 3.5.0
* Ubuntu Linux 12.04 LTS                 R-oldrel  3.4.4

### AppVeyor
* Windows Server 2012 R2 (x64) x64 GCC   R-release 3.5.0
* Windows Server 2012 R2 (x64) i386 GCC  R-stable  3.5.0
* Windows Server 2012 R2 (x64) i386 GCC  R-patched 3.5.0 2018-06-15 r74904
* Windows Server 2012 R2 (x64) i386 GCC  R-oldrel  3.4.4

### Rhub
* Ubuntu Linux 16.04 LTS GCC             R-release 3.4.4

## R CMD check results
0 errors | 0 warnings | 3 notes

> Found the following (possibly) invalid URLs:
>   URL: https://msdn.microsoft.com/en-us/library/dd922181(v=office.12).aspx
>     From: man/excel_functions.Rd
>     Status: Error
>     Message: libcurl error code 60:
>     	SSL certificate problem: unable to get local issuer certificate
>     	(Status without verification: OK)

The URL is valid.

> * checking installed package size ... NOTE
>   installed size is 11.6Mb
>   sub-directories of 1Mb or more:
>     libs  10.6Mb

That is all compiled code in the libs/ directory.

> Possibly mis-spelled words in DESCRIPTION:
>   Tokenizes (12:5)

'tokenizes' is correct.

## Downstream dependencies

There is one downstream dependency, unpivotr, which passes.
