# xls2xlsx

This is a quick and dirty .xls to .xlsx file converter. Don't expect much.

## Installation

Select the latest release `.tar.gz` file by clicking the latest project release. Copy the URL for the `.tar.gz` file and paste it into a `cpanm` command.

_Example:_

```sh
cpanm https://github.com/hashref/xls2xlsx/archive/refs/tags/v0.0.3.tar.gz
```

Or, if you are more familiar with Perl, you can clone this repo, run `perl Makefile.PL && make && make dist`. From there install the tarball that it creates with `cpanm`.

_Example:_

```sh
$ git clone https://github.com/hashref/xls2xlsx.git
$ cd xls2xlsx
$ perl Makefile.PL
$ make
$ make test
$ make dist
$ cpanm xls2xlsx-0.0.3.tar.gz
```

## Running

_Example:_

```sh
$ xlsx2xlsx ./file.xls
$ ls -l
file.xls
file.xlsx
```

Or, you can name the output file whatever you want...

_Example:_

```sh
$ xlsx2xlsx -o foobar.xlsx ./file.xls
$ ls -l
file.xls
foobar.xlsx
```
