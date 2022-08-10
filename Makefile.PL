use strict;
use warnings;

use 5.009_001;
use ExtUtils::MakeMaker;

WriteMakefile(
  NAME             => 'xls2xlsx',
  AUTHOR           => 'David Betz <hashref@gmail.com>',
  VERSION_FROM     => 'script/xls2xlsx',
  ABSTRACT         => 'A quick and dirty .xls to .xlsx file converter.',
  LICENSE          => 'perl',
  MIN_PERL_VERSION => '5.9.1',
  PREREQ_PM        => {
    'Term::Emit'              => '0.0.4',
    'Spreadsheet::ParseExcel' => '0.65',
    'Spreadsheet::ParseExcel' => '0.65',
  },
  TEST_REQUIRES => {
    "Test2::V0" => "0.000060",
    "Test::Script" => "1.29",
  },
  EXE_FILES     => ['script/xls2xlsx'],
  META_MERGE    => {
    'meta-spec' => { version => 2 },
    resources   => {
      repository => {
        type => 'git',
        url  => 'https://github.com/hashref/xls2xlsx.git',
        web  => 'https://github.com/hashref/xls2xlsx',
      },
      bugtracker => { web => 'https://github.com/hashref/xls2xlsx/issues' },
    },
  },
);
