#!/usr/bin/env perl

use strict;
use warnings;

use Test2::V0;
use Test::Script;

my $SCRIPT_PATH   = './script/xls2xlsx';
my $XLS_FILE_PATH = './t/test_data/test.xls';
( my $XLSX_FILE_PATH        = $XLS_FILE_PATH )  =~ s/$/x/;
( my $XLSX_FILE_PATH_RENAME = $XLSX_FILE_PATH ) =~ s/test\./test2\./;

plan(7);

script_compiles( $SCRIPT_PATH, 'script compiles' );

script_fails $SCRIPT_PATH , { exit => 1 }, 'execution fails with missing CL arguments';

# Test Command Line Options
script_runs( [ $SCRIPT_PATH, '-h' ], '-h option works' );
script_runs( [ $SCRIPT_PATH, '-m' ], '-m option works' );
`$SCRIPT_PATH -o $XLSX_FILE_PATH_RENAME $XLS_FILE_PATH`;
( my $output_file_test = `ls -1 $XLSX_FILE_PATH_RENAME` ) =~ s/\n$//;
if ( $output_file_test eq $XLSX_FILE_PATH_RENAME ) {
  pass('-o options works');
}
else {
  fail('-o option works');
}

# Test File Conversion
script_runs( [ $SCRIPT_PATH, $XLS_FILE_PATH ], 'execute with xls file' );
( my $file_test = `file -b $XLSX_FILE_PATH` ) =~ s/\n$//;
if ( $file_test eq 'Microsoft Excel 2007+' ) {
  pass('test convertion');
}
else {
  fail('test conversion');
}

done_testing();

END {
  # Cleanup
  unlink $XLSX_FILE_PATH, $XLSX_FILE_PATH_RENAME;
}
