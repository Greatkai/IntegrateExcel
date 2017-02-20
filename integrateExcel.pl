#!/usr/bin/env perl

use File::Basename;
use Data::Dumper;
use Spreadsheet::XLSX;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtUnicode;
use Excel::Writer::XLSX;
use MyExcelFormatter;
use Text::Iconv;
use Encode;
use Unicode::Map;
use utf8;

$currentDir = dirname(__FILE__);
$excelPath = $currentDir . '/PutExcelsHere';

#my $converter = Text::Iconv -> new ("gb2312", "gbk");
#my $fmt = new MyExcelFormatter();

my $workbook = Excel::Writer::XLSX->new( $currentDir . '/iResults.xlsx' );
# Add a worksheet
$worksheet = $workbook->add_worksheet();
#  Add and define a format
$format = $workbook->add_format();
$format->set_bold();
$format->set_color( 'red' );
$format->set_align( 'center' );

my $workRow = 0;
my $firstFile = 1;
my $targetDir = "$excelPath";
my @fileDirs;
get_dir_file($targetDir);
foreach $file (@fileDirs){
  if($file =~ /\.xlsx$/){
    if($firstFile == 1){
      $minRow = 0;
    }else{
      $minRow = 1;
    }

     my $excel = Spreadsheet::XLSX -> new ($file);
     @sheets = @{$excel -> {Worksheet}};
     $sheet = $sheets[0];
     printf("Sheet: %s\n", $sheet->{Name});
     $sheet -> {MaxCol} ||= $sheet -> {MinCol};
     foreach $row ($minRow .. 1){
       foreach my $col ($sheet -> {MinCol} ..  $sheet -> {MaxCol}) {
         my $cell = $sheet -> {Cells} [$row] [$col];
         if ($cell) {
          # print Dumper($cell);
           $a = $cell -> {Val};
           if($a =~ /^([^a-z']+)/){
             $a = $1;
           }
            #printf("( %s , %s ) => %s\n", $row, $col, $a);

            $worksheet->write_string( $workRow, $col, encode("unicode", decode("utf8", $a)));
          }
        }
        $workRow++;
     }
    $firstFile = 0;
  }

  if($file =~ /\.xls$/){

    if($firstFile == 1){
      $minRow = 0;
    }else{
      $minRow = 1;
    }

    my $oExcel = new Spreadsheet::ParseExcel;
    my $oBook = $oExcel->Parse( $file );
    if ( !defined $oBook ) {
      die $parser->error(), ".\n";
    }

    @sheets = @{$oBook->{Worksheet}};
    $oWkS = $sheets[0];

    foreach $iR ($minRow .. 1){
      for( $iC = $oWkS->{MinCol} ; defined $oWkS->{MaxCol} && $iC <= $oWkS->{MaxCol} ; $iC++) {
          $oWkC = $oWkS->{Cells}[$iR][$iC];
          $a = $oWkC -> {_Value};
          #print Dumper($oWkC)."\n";
          $worksheet->write_string( $workRow, $iC, encode('unicode',$a));
      }
      $workRow++;
    }
    $firstFile = 0;
  }
}

sub get_dir_file
{
    my $path = shift @_;
#    print $path;
    opendir(TEMP, $path) || die "open $path fail...$!";
    my @FILES = readdir TEMP;
    for my $filename (@FILES) {
        if ($filename eq "Thumbs.db" || $filename eq "." || $filename eq ".." || $filename eq ".svn" || $filename eq "get_file_list.pl")
        {
        }
        else
        {
            if (-d "$path/$filename") {
                #print "$path/$filename"."\n";
                get_dir_file("$path/$filename");
            }
            else {
                write_to_file("$path/$filename");
            }
        }
    }
    closedir(TEMP);
}

sub write_to_file
{
    my $text = shift @_;
    #print "$text\n";
    push @fileDirs,$text;
}
