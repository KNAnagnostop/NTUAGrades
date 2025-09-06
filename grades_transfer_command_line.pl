#!/usr/bin/perl
#
#---------------------------------------------------------------------
use utf8;use open qw(:std :utf8); use Encode qw(decode);#use Unicode::Collate; # for unicode string comparison
#---------------------------------------------------------------------
# Defaults:
$ixlsx="students.xlsx"; # xlsx containing the gradet, use -i to change
$bxlsx="Grammateia/ΒΑΘΜΟΛΟΓΙΟ ΠΡΟΓΡΑΜΜΑΤΙΣΜΟΣ ΜΕ ΕΦΑΡΜΟΓΕΣ ΣΤΗΝ ΕΠΙΣΤΗΜΗ ΤΟΥ ΜΗΧΑΝΙΚΟΥ (9348) 2024-2025 Εαρινή.xlsx"; # template xlsx file from the Ba0mologio https://e-sec.ntua.gr/
$colAM=5;               # Column of Student Number (starting from 1) , use -a option to change Note: Spreadsheet::ParseXLSX counts from 0, we reduce by one later
$colGR=8;               # Column of Grade          (starting from 1) , use -g option to change
#---------------------------------------------------------------------
# Probably never change:
$colAO=1;               # Column of Student Number in Ba0mologio (starting from 1), use the -A option to change
$colGO=7;               # Column of Grade in Ba0mologio (starting from 1), use the -o option to change
$GLOG = *STDERR;        # will change within function transfer_grades()
$DEBUG= 0;
#---------------------------------------------------------------------
use Spreadsheet::ParseXLSX;       # wrapper: see also perldoc Spreadsheet::ParseExcel
use Spreadsheet::ParseExcel::Cell;# more ways to read a cell, like unformatted()
use Excel::Writer::XLSX;          # use Spreadsheet::WriteExcel;
use Getopt::Std; use Switch 'Perl6'; use Sys::Hostname; use Time::HiRes qw(usleep); use File::Basename; use File::Copy; use File::Find; use Cwd;use File::Path;use Time::Piece;
# --------------------- OPTIONS       --------------------------------
$cdir = cwd;chomp($cdir);$home=$ENV{$HOME};$host = hostname;
$prog = basename ($0   );
# Correctly decode command-line arguments
foreach my $arg (@ARGV) {$arg = decode('UTF-8', $arg);}
getopt('iagotb', \%opts); #here we add only options with args
foreach $opt (keys %opts){
 given ($opt){
  when /[i]/ {$ixlsx = $opts{$opt};last;}
  when /[t]/ {$bxlsx = $opts{$opt};last;}
  when /[a]/ {$colAM = $opts{$opt};last;}
  when /[g]/ {$colGR = $opts{$opt};last;}
  when /[o]/ {$colGO = $opts{$opt};last;}
  when /[d]/ {$DEBUG = 1;last;}
  when 'h' {usage();last;}
  default  {usage("Not a valid option: $opt");}
 }
}
# reduce by one, since columns are counted from 0 from here on:
$colAM--; $colGR--; $colGO--; $colAO--;

transfer_grades($ixlsx,$bxlsx,$colAM,$colGR,$colGO,$colAO);

# END of main program
#----------------------------------------------------------------------
# call: transfer_grades($ixlsx,$bxlsx,$colAM,$colGR,$colGO,$colAO);
sub transfer_grades{
 my $ixlsx = shift;     # local variables to make the function independent of main program
 my $bxlsx = shift;
 my $colAM = shift;
 my $colGR = shift;
 my $colGO = shift;
 my $colAO = shift;
 my %grades;
# -------------------------------------------------------------------
# check validity+existence of filenames:
 my @fnames = ($ixlsx,$bxlsx); foreach my $fname (@fnames){check_valid_filename($fname);}

# -------------------------------------------------------------------
# create filenames:
 my ($fnixlsx,$dirixlsx,$extixlsx) = fileparse($ixlsx,qr/\.[^.]*$/);
 $oxlsx=$fnixlsx."-ba0mologio.xlsx";
 $log  =$fnixlsx."_log.txt";
# -------------------------------------------------------------------
# start computing:
 open my $LOG, ">", $log || usage("Failed to open $log");     $GLOG=$LOG;
 print $LOG "Date: ".localtime->strftime("%y%m%d%H%M%S");
 print $LOG "\n\nTransfering grades:\nFrom: $ixlsx\nTo: $oxlsx\nTemplate: $bxlsx\n";
 print $LOG "----------------------------------------------------------------------------------\n";

#---------------------------------------------------------------------
 print $LOG "Reading input grades from $ixlsx\n";
 print $LOG "----------------------------------------------------------------------------------\n";
 my $rparser       = Spreadsheet::ParseXLSX->new;
 my $rworkbook     = $rparser->parse($ixlsx);

 if( !defined $rworkbook){usage("Workbook failed to open from $ixlsx");}
 for $sheet  ($rworkbook->worksheets() ){                       # process all sheets
  $nsheets++;if($nsheets > 1 ){last;}                           # process only the 1st sheet
  $sname        = $sheet->get_name ();                    # perldoc Spreadsheet::ParseExcel::Worksheet :: the name of the current sheet

  ($rmin,$rmax) = $sheet->row_range(); if($rmax > 5000){$rmax=5000;}
  ($cmin,$cmax) = $sheet->col_range();

  debug("In new worksheet");
 
  $n = 0;
  for  $row  ($rmin .. $rmax){                            # loop over each row
   $cell = $sheet->get_cell( $row, $rmin );  
   if($sheet->get_cell( $row, $colAM )){ $AM = $sheet->get_cell( $row, $colAM)->unformatted();$AM =~ s/^\s+|\s+$//g;}else{$AM="";} # use get_value() to read as string, not as number
   if($sheet->get_cell( $row, $colGR )){ $GR = $sheet->get_cell( $row, $colGR)->value      ();$GR =~ s/^\s+|\s+$//g;}else{$GR="";}

#  debug("$row :: $AM :: $GR :: cAM= $colAM :: cGR= $colGR");

   if($AM !~ /^0[0-9][0-9][0-9][0-9][0-9][0-9][0-9]$/){next;}
#  if($GR eq "" || $GR < 0 || $GR > 10){next;} # allow floating point grades
   if($GR !~ /^\d+$/ || $GR < 0 || $GR > 10){next;}
   $grades{$AM} = $GR;
   $n++; debug( "${n}. $sname :: $row :: $AM :: $GR :: $colAM  :: $colGR :: grade: $grades{$AM}");

  }    # for  $row  ($rmin .. $rmax)
 }     # for $sheet  ($rworkbook->worksheets() )

#foreach my $ams (keys %grades){debug("KEYS: $row :: $ams :: $grades{$ams}");}; exit;

#---------------------------------------------------------------------
 print $LOG "Reading template $bxlsx\n";
 print $LOG "Output  in       $oxlsx\n";
 print $LOG "----------------------------------------------------------------------------------\n";

 my $tparser  = Spreadsheet::ParseXLSX->new;
 my $tworkbook = $tparser->parse($bxlsx);
 if(!defined $tworkbook){usage("Workbook failed to open from $bxlsx");}

 my $workbook_out = Excel::Writer::XLSX->new($oxlsx) || usage("Failed to create new xlsx file: $!");

 for my $tsheet ($tworkbook->worksheets()){
  my $worksheet_out = $workbook_out->add_worksheet($tsheet->get_name);
  my ($rmin, $rmax) = $tsheet->row_range();
  my ($cmin, $cmax) = $tsheet->col_range();

  my $width_default = 10;
  my @widths=(16, 40, 25, 25, 70, 25, 10); # Define an array for widths. Number of elements define merging of 1st row of output
  if($cmax < $#widths){$cmax = $#widths;}

  # Set column widths
  for my $col ($cmin .. $cmax){
   if ($col <= $#widths){
    $worksheet_out->set_column($col, $col, $widths[$col]);
   }else{
    $worksheet_out->set_column($col, $col, $width_default);
   }
  }

  #The first row is the title:
  my $title_cell = $tsheet->get_cell(0, 0);
  my $title      = (defined $title_cell) ? $title_cell->value() : "";
  # Merge cells in row 1 (row index 0) and add the title
  my $mergeformat = $workbook_out->add_format(align=>'center'); # add_format(border => 6, valign => 'vcenter', align  => 'center')
  $worksheet_out->merge_range(0, 0, 0, $#widths, $title,$mergeformat);

  $rmin = 1; # skip 1st line
  for my $row ($rmin .. $rmax){
   my $cell_am = $tsheet->get_cell($row, $colAO); # Assuming student ID is in column $colAO
   my $student_id = defined $cell_am ? $cell_am->unformatted() : "";

   for my $col ($cmin .. $cmax){
    my $value;
    my $cell = $tsheet->get_cell($row, $col);

    #next unless defined $cell;
    if(defined $cell){$value = $cell->value();}else{$value = "";}
       
    debug("XXX $row :: $col :: $colGO :: $student_id :: $grades{$student_id} :: cmin= $cmin  cmax= $cmax  col= $col");
    # Check if we have a grade for this student and if the current column is $colGO
    if($col == $colAO){
     $worksheet_out->write_string($row, $col, $student_id);   # write $AM as a string, not number
    } 
    elsif ($col == $colGO && exists $grades{$student_id}) {
     $worksheet_out->write($row, $col, $grades{$student_id});
    }  else{
     # Copy all other cells as-is
     $worksheet_out->write($row, $col, $value);
    }    # if ($col == 6 && exists $grades{$student_id}) 
   }     # for my $col    ($cmin .. $cmax)
  }      # for my $row    ($rmin .. $rmax)
 }       # for my $tsheet ($tworkbook->worksheets())
 $workbook_out->close();
}        # sub transfer_grades
#-------------------------------------------------------------------
#check_valid_filename("Grammateia/ΒΑΘΜΟΛΟΓΙΟ ΠΡΟΓΡΑΜΜΑΤΙΣΜΟΣ ΜΕ ΕΦΑΡΜΟΓΕΣ ΣΤΗΝ ΕΠΙΣΤΗΜΗ ΤΟΥ ΜΗΧΑΝΙΚΟΥ (9348) 2024-2025 Εαρινή.xlsx");
#check_valid_filename("students.xlsx");
sub check_valid_filename(){

 my $fname  = shift;
 my $maxlen = 150;

 unless(-r     $fname         and -s $fname                                            ) {usage("$fname does not exist (or has zero size)");}
 unless( length($fname) <= $maxlen ){usage("$fname is too long (> $maxlen characters) length = ".length($fname));}

 my $regex = qr/[^\_\(\)\.\- \/a-zA-Z0-9\x{0370}-\x{03ff}]/;  # valid characters

 if ($fname =~ $regex) {
  my $invalid_char = $&; # The regex matches the first invalid character and captures it
  my $ordinal      = ord($invalid_char);
  usage("$fname does not have a valid name. Invalid character found: '$invalid_char' (U+$ordinal).");
 }

}

# ----------------------- HELP MESSAGE ------------------------------
sub usage(){
 if(@_){$message  = "Fatal Error: ".shift(@_)."\n";}
 $message .= << "EOF";
----------------------------------------------------------------------------------
Usage: ${prog} [options]
      -i <input>       xlsx file with grades in column determined by the -a option                 (default: $colAM) and student number in column determined by by the -g option (default: $colGR)
      -t <template>    xlsx file from e-sec.ntua.gr. Existing grades will be either replaced or left as is
      -a <col num>     Column number where the AM  student number is recorded                      (default: $colAM). Columns are counted from 1: Column A is 1, column B is 2 and so on.
      -g <col num>     Column number where the new student grade  is recorded                      (default: $colGR)
      -A <col num>     Column number where the AM  student number is recorded in the template file (default: $colAO)
      -o <col num>     Column number where the students grades are in the template file            (default: $colGO) - clean this column if you want to replace grades
      -d               Sets debugging mode

Columns in xlsx files are counted from 1 (column A is 1, column B is 2 and so on)
The script ADDS/REPLACES grades: If you have a preexisting template that already contains grades, it will add new grades from the input xlsx file, replace old grades, but will leave intact already existing grades.
----------------------------------------------------------------------------------
EOF
 print $GLOG $message;
 exit(1);
}
sub main::HELP_MESSAGE(){ usage();} #for --help (does not work when default?)
# -------------------------------------------------------------------
sub debug(){
 if( $DEBUG == 1 ){
  $message = shift;
  print $GLOG "DBG: $message\n";
 }
}
# -------------------------------------------------------------------
if($#ARGV < -1){usage();} #$ARGV[0-] arguments (not progname) $#ARGV=-1 (noarg)

# $f = /d/f.e => "e" = extension($f);"f"= filename($f);"/d"=dirname($f);"f.e"=basename($f);
sub extension(){($f,$d,$e)=fileparse(@_,qr/\.[^.]*$/);return $e}
sub filename (){($f,$d,$e)=fileparse(@_,qr/\.[^.]*$/);return $f}


#  ---------------------------------------------------------------------
#  Copyright by Konstantinos N. Anagnostopoulos (2025)
#  Physics Dept., National Technical University,
#  konstant@mail.ntua.gr, www.physics.ntua.gr/konstant
#  
#  This program is free software: you can redistribute it and/or modify
#  it under the terms of the GNU General Public License as published by
#  the Free Software Foundation, version 3 of the License.
#  
#  This program is distributed in the hope that it will be useful, but
#  WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
#  General Public License for more details.
#  
#  You should have received a copy of the GNU General Public Liense along
#  with this program.  If not, see <http://www.gnu.org/licenses/>.
#  -----------------------------------------------------------------------

