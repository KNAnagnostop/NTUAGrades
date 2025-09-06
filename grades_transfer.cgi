#!/usr/bin/perl
use strict;
use warnings;
use CGI;
use File::Spec;
use File::Basename;
use Time::Piece;
use Spreadsheet::ParseXLSX;
use Spreadsheet::ParseExcel::Cell;
use Excel::Writer::XLSX;
use File::Path   qw(mkpath);
use Sys::Hostname;
use URI::Escape;
use Scalar::Util qw(looks_like_number);
use Encode       qw(decode encode_utf8);
# ---------------------------------------------------------------------
# Define paths and URLs
my $pubdir = "/home/konstant/public_html/PUB/GradesForms";        # The path that contains the contents of $puburl
my $puburl = "https://physics.ntua.gr/konstant/PUB/GradesForms";  # constructs the URL that leads to the output files
# ---------------------------------------------------------------------
my $row_max = 5000; # we ignore lines above         this limit
my $col_max = 20;   # we ignore columns larger than this limit, should not be less than 8 since grades in Ba0mologia is in 7th column (your responsibility!)
# ---------------------------------------------------------------------
my $cgi = CGI->new;
my $log;
my $GLOG = *STDOUT;

# Set CGI headers and start HTML
print $cgi->header('text/html; charset=UTF-8');
#binmode STDOUT, ':utf8';
print <<EOF;
<!DOCTYPE html>
<html lang="el">
<head>
    <meta charset="UTF-8">
    <title>Grade Transfer Results</title>
    <style>
        body { font-family: sans-serif; max-width: 800px; margin: auto; padding: 20px; line-height: 1.6; }
        .success      { color: green;     font-weight: bold; }
        .error        { color: red;       font-weight: bold; }
        .errormessage { color: #1a5f3b;   font-weight: normal; }
        .result-box   { border: 1px solid #ccc; padding: 15px; margin-top: 20px; border-radius: 8px; }
        h1, h2 { color: #333; }
        a { color: #0066cc; text-decoration: none; }
        a:hover { text-decoration: underline; }
    </style>
</head>
<body>
    <h1>Αποτελέσματα Μεταφοράς Βαθμολογίας</h1>
EOF

#  If an error occurs, the code inside the eval block stops executing, but the script itself does not die. Instead, the error message is stored in the special Perl variable $@
eval {  
 # Generate unique filename based on date and random characters
 my $timestamp        = localtime->strftime("%Y%m%d%H%M%S");
 my $random_chars     = join '', map { ('a'..'z','A'..'Z','0'..'9')[rand(62)] } 1..40;
 my $filename_prefix  = "${timestamp}_${random_chars}";

 # Create directory if it doesn't exist
 # mkpath($pubdir) or die "Failed to create directory $pubdir: $!";

 # Get the original filenames to validate their names (it does not matter that this is after  my $input_file_fh    = $cgi->upload('input_grades'   );
 my $input_original_filename    = $cgi->param('input_grades'   );
 my $template_original_filename = $cgi->param('template_grades');
 check_valid_filename(   $input_original_filename);
 check_valid_filename($template_original_filename);

 # Get uploaded files and form data
 my $input_file_fh    = $cgi->upload('input_grades'   ); # returns a filehandle to the uploaded file. Data has not been uploaded until the 
 my $template_file_fh = $cgi->upload('template_grades');
 my $colAM            = $cgi->param ('colAM'          ); # Column where Student ID is in input    xlsx file $ixlsx, counting from 1 (later we reduce by one to conform with parsexlsx conventions counting from 0)
 my $colGR            = $cgi->param ('colGR'          ); # Column where Grade      is in input    xlsx file $ixlsx, counting from 1
 my $colAO            = 1                              ; # Column where Student ID is in template xlsx file $bxlsx, counting from 1
 my $colGO            = 7                              ; # Column where Grade      is in template xlsx file $bxlsx, counting from 1
 unless ($input_file_fh && $template_file_fh) {die "Fatal Error: Both input and template files must be uploaded.";}

#check_valid_filename(   $input_file_fh); # print "<p>Input    file handle: $input_file_fh   </p>";
#check_valid_filename($template_file_fh); # print "<p>Template file handle: $template_file_fh</p>";


 if    ($colAM > $col_max     ){die "Max numbers of columns in Excel files is $col_max. colAM= $colAM  (Student number column, input    file)";}
 if    ($colGR > $col_max     ){die "Max numbers of columns in Excel files is $col_max. colGR= $colGR  (Grade          column, input    file)";}
 if    ($colAO > $col_max     ){die "Max numbers of columns in Excel files is $col_max. colAO= $colAO  (Student number column, template file)";}
 if    ($colGO > $col_max     ){die "Max numbers of columns in Excel files is $col_max. colGO= $colGO  (Grade          column, template file)";}
 unless($colAM =~ /^[1-9]\d*$/){die "Η Στήλη Αριθμού Μητρώου πρέπει να είναι θετικός ακέραιος. colAM= $colAM";}
 unless($colGR =~ /^[1-9]\d*$/){die "Η Στήλη Βαθμολογίας     πρέπει να είναι θετικός ακέραιος. colGR= $colGR";}
 unless($colAO =~ /^[1-9]\d*$/){die "Η Στήλη Αριθμού Μητρώου πρέπει να είναι θετικός ακέραιος. colAO= $colAO (template file)";}
 unless($colGO =~ /^[1-9]\d*$/){die "Η Στήλη Βαθμολογίας     πρέπει να είναι θετικός ακέραιος. colGO= $colGO (template file)";}

 # Define file paths
 my $ixlsx            = File::Spec->catfile($pubdir, "${filename_prefix}_input.xlsx"   );
 my $bxlsx            = File::Spec->catfile($pubdir, "${filename_prefix}_template.xlsx");
 my $oxlsx            = File::Spec->catfile($pubdir, "${filename_prefix}_output.xlsx"  );
 my $log_file         = File::Spec->catfile($pubdir, "${filename_prefix}_log.txt"      );
 my $oxlsxname        =                              "${filename_prefix}_output.xlsx"   ;
 my $log_filename     =                              "${filename_prefix}_log.txt"       ;
     
 # Save uploaded files
 open my $ifh, '>', $ixlsx or die "Failed to save input    file: $!";
 binmode $ifh;
 while (<$input_file_fh>   ) { print $ifh $_; }
 close   $ifh;

 open my $tfh, '>', $bxlsx or die "Failed to save template file: $!";
 binmode $tfh;
 while (<$template_file_fh>) { print $tfh $_; }
 close   $tfh;

 # Check for file existence and non-zero size after saving
 unless(-r $ixlsx and -s $ixlsx and -r $bxlsx and -s $bxlsx){
  die "Fatal Error: Uploaded files could not be saved correctly or have zero size.";
 }

 # Set up logging to a file
 open $log, '>', $log_file or die "Failed to open log file: $!";
 $GLOG = $log;

 # Call the main logic
 transfer_grades($ixlsx, $bxlsx, $colAM, $colGR, $colGO, $colAO, $oxlsx, $log_file);

 # Compute output URLs
 my $uxlsx = "${puburl}/$oxlsxname"   ; 
 my $ulog  = "${puburl}/$log_filename"; 

 # Print success message and links
 print "<div class=\"result-box success\">";
 print "<h2>Επιτυχία!</h2>";
 print "<p>Η μεταφορά βαθμολογίας ολοκληρώθηκε με επιτυχία.</p>";
 print "<p>Μπορείτε να κατεβάσετε το νέο αρχείο Excel από τον ακόλουθο σύνδεσμο:</p>";
 print "<p><a href=\"" . $uxlsx . "\">Κατεβάστε το αρχείο Excel</a></p>";
 print "<p>Μπορείτε να δείτε το αρχείο καταγραφής από εδώ:</p>";
 print "<p><a href=\"" . $ulog  . "\">Αρχείο Καταγραφής (logfile)</a></p>";
 print "</div>";
}; # eval

#  If an error occurs, the code inside the eval block stops executing, but the script itself does not die. Instead, the error message is stored in the special Perl variable $@
if ($@) {
 # Handle errors gracefully
 print "<div class=\"result-box error\">";
 print "<h2>Σφάλμα</h2>";
 print "<p>Παρουσιάστηκε ένα σφάλμα κατά την εκτέλεση του script:</p>";
 print "<p class=\"errormessage\"> $@ </p>";
 print "</div>";
    
 # Log the error if the log file is open
 if ($log) {print $log "Fatal Error: $@\n";}

} # if ($@)

print "</body></html>";
close $log if $log;

# The transfer_grades function from your existing script,
# adapted to work within the CGI environment.
sub transfer_grades {
 my ($ixlsx, $bxlsx, $colAM, $colGR, $colGO, $colAO, $oxlsx, $log_file) = @_;
    
 # reduce by one, since columns are counted from 0 from here on:
 $colAM--; $colGR--; $colGO--; $colAO--;

 if($colAM > $col_max){die "Excel files must have at most $col_max columns. Student ID column colAM= $colAM";}
 if($colGR > $col_max){die "Excel files must have at most $col_max columns. Grade      column colGR= $colGR";}
 if($colGO > $col_max){die "Excel files must have at most $col_max columns. Grade             colGO= $colGO";}
 if($colAO > $col_max){die "Excel files must have at most $col_max columns. Student ID column colAO= $colAO";}
 # -------------------------------------------------------------------
 # check validity+existence of filenames:
 my @fnames = ($ixlsx, $bxlsx); foreach my $fname (@fnames){check_valid_filename($fname, $log_file);}

 # -------------------------------------------------------------------
 print $GLOG "Reading input grades from $ixlsx\n";
 print $GLOG "----------------------------------------------------------------------------------\n";

 my $rparser   = Spreadsheet::ParseXLSX->new;
 my $rworkbook = $rparser->parse($ixlsx);

 if (!defined $rworkbook){usage($log_file, "Workbook failed to open from $ixlsx");}

 my %grades;
 my $ngrades        = 0;
 my $nsheets        = 0; 
 for my $sheet ($rworkbook->worksheets()) {
  $nsheets++; if ($nsheets > 1) {last;}
  my ($rmin, $rmax) = $sheet->row_range();
  my ($cmin, $cmax) = $sheet->col_range();
  $rmax             = $row_max if $rmax > $row_max;
  $cmax             = $col_max if $cmax > $col_max;
        
  for my $row ($rmin .. $rmax) {
   my $AM = ''; my $GR = '';
   if (     $sheet->get_cell($row, $colAM)) {
    $AM   = $sheet->get_cell($row, $colAM)->unformatted();
    $AM   =~ s/^\s+|\s+$//g;
   }
   if (     $sheet->get_cell($row, $colGR)) {
    $GR   = $sheet->get_cell($row, $colGR)->value();
    $GR   =~ s/^\s+|\s+$//g;
   }
   if ($AM !~ /^0[0-9]{7}$/ && !looks_like_number($AM)) {next;}
   if ($GR !~ /^\d+$/ || $GR < 0 || $GR > 10) {next;}
   $grades{$AM} = $GR;
   $ngrades++;
   print $GLOG "READSGRADES:  ${ngrades}. row= $row :: AM= $AM  grade= $GR :: column_AM= $colAM  column_GR= $colGR\n";
  } # for my $row   ($rmin .. $rmax)
 }  # for my $sheet ($rworkbook->worksheets())

 if($ngrades == 0){die "<span class=\"error\">Reason:</span></br>No grades recorded from input file.</br>\nHave you chosen the correct columns for Student ID and Grade?</br></br>";}
 print $GLOG "----------------------------------------------------------------------------------\n";
 print $GLOG "Reading template $bxlsx\n";
 print $GLOG "Output in $oxlsx\n";
 print $GLOG "----------------------------------------------------------------------------------\n";

 my $tparser        = Spreadsheet::ParseXLSX->new;
 my $tworkbook      = $tparser->parse($bxlsx);
 if (!defined $tworkbook){usage($log_file, "Workbook failed to open from $bxlsx");}

 my $workbook_out   = Excel::Writer::XLSX->new($oxlsx) || usage($log_file, "Failed to create new xlsx file: $!");
    
 for my $tsheet ($tworkbook->worksheets()) {
  my $worksheet_out = $workbook_out->add_worksheet($tsheet->get_name);
  my ($rmin, $rmax) = $tsheet->row_range();
  my ($cmin, $cmax) = $tsheet->col_range();
  $rmax             = $row_max if $rmax > $row_max;
  $cmax             = $col_max if $cmax > $col_max;
  
  my $width_default = 10;
  my @widths        = (16, 40, 25, 25, 70, 25, 10);  # widths of each cell in output xlsx file: put by hand here
        
  for my $col ($cmin .. $cmax) {
   if ($col <= $#widths) {
    $worksheet_out->set_column($col, $col, $widths[$col]);
   } else {
    $worksheet_out->set_column($col, $col, $width_default);
   } # if ($col <= $#widths)
  }  # for my $col ($cmin .. $cmax)
        
  my $title_cell    = $tsheet->get_cell(0, 0);
  my $title         = (defined $title_cell) ? $title_cell->value() : "";
  my $mergeformat   = $workbook_out->add_format(align => 'center');
  $worksheet_out->merge_range(0, 0, 0, $#widths, $title, $mergeformat);
        
  $rmin = 1;
  for my $row ($rmin .. $rmax) {
   my $student_id   = '';
   my $cell_am      = $tsheet->get_cell($row, $colAO);
   $student_id      = defined $cell_am ? $cell_am->unformatted() : '';

   print $GLOG "WRITEGRADES:  row= $row :: AM= $student_id grade= $grades{$student_id} :: column_AM= $colAO  column_GR= $colGO \n";

   for my $col ($cmin .. $cmax) {
    my $cell        = $tsheet->get_cell($row, $col);
    my $value       = defined $cell ? $cell->value() : '';
    
    if ($col       == $colAO) {
     $worksheet_out->write_string($row, $col,         $student_id );
    } elsif ($col  == $colGO && exists        $grades{$student_id}) {
     $worksheet_out->write       ($row, $col, $grades{$student_id});
    } else {
     $worksheet_out->write       ($row, $col, $value              );
    } # if ($col       == $colAO)
   }  # for my $col    ($cmin .. $cmax)
  }   # for my $row    ($rmin .. $rmax)
 }    # for my $tsheet ($tworkbook->worksheets())
 $workbook_out->close();
}     # sub transfer_grades

sub check_valid_filename {
 my ($fname, $log_file) = @_;
 my $maxlen = 200;
 my $regex = qr/[^a-zA-Z0-9\x{0370}-\x{03ff}\_\-\.\(\)\/ ]/u; # The '/u' flag is crucial
 
 # Ensure the input string is treated as UTF-8
 my $utf8_fname = decode('UTF-8', $fname);
 
 if ($utf8_fname =~ $regex) {
  my $invalid_char = $&;
  my $ordinal = ord($invalid_char);
  usage($log_file, "'$fname' does not have a valid name. Invalid character found: '$invalid_char' (U+$ordinal).");
 }
}
# sub check_valid_filename {
#  my ($fname, $log_file) = @_;
#  my $maxlen = 200;

# #This check must be done AFTER files have been saved on disk
# #unless (-r     $fname and -s $fname ) {usage($log_file, "$fname does not exist (or has zero size)"                            );}
# #unless (length($fname)   <=  $maxlen) {usage($log_file, "$fname is too long (> $maxlen characters) length = " . length($fname));}

#  my $regex = qr/[^\_\(\)\.\-\/a-zA-Z0-9\x{0370}-\x{03ff}\ ]/; #Greek UTF8 alphabet, space, and  _().-/
#  if ($fname         =~ $regex) {
#   my $invalid_char  = $&;
#   my $ordinal       = ord($invalid_char);
#   usage($log_file, "'$fname' does not have a valid name. Invalid character found:  (U+$ordinal).");
# # usage($log_file, "'$fname' does not have a valid name. Invalid character found: '$invalid_char' (U+$ordinal).");
#  } # if ($fname         =~ $regex)

#  # if ($fname =~ $regex) {
#  #  my $invalid_char = $&;
#  #  my $ordinal = ord($invalid_char);
#  #  my $error_message = encode_utf8("$fname does not have a valid name. Invalid character found: '$invalid_char' (U+$ordinal).");
#  #  usage($log_file, $error_message);
#  # } # if ($fname =~ $regex)
# }  # sub check_valid_filename
 
sub usage {
 my ($log_file, $message) = @_;
 my $prog = basename($0);
 my $output_message = "Fatal Error: $message\n";
 $output_message .= << "EOF";
----------------------------------------------------------------------------------
Usage: This script is intended to be run as a CGI script. It takes no command-line options.
----------------------------------------------------------------------------------
EOF
 if($GLOG ne *STDOUT){print $GLOG $output_message;} # checks if log file has been opened
 die "<span class=\"error\">Reason:</span></br>" . $message; #exit(1);
#die "<span class=\"error\">Reason:</span></br>" . encode_utf8($message); #exit(1);
} # sub usage

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
