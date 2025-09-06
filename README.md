# Μεταφορά Βαθμολογίας σε Φοιτητολόγιο ΕΜΠ

The tools provided here transfer student grades from your personal Excel (xlxs) file, to a file that can be imported to the [official student records](https://e-sec.ntua.gr) of the [National Technical University of Athens](https://www.ntua.gr).

There is an input file, which must be an xlsx file, containing student numbers in a numeric form starting with 0 in a column of your choice, and the grade in integer form in another column of your choice. The restrictions are quite mild, because the script ignores all lines that don't contain a student number, therefore ti may contain a lot of useful information for you, without affecting the function of the script. Your file does not need to have records for all students, only the existing ones will be processed.

Then, there is a template file, which is the export of the grades section as an xlsx file, taken from  [e-sec.ntua.gr](https://e-sec.ntua.gr). The format of the file is very strict, and you should leave it as it is.

The script takes as input:
  1. The name of the input xlsx file
  2. The name of the output xlsx file
  3. The column numbers where student numbers and grades are recorded in your input xlsx file.
  
Columns are counted by starting from 1 for column A.

Then the script provides an output xlsx file, of the same form as the template file, which can be imported to  [e-sec.ntua.gr](https://e-sec.ntua.gr). If there are already grades recorded in the input file, the grades from the input file will be **added** to the output file. Old grades will be overwritten, or left intact. 

# Web Interface

An example of the web interface with the script can be [found here](https://physics.ntua.gr/grades.html).

## Relevant Files:

1. `grades_transfer.cgi`: the actual script running as a CGI-script. It should be placed in the cgi-bin directory of your web server, after editing it and put the correct paths to your filesystem
2. `grades_transfer_upload_form.html`: an example of an html file that provides a form that runs the CGI-script

The script is written in Perl, assumes a Unix filesystem (like in Linux), and uses some non-standard modules to process xlxs files, which you should install. See the Prerequisites section below. You should edit it and set the variables that point to pathes in your filesystem.

For safety, there are many checks that put restrictions to the filenames, see `grades_transfer_upload_form.html`.

# CLI Interface

There is also a command line script `grades_transfer_command_line.pl`, written in Perl. Use the `-h` flag to see the help message:

```
Usage: grades_transfer_command_line.pl [options]

      -i <input>       xlsx file with grades in column determined by the -a option                 (default: 5) and student number in column determined by by the -g option (default: 8)
      -t <template>    xlsx file from e-sec.ntua.gr. Existing grades will be either replaced or left as is
      -a <col num>     Column number where the AM  student number is recorded                      (default: 5). Columns are counted from 1: Column A is 1, column B is 2 and so on.
      -g <col num>     Column number where the new student grade  is recorded                      (default: 8)
      -A <col num>     Column number where the AM  student number is recorded in the template file (default: 1)
      -o <col num>     Column number where the students grades are in the template file            (default: 7) - clean this column if you want to replace grades
      -d               Sets debugging mode

Columns in xlsx files are counted from 1 (column A is 1, column B is 2 and so on)
The script ADDS/REPLACES grades: If you have a preexisting template that already contains grades, it will add new grades from the input xlsx file, replace old grades, but will leave intact already existing grades.
```

There are no restrictions in filenames anymore, since you are in your (hopefully) safe environament. 

The relevant files function the same way as in the web interface.

# Prerequisites

You need Perl to run in a Unix-like environment.

Perl uses the following modules that need to be installed:

```
Spreadsheet::ParseXLSX, Spreadsheet::ParseExcel::Cell, Excel::Writer::XLSX
```

You can use `cpan -i <module name>` to install a module with all its dependencies.

The full list of modules used, are listed in the beginning of the files `grades_transfer.cgi` and `grades_transfer_command_line.pl`.




