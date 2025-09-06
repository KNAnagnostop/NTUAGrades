#!/bin/bash

cdir=$(pwd)
prog=$(basename $0)
# Configuration
URL="https://physics.ntua.gr/cgi-a/Grades/grades.cgi"
INPUT_FILE="${HOME}/Downloads/00_hi.xlsx"
TEMPLATE_FILE="${HOME}/Downloads/00_hi-t.xlsx"
COL_AM="1"
COL_GR="2"
filename  () { basename "$1" ".${1##*.}"; } # filename  file.tar.gz = file.tar
html="$(filename ${prog})_out.html"
# Check if required files exist
if [ ! -f "$INPUT_FILE" ]; then
    echo "Error: Input file '$INPUT_FILE' not found."
    exit 1
fi

if [ ! -f "$TEMPLATE_FILE" ]; then
    echo "Error: Template file '$TEMPLATE_FILE' not found."
    exit 1
fi

echo "Sending files to CGI script..."

# Use curl to send the files and form data
curl --silent \
     -F "input_grades=@$INPUT_FILE" \
     -F "template_grades=@$TEMPLATE_FILE" \
     -F "colAM=$COL_AM" \
     -F "colGR=$COL_GR" \
     "$URL" > $html

echo "#-------------------------------------------------------------------"
echo "#----------- HTML response in ${html}: -----------------------------"
cat $html
echo "#-------------------------------------------------------------------"

fnam=$(cat $html | perl -nle 'if(m{PUB/GradesForms/(.*?)\_output.xlsx}){print $1;}')
xlsx=$(cat $html | perl -nle 'if(m/(https:.*?_output.xlsx)/){print $1;}')
logf=$(cat $html | perl -ne  'if(/href="([^"]+)_log\.txt"/){print "$1_log.txt\n"}')

echo "Download with the commands:"
echo "wget \"$xlsx\""
echo "wget \"$logf\""
echo " "
echo "Or, using curl:"
echo "curl \"$xlsx\" --output \"${fnam}_output.xlsx\""
echo "curl \"$logf\" --output \"${fnam}_log.txt\""

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
