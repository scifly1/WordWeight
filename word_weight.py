#! /usr/bin/env python

# Copyright 2010 - 2011 Paul Elms
# 
# word_weight is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# 
# word_weight is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with word_weight.  If not, see <http://www.gnu.org/licenses/>.

from optparse import OptionParser
import ODS
import ODT


def __read_cmd_args():
    """
    Set up the Option Parser to get the ods filename and output
    filename from the command line parameters.
    """
    usage = "usage: ./word_weight.py -f ODS_FILE [-o OUTPUT_FILE] -w\n\nTo run, supply the ods file with the data and an optional output filename."
    parser = OptionParser(usage=usage)
    parser.set_defaults(output_filename="./output")
    parser.add_option("-f", "--filename", metavar="ODS_FILE", action="store", type="string", dest="ods_filename", help="Specify an .ODS_FILE to get data from.")
    parser.add_option("-o", "--output-file", metavar="OUTPUT_FILE", action="store", type="string", dest="output_filename", help="Specify an OUTPUT_FILE.")
    parser.add_option("-w", "--wordle", action="store_true", dest="wordle", default=False, help="Runs in Wordle mode.")
    return parser.parse_args()

    
#Program starts here..
options, args = __read_cmd_args()

ods = ODS.ODS(options.ods_filename)
word_list = ods.parse_data()

if options.wordle:
    print "Wordle mode selected"
    
new_odt = ODT.ODT()
new_odt.parse_word_list(word_list)
new_odt.save(options.output_filename)
new_odt.finish()




