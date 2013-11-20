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
import sys
import ODS
import ODT


def __get_parser():
    """
    Set up the Option Parser to get the ods filename and output
    filename from the command line parameters.
    """
    usage = "usage: ./word_weight.py [-w:-a] [-o OUTPUT_FILE] ODS_FILE\n\nTo run, supply the ods file with the data and an optional output filename."
    parser = OptionParser(usage=usage)
    parser.set_defaults(output_filename="./output")
    parser.add_option("-o", "--output-file", metavar="OUTPUT_FILE", action="store", type="string", dest="output_filename", help="Specify an OUTPUT_FILE.")
    parser.add_option("-w", "--wordle", action="store_true", dest="wordle", default=False, help="Runs in Wordle mode.")
    parser.add_option("-a", "--wordle-advanced", action="store_true", dest="wordle_adv", default=False, help="Runs in Wordle advanced mode.")
    return parser

    
#Program starts here..
parser = __get_parser()
options, args = parser.parse_args()

if len(args) != 1 or args[0] == None:
    print("\nYou must supply a source ODS_FILE filename\n")
    parser.print_help()
    sys.exit(0)
 
ods = ODS.ODS(args[0])
word_list = ods.parse_data()
    
new_odt = ODT.ODT()
if options.wordle:
    print "Wordle mode selected"
    new_odt.parse_words_for_wordle(word_list)
elif options.wordle_adv:
    print "Wordle advanced mode selected"
    new_odt.parse_words_for_wordle(word_list, True)
else:
    new_odt.parse_word_list(word_list)
    
new_odt.save(options.output_filename)
new_odt.finish()




