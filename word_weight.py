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
from xml.dom import minidom
import ODT
import tempfile
import zipfile


def __read_cmd_args():
    """
    Set up the Option Parser to get the ods filename and output
    filename from the command line parameters.
    """
    usage = "usage: ./word_weight.py -f ODS_FILE [-o OUTPUT_FILE]\n\nTo run supply the ods file with the data and an optional output filename."
    parser = OptionParser(usage=usage)
    parser.add_option("-f", "--filename", metavar="ODS_FILE", action="store", type="string", dest="ods_filename", help="Specify an .ODS_FILE to get data from.")
    parser.add_option("-o", "--output-file", metavar="OUTPUT_FILE", action="store", type="string", dest="output_filename", help="Specify an OUTPUT_FILE.")
    return parser.parse_args()


def __get_rows(filename):
    """
    Open the ods spreadsheet and get each row in to a list to return.
    """
    ods_file = filename
    ods = zipfile.ZipFile(ods_file,'r')
    tmp_ods_dir = tempfile.mkdtemp()
    ods.extract('content.xml',tmp_ods_dir)
    ods.close()

    ods_xml = minidom.parse(tmp_ods_dir + '/content.xml')
    rows = ods_xml.getElementsByTagName('table:table-row')

    return rows

def __parse_rows(rows):
    """
    Parse rows from the ods file and return a list of tuples of each word and
    its associated weighting.
    """
    word_list = []
    for row in rows:
        if row.childNodes.length > 1:
            if row.childNodes[1].getAttribute('office:value-type') == 'float':
                text = row.childNodes[0].childNodes[0].childNodes[0].nodeValue
                size = row.childNodes[1].childNodes[0].childNodes[0].nodeValue
    
                word_list.append((text, size))
    return word_list
                
    
#Program starts here..
options, args = __read_cmd_args()
rows = __get_rows(options.ods_filename)
word_list = __parse_rows(rows)

new_odt = ODT.ODT()
new_odt.parse_word_list(word_list)
new_odt.save(options.output_filename)
new_odt.finish()


"""

#Run through list of font sizes to find all unique sizes and add them to the odt template

for word in word_list:
    size = word[1]
    if size not in size_list:
        size_list.append(size)
        #create new style node and text-properties node and append them to the document
        elem = odt_xml.createElement('style:style')
        elem.setAttribute('style:name', str.format('T{0}', size))
        elem.setAttribute('style:family','text')
        styles[0].appendChild(elem)

        elem = odt_xml.createElement('style:text-properties')
        pt_size = str.format('{0}pt', size)
        elem.setAttribute('fo:font-size', pt_size)
        elem.setAttribute('style:font-size-asian', pt_size)
        elem.setAttribute('style:font-size-complex', pt_size)
        styles[0].lastChild.appendChild(elem)        
    #Make sure each word has a space after it.
    text = word[0]
    if text[-1:] != ' ':
        text = word[0] + " "
        
        
    #Create the text:span node from each piece of text and its associated style
    elem = odt_xml.createElement('text:span')
    elem.setAttribute('text:style-name',str.format('T{0}', size))
    text_elem = odt_xml.createTextNode(text)
    elem.appendChild(text_elem)
    paragraph[0].appendChild(elem)

#Save modifyed content.xml in a new odt file
f = open(odt_path,'w')
odt_xml.writexml(f, encoding='utf-8')
f.close()
new_odt = zipfile.ZipFile(options.output_filename, 'w')
for name in odt_ls:
    new_odt.write(tmp_odt_dir + '/' + name, name)

new_odt.close()

shutil.rmtree(tmp_odt_dir)
"""


