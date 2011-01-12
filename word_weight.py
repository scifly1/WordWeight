#! /usr/bin/env python

# Copyright 2010 Paul Elms
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

import zipfile
from xml.dom import minidom
from optparse import OptionParser
import tempfile
import shutil

#TODO Break up into functions to improve readability

# Set up optparse to get ods filename and to supply an output filename.
def __read_cmd_args():
    usage = "usage: ./word_weight.py -f ODS_FILE [-o OUTPUT_FILE]\n\nTo run supply the ods file with the data and an optional output filename."
    parser = OptionParser(usage=usage)
    parser.add_option("-f", "--filename", metavar="ODS_FILE", action="store", type="string", dest="ods_filename", help="Specify an .ODS_FILE to get data from.")
    parser.add_option("-o", "--output-file", metavar="OUTPUT_FILE", action="store", type="string", dest="output_filename", help="Specify an OUTPUT_FILE.")
    return parser.parse_args()

options, args = __read_cmd_args()

#open ods spreadsheet and get each word and its size in a list of tuples.
ods_file = options.ods_filename
ods = zipfile.ZipFile(ods_file,'r')
tmp_ods_dir = tempfile.mkdtemp()
ods.extract('content.xml',tmp_ods_dir)
ods.close()

ods_xml = minidom.parse(tmp_ods_dir + '/content.xml')
rows = ods_xml.getElementsByTagName('table:table-row')

word_list = []
#Create word and associated size list from ods file.
for row in rows:
    if row.childNodes.length > 1:
        if row.childNodes[1].getAttribute('office:value-type') == 'float':
            text = row.childNodes[0].childNodes[0].childNodes[0].nodeValue
            size = row.childNodes[1].childNodes[0].childNodes[0].nodeValue
    
            word_list.append((text, size))

#Extract the template odt file and parse content.xml
odt_template = 'template-blank.odt'
odt = zipfile.ZipFile(odt_template, 'r')
odt_ls = odt.namelist()
tmp_odt_dir = tempfile.mkdtemp()
odt.extractall(tmp_odt_dir)
odt.close()
odt_path = tmp_odt_dir + '/content.xml'
odt_xml = minidom.parse(odt_path)

styles = odt_xml.getElementsByTagName('office:automatic-styles')
paragraph = odt_xml.getElementsByTagName('text:p')

#Run through list of font sizes to find all unique sizes and add them to the odt template
size_list = []
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
shutil.rmtree(tmp_ods_dir)
shutil.rmtree(tmp_odt_dir)

