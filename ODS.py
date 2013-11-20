from xml.dom import minidom
import tempfile
import zipfile

# Copyright 2010 - 2011 Paul Elms
# 
# This file is part of word_weight.
# 
# It is free software: you can redistribute it and/or modify
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


class ODS(object):
    """
    Class for dealing with an ods file.  It opens the file and 
    parses the word and weight data.
    """
    
    def __init__(self, filename):
        """
        Constructor.
        Opens the ods archive, extracts the content.xml into a temp
        directory and closes the file again.
        """
        ods = zipfile.ZipFile(filename,'r')
        self.temp_dir = tempfile.mkdtemp()
        ods.extract('content.xml',self.temp_dir)
        ods.close()
        
    def parse_data(self):
        """
        Parses the content of the ods. Extracts the words and weights from
        the file.
        Returns a list of word, weight tuples.
        """
        ods_xml = minidom.parse(self.temp_dir + '/content.xml')
        rows = ods_xml.getElementsByTagName('table:table-row')
        
        word_list = []
        for row in rows:
            if row.childNodes.length > 1:
                if row.childNodes[1].getAttribute('office:value-type') == 'float':
                    text = row.childNodes[0].childNodes[0].childNodes[0].nodeValue
                    size = row.childNodes[1].childNodes[0].childNodes[0].nodeValue
    
                    word_list.append((text, size))
        return word_list
        
