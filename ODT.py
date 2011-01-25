
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


from xml.dom import minidom
import shutil
import tempfile
import zipfile


class ODT(object):
    '''
    ODT is a class representing an odt file.  It has methods to create a
    new odt file, to add text to and set the font sizes of that text in the 
    odt file.
    '''
    
    odt_template_filename = 'template-blank.odt'    

    def __init__(self):
        '''
        Constructor
        '''
        odt = zipfile.ZipFile(self.odt_template_filename, 'r')
        self.odt_ls = odt.namelist()
        #Create temp dir for constuction of a new odt file.
        self.tmp_odt_dir = tempfile.mkdtemp()
        odt.extractall(self.tmp_odt_dir)
        odt.close()
        
        self.content_path = self.tmp_odt_dir + '/content.xml'
        
        self.__get_content_xml()
        
    def __get_content_xml(self):
        """
        Returns the DOM of the content.xml file in the odt archive.
        """
        self.content_xml = minidom.parse(self.content_path)
        
    def parse_word_list(self, word_list):
        """
        Takes a list of words and the weighting to be applied to each one
        as a list of tuples.
        Parses each word and weighting and then adds each them to the
        content.xml file of the odt archive.  It creates the correct style 
        elements for each weighting/font size. 
        """
        self.__setup_styles_and_para_elements()
        self.size_list = []
        
        for word in word_list:
            size = word[1]
            if size not in self.size_list:
                self.size_list.append(size)
                self.__create_new_style_node(size)
            #Check each word has a trailing space
            text = self.__check_spaces(word[0])
            #Create the text:span node from each piece of text
            #and its associated style
            self.__create_text_nodes(text, size)
            
    def parse_words_for_wordle(self,word_list, advanced=False):
        """
        Takes a list of words and the weighting to be applied to each one
        as a list of tuples.
        Parses each word and weighting and then concatenates weighting number
        of each word together.
        """
        self.text = self.content_xml.getElementsByTagName('office:text')
        self.size_list = []
        
        for word in word_list:
            size = word[1]
            #Check each word has a trailing space
            text = self.__check_spaces(word[0])
            #Concatenate
            if advanced is True: 
                text = text + ':' + size
            else:
                text = text * int(size)
                
            self.__create_wordle_text_nodes(text, size)
             
        
    def __setup_styles_and_para_elements(self):
        """
        Capture the automatic-styles and paragraph elements that hold the font size
        data and the text respectively.
        """
        self.styles = self.content_xml.getElementsByTagName('office:automatic-styles')
        self.paragraph = self.content_xml.getElementsByTagName('text:p')
        
    def __create_new_style_node(self,size):
        """
        Creates a new style node and text-properties node and appends 
        them to the style element
        """
        elem = self.content_xml.createElement('style:style')
        elem.setAttribute('style:name', str.format('T{0}', size))
        elem.setAttribute('style:family','text')
        self.styles[0].appendChild(elem)

        elem = self.content_xml.createElement('style:text-properties')
        pt_size = str.format('{0}pt', size)
        elem.setAttribute('fo:font-size', pt_size)
        elem.setAttribute('style:font-size-asian', pt_size)
        elem.setAttribute('style:font-size-complex', pt_size)
        self.styles[0].lastChild.appendChild(elem) 
        
    def __check_spaces(self, word):
        """
        Each word needs a space after it to be correctly formatted 
        in the re-constructed sentance.
        """
        text = word[0]
        if text[-1:] != ' ':
            text = word + " " 
        return text   
        
    def __create_text_nodes(self, text, size):
        """
        Each text span node contains the text in that span and its 
        associated style
        """ 
        elem = self.content_xml.createElement('text:span')
        elem.setAttribute('text:style-name',str.format('T{0}', size))
        text_elem = self.content_xml.createTextNode(text)
        elem.appendChild(text_elem)
        self.paragraph[0].appendChild(elem)
        
    def __create_wordle_text_nodes(self, text,size):
        """
        Each text:p node contains the text and its 
        weighting.
        """ 
        elem = self.content_xml.createElement('text:p')
        text_elem = self.content_xml.createTextNode(text)
        elem.appendChild(text_elem)
        self.text[-1].appendChild(elem)
        
    def save(self,filename):
        """
        Saves the new content.xml to a file and then recreates 
        the odt archive
        """
        content = open(self.content_path,'w')
        self.content_xml.writexml(content, encoding='utf-8')
        content.close()
        
        new_odt = zipfile.ZipFile(filename, 'w')
        for file in self.odt_ls:
            new_odt.write(self.tmp_odt_dir + '/' + file, file)
            
        new_odt.close()
        
    def finish(self):
        """
        Clean up.
        Temp odt directory is deleted here.
        """
        shutil.rmtree(self.tmp_odt_dir)
        
        