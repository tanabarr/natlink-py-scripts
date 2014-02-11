# NatLink macro definitions for NaturallySpeaking
# coding: latin-1
# Generated by vcl2py 2.8.1, Tue Feb 11 11:32:14 2014

import natlink
from natlinkutils import *
from VocolaUtils import *


class ThisGrammar(GrammarBase):

    gramSpec = """
        <dgndictation> imported;
        <1> = 'Search bar' ;
        <2> = 'Search for' <dgndictation> ;
        <next_back> = ('next' | 'back' ) ;
        <3> = <next_back> 'track' ;
        <4> = ('play' | 'pause' ) ;
        <5> = 'volume' ('Up' | 'Down' ) ;
        <any> = <1>|<2>|<3>|<4>|<5>;
        <sequence> exported = <any>;
    """
    
    def initialize(self):
        self.load(self.gramSpec)
        self.currentModule = ("","",0)
        self.ruleSet1 = ['sequence']

    def gotBegin(self,moduleInfo):
        # Return if wrong application
        window = matchWindow(moduleInfo,'spotify','')
        if not window: return None
        self.firstWord = 0
        # Return if same window and title as before
        if moduleInfo == self.currentModule: return None
        self.currentModule = moduleInfo

        self.deactivateAll()
        title = string.lower(moduleInfo[1])
        if string.find(title,'') >= 0:
            for rule in self.ruleSet1:
                try:
                    self.activate(rule,window)
                except BadWindow:
                    pass

    def convert_number_word(self, word):
        if   word == '0':
            return '0'
        else:
            return word

    # 'Search bar'
    def gotResults_1(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+l}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_1(words[1:], fullResults)
        except Exception, e:
            handle_error('spotify.vcl', 5, '\'Search bar\'', e)
            self.firstWord = -1

    # 'Search for' <_anything>
    def gotResults_2(self, words, fullResults):
        if self.firstWord<0:
            return
        fullResults = combineDictationWords(fullResults)
        opt = 1 + self.firstWord
        if opt >= len(fullResults) or fullResults[opt][1] != 'converted dgndictation':
            fullResults.insert(opt, ['', 'converted dgndictation'])
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+l}'
            word = fullResults[1 + self.firstWord][0]
            top_buffer += word
            top_buffer += '{Enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
        except Exception, e:
            handle_error('spotify.vcl', 6, '\'Search for\' <_anything>', e)
            self.firstWord = -1

    def get_next_back(self, list_buffer, functional, word):
        if word == 'next':
            list_buffer += 'Right'
        elif word == 'back':
            list_buffer += 'Left'
        return list_buffer

    # <next_back> 'track'
    def gotResults_3(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Esc}{Ctrl+'
            word = fullResults[0 + self.firstWord][0]
            top_buffer = self.get_next_back(top_buffer, False, word)
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
        except Exception, e:
            handle_error('spotify.vcl', 10, '<next_back> \'track\'', e)
            self.firstWord = -1

    # ('play' | 'pause')
    def gotResults_4(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += ' '
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_4(words[1:], fullResults)
        except Exception, e:
            handle_error('spotify.vcl', 12, '(\'play\' | \'pause\')', e)
            self.firstWord = -1

    # 'volume' ('Up' | 'Down')
    def gotResults_5(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+'
            word = fullResults[1 + self.firstWord][0]
            top_buffer += word
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_5(words[2:], fullResults)
        except Exception, e:
            handle_error('spotify.vcl', 13, '\'volume\' (\'Up\' | \'Down\')', e)
            self.firstWord = -1

thisGrammar = ThisGrammar()
thisGrammar.initialize()

def unload():
    global thisGrammar
    if thisGrammar: thisGrammar.unload()
    thisGrammar = None
