# NatLink macro definitions for NaturallySpeaking
# coding: latin-1
# Generated by vcl2py 2.8.1, Tue Aug 26 00:14:53 2014

import natlink
from natlinkutils import *
from VocolaUtils import *


class ThisGrammar(GrammarBase):

    gramSpec = """
        <1> = 'Pause' ;
        <2> = 'Play' ;
        <functions_arrow> = ('Next' | 'Previous' | 'Increase' | 'Decrease' ) ;
        <3> = (('Next' | 'Previous' | 'Increase' | 'Decrease' ) ) ;
        <4> = 'podcast view' ;
        <5> = 'right side' ;
        <6> = 'left side' ;
        <7> = 'current podcast' ;
        <any> = <1>|<2>|<3>|<4>|<5>|<6>|<7>;
        <sequence> exported = <any>;
    """
    
    def initialize(self):
        self.load(self.gramSpec)
        self.currentModule = ("","",0)
        self.ruleSet1 = ['sequence']

    def gotBegin(self,moduleInfo):
        # Return if wrong application
        window = matchWindow(moduleInfo,'itunes','')
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

    # 'Pause'
    def gotResults_1(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += ' '
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_1(words[1:], fullResults)
        except Exception, e:
            handle_error('itunes.vcl', 5, '\'Pause\'', e)
            self.firstWord = -1

    # 'Play'
    def gotResults_2(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += ' '
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_2(words[1:], fullResults)
        except Exception, e:
            handle_error('itunes.vcl', 6, '\'Play\'', e)
            self.firstWord = -1

    def get_functions_arrow(self, list_buffer, functional, word):
        if word == 'Next':
            list_buffer += 'Right'
        elif word == 'Previous':
            list_buffer += 'Left'
        elif word == 'Increase':
            list_buffer += 'Up'
        elif word == 'Decrease':
            list_buffer += 'Down'
        return list_buffer

    # (('Next' | 'Previous' | 'Increase' | 'Decrease'))
    def gotResults_3(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+'
            word = fullResults[0 + self.firstWord][0]
            if word == 'Next':
                top_buffer += 'Right'
            elif word == 'Previous':
                top_buffer += 'Left'
            elif word == 'Increase':
                top_buffer += 'Up'
            elif word == 'Decrease':
                top_buffer += 'Down'
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_3(words[1:], fullResults)
        except Exception, e:
            handle_error('itunes.vcl', 8, '((\'Next\' | \'Previous\' | \'Increase\' | \'Decrease\'))', e)
            self.firstWord = -1

    # 'podcast view'
    def gotResults_4(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Esc}{Tab_7}{Enter}{Down_2}{Enter}'
            limit = ''
            limit += '6'
            for i in range(to_long(limit)):
                top_buffer += '{Shift+Tab}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_4(words[1:], fullResults)
        except Exception, e:
            handle_error('itunes.vcl', 9, '\'podcast view\'', e)
            self.firstWord = -1

    # 'right side'
    def gotResults_5(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Tab_7}{Down}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_5(words[1:], fullResults)
        except Exception, e:
            handle_error('itunes.vcl', 10, '\'right side\'', e)
            self.firstWord = -1

    # 'left side'
    def gotResults_6(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            limit = ''
            limit += '7'
            for i in range(to_long(limit)):
                top_buffer += '{Shift+Tab}'
            top_buffer += '{Down}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_6(words[1:], fullResults)
        except Exception, e:
            handle_error('itunes.vcl', 11, '\'left side\'', e)
            self.firstWord = -1

    # 'current podcast'
    def gotResults_7(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+l}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_7(words[1:], fullResults)
        except Exception, e:
            handle_error('itunes.vcl', 12, '\'current podcast\'', e)
            self.firstWord = -1

thisGrammar = ThisGrammar()
thisGrammar.initialize()

def unload():
    global thisGrammar
    if thisGrammar: thisGrammar.unload()
    thisGrammar = None
