# NatLink macro definitions for NaturallySpeaking
# coding: latin-1
# Generated by vcl2py 2.8.1, Thu Mar 20 14:44:57 2014

import natlink
from natlinkutils import *
from VocolaUtils import *


class ThisGrammar(GrammarBase):

    gramSpec = """
        <n> = ('zero' | 'one' | 'two' | 'three' | 'four' | 'five' | 'six' | 'seven' | 'eight' | 'nine') ;
        <1> = 'zoom' ('in' | 'out' ) <n> ;
        <19> = 'zoom' ('in' | 'out' ) ;
        <2> = 'save' ;
        <3> = 'new' ;
        <4> = 'last' ;
        <5> = 'new window' ;
        <6> = 'next' <n> ;
        <7> = 'previous' <n> ;
        <8> = 'switch tab' <n> ;
        <9> = 'private' ;
        <10> = 'close' ;
        <11> = 'bookmark' ;
        <12> = 'tools' ;
        <13> = 'reload' ;
        <14> = 'back page' ;
        <15> = ('Copy' | 'Paste' | 'Go' ) ('Address' | 'URL' ) ;
        <16> = 'text box' ;
        <17> = 'show links' ;
        <18> = 'copy links' ;
        <any> = <1>|<19>|<2>|<3>|<4>|<5>|<6>|<7>|<8>|<9>|<10>|<11>|<12>|<13>|<14>|<15>|<16>|<17>|<18>;
        <sequence> exported = <any>;
    """
    
    def initialize(self):
        self.load(self.gramSpec)
        self.currentModule = ("","",0)
        self.ruleSet1 = ['sequence']

    def gotBegin(self,moduleInfo):
        # Return if wrong application
        window = matchWindow(moduleInfo,'chrome','')
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
        if   word == 'zero':
            return '0'
        elif word == 'one':
            return '1'
        elif word == 'two':
            return '2'
        elif word == 'three':
            return '3'
        elif word == 'four':
            return '4'
        elif word == 'five':
            return '5'
        elif word == 'six':
            return '6'
        elif word == 'seven':
            return '7'
        elif word == 'eight':
            return '8'
        elif word == 'nine':
            return '9'
        else:
            return word

    def get_n(self, list_buffer, functional, word):
        list_buffer += self.convert_number_word(word)
        return list_buffer

    # 'zoom' ('in' | 'out') <n>
    def gotResults_1(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            when_value = ''
            word = fullResults[2 + self.firstWord][0]
            when_value = self.get_n(when_value, True, word)
            if when_value != "":
                limit2 = ''
                word = fullResults[2 + self.firstWord][0]
                limit2 = self.get_n(limit2, True, word)
                for i in range(to_long(limit2)):
                    top_buffer += '{Ctrl+'
                    word = fullResults[1 + self.firstWord][0]
                    if word == 'in':
                        top_buffer += 'plus'
                    elif word == 'out':
                        top_buffer += 'minus'
                    top_buffer += '}'
                    top_buffer = do_flush(False, top_buffer);
                    dragon3_arg1 = ''
                    dragon3_arg1 += '100'
                    call_Dragon('Wait', 'i', [dragon3_arg1])
            else:
                limit2 = ''
                limit2 += '1'
                for i in range(to_long(limit2)):
                    top_buffer += '{Ctrl+'
                    word = fullResults[1 + self.firstWord][0]
                    if word == 'in':
                        top_buffer += 'plus'
                    elif word == 'out':
                        top_buffer += 'minus'
                    top_buffer += '}'
                    top_buffer = do_flush(False, top_buffer);
                    dragon3_arg1 = ''
                    dragon3_arg1 += '100'
                    call_Dragon('Wait', 'i', [dragon3_arg1])
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 3
        except Exception, e:
            handle_error('chrome.vcl', 8, '\'zoom\' (\'in\' | \'out\') <n>', e)
            self.firstWord = -1

    # 'zoom' ('in' | 'out')
    def gotResults_19(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            when_value = ''
            when_value += ''
            if when_value != "":
                limit2 = ''
                limit2 += ''
                for i in range(to_long(limit2)):
                    top_buffer += '{Ctrl+'
                    word = fullResults[1 + self.firstWord][0]
                    if word == 'in':
                        top_buffer += 'plus'
                    elif word == 'out':
                        top_buffer += 'minus'
                    top_buffer += '}'
                    top_buffer = do_flush(False, top_buffer);
                    dragon3_arg1 = ''
                    dragon3_arg1 += '100'
                    call_Dragon('Wait', 'i', [dragon3_arg1])
            else:
                limit2 = ''
                limit2 += '1'
                for i in range(to_long(limit2)):
                    top_buffer += '{Ctrl+'
                    word = fullResults[1 + self.firstWord][0]
                    if word == 'in':
                        top_buffer += 'plus'
                    elif word == 'out':
                        top_buffer += 'minus'
                    top_buffer += '}'
                    top_buffer = do_flush(False, top_buffer);
                    dragon3_arg1 = ''
                    dragon3_arg1 += '100'
                    call_Dragon('Wait', 'i', [dragon3_arg1])
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_19(words[2:], fullResults)
        except Exception, e:
            handle_error('chrome.vcl', 8, '\'zoom\' (\'in\' | \'out\')', e)
            self.firstWord = -1

    # 'save'
    def gotResults_2(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+s}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_2(words[1:], fullResults)
        except Exception, e:
            handle_error('chrome.vcl', 9, '\'save\'', e)
            self.firstWord = -1

    # 'new'
    def gotResults_3(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+t}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_3(words[1:], fullResults)
        except Exception, e:
            handle_error('chrome.vcl', 10, '\'new\'', e)
            self.firstWord = -1

    # 'last'
    def gotResults_4(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+T}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_4(words[1:], fullResults)
        except Exception, e:
            handle_error('chrome.vcl', 11, '\'last\'', e)
            self.firstWord = -1

    # 'new window'
    def gotResults_5(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+n}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_5(words[1:], fullResults)
        except Exception, e:
            handle_error('chrome.vcl', 12, '\'new window\'', e)
            self.firstWord = -1

    # 'next' <n>
    def gotResults_6(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            limit = ''
            word = fullResults[1 + self.firstWord][0]
            limit = self.get_n(limit, True, word)
            for i in range(to_long(limit)):
                top_buffer += '{Ctrl+tab}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
        except Exception, e:
            handle_error('chrome.vcl', 13, '\'next\' <n>', e)
            self.firstWord = -1

    # 'previous' <n>
    def gotResults_7(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            limit = ''
            word = fullResults[1 + self.firstWord][0]
            limit = self.get_n(limit, True, word)
            for i in range(to_long(limit)):
                top_buffer += '{Ctrl+Shift+tab}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
        except Exception, e:
            handle_error('chrome.vcl', 14, '\'previous\' <n>', e)
            self.firstWord = -1

    # 'switch tab' <n>
    def gotResults_8(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl'
            top_buffer += '+'
            word = fullResults[1 + self.firstWord][0]
            top_buffer = self.get_n(top_buffer, False, word)
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
        except Exception, e:
            handle_error('chrome.vcl', 15, '\'switch tab\' <n>', e)
            self.firstWord = -1

    # 'private'
    def gotResults_9(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+e}'
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '10'
            call_Dragon('Wait', 'i', [dragon_arg1])
            top_buffer += '{i}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_9(words[1:], fullResults)
        except Exception, e:
            handle_error('chrome.vcl', 16, '\'private\'', e)
            self.firstWord = -1

    # 'close'
    def gotResults_10(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+w}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_10(words[1:], fullResults)
        except Exception, e:
            handle_error('chrome.vcl', 17, '\'close\'', e)
            self.firstWord = -1

    # 'bookmark'
    def gotResults_11(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+d}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_11(words[1:], fullResults)
        except Exception, e:
            handle_error('chrome.vcl', 18, '\'bookmark\'', e)
            self.firstWord = -1

    # 'tools'
    def gotResults_12(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+e}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_12(words[1:], fullResults)
        except Exception, e:
            handle_error('chrome.vcl', 19, '\'tools\'', e)
            self.firstWord = -1

    # 'reload'
    def gotResults_13(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{f5}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_13(words[1:], fullResults)
        except Exception, e:
            handle_error('chrome.vcl', 20, '\'reload\'', e)
            self.firstWord = -1

    # 'back page'
    def gotResults_14(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{backspace}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_14(words[1:], fullResults)
        except Exception, e:
            handle_error('chrome.vcl', 21, '\'back page\'', e)
            self.firstWord = -1

    # ('Copy' | 'Paste' | 'Go') ('Address' | 'URL')
    def gotResults_15(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+d}'
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '20'
            call_Dragon('Wait', 'i', [dragon_arg1])
            word = fullResults[0 + self.firstWord][0]
            if word == 'Copy':
                top_buffer += '{Ctrl+c}'
            elif word == 'Paste':
                top_buffer += '{Ctrl+v}'
            elif word == 'Go':
                top_buffer += ''
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_15(words[2:], fullResults)
        except Exception, e:
            handle_error('chrome.vcl', 22, '(\'Copy\' | \'Paste\' | \'Go\') (\'Address\' | \'URL\')', e)
            self.firstWord = -1

    # 'text box'
    def gotResults_16(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Esc}gi'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_16(words[1:], fullResults)
        except Exception, e:
            handle_error('chrome.vcl', 24, '\'text box\'', e)
            self.firstWord = -1

    # 'show links'
    def gotResults_17(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Esc}f'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_17(words[1:], fullResults)
        except Exception, e:
            handle_error('chrome.vcl', 25, '\'show links\'', e)
            self.firstWord = -1

    # 'copy links'
    def gotResults_18(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Esc}yf'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_18(words[1:], fullResults)
        except Exception, e:
            handle_error('chrome.vcl', 28, '\'copy links\'', e)
            self.firstWord = -1

thisGrammar = ThisGrammar()
thisGrammar.initialize()

def unload():
    global thisGrammar
    if thisGrammar: thisGrammar.unload()
    thisGrammar = None
