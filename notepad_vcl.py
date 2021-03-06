# NatLink macro definitions for NaturallySpeaking
# coding: latin-1
# Generated by vcl2py 2.8.1I+, Fri Oct 31 17:34:49 2014

import natlink
from natlinkutils import *
from VocolaUtils import *


class ThisGrammar(GrammarBase):

    gramSpec = """
        <key> = ('alpha' | 'bravo' | 'charlie' | 'delta' | 'echo' | 'foxtrot' | 'golf' | 'hotel' | 'india' | 'juliett' | 'kilo' | 'lima' | 'mike' | 'november' | 'oscar' | 'papa' | 'quebec' | 'romeo' | 'sierra' | 'tango' | 'uniform' | 'victor' | 'whiskey' | 'xray' | 'yankee' | 'zulu' | '0' | '1' | '2' | '3' | '4' | '5' | '6' | '7' | '8' | '9' | '!' | '@' | '#' | '$' | '%' | '^' | '&' | '*' | '(' | ')' | '`' | '~' | '-' | '_' | '=' | '+' | '\\' | '|' | '[' | '{' | ']' | '}' | ';' | ':' | "'" | '"' | ',' | '<' | '.' | '>' | '/' | '?' | 'Left' | 'Right' | 'Up' | 'Down' | 'space-bar' | 'tab-key' | 'Enter' | 'page-up' | 'page-down' | 'Backspace' | 'delete' | 'Escape' | 'Home' | 'End' ) ;
        <1> = 'Press' <key> ;
        <2> = <key> 'Here' ;
        <3> = 'Space Bar' ;
        <4> = 'Tab Key' ;
        <special> = ('Left' | 'Right' | 'Up' | 'Down' | 'space-bar' | 'tab-key' | 'Enter' | 'page-up' | 'page-down' | 'Backspace' | 'delete' | 'escape' ) ;
        <mod> = 'Shift' | 'control-key' | 'Alt' ;
        <nn> = (1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 | 11 | 12 | 13 | 14 | 15 | 16 | 17 | 18 | 19 | 20 | 21 | 22 | 23 | 24 | 25 | 26 | 27 | 28 | 29 | 30 | 31 | 32 | 33 | 34 | 35 | 36 | 37 | 38 | 39 | 40 | 41 | 42 | 43 | 44 | 45 | 46 | 47 | 48 | 49 | 50) ;
        <5> = 'Press' <special> <nn> ;
        <14> = (('Left' | 'Right' | 'Up' | 'Down' | 'space-bar' | 'tab-key' | 'Enter' | 'page-up' | 'page-down' | 'Backspace' | 'delete' | 'escape' ) ) <nn> ;
        <6> = 'Press' <mod> <key> <nn> ;
        <15> = ('Shift' | 'control-key' | 'Alt' ) <key> <nn> ;
        <7> = 'Press' <mod> <mod> <key> <nn> ;
        <16> = ('Shift' | 'control-key' | 'Alt' ) <mod> <key> <nn> ;
        <8> = 'Press' <mod> <mod> <mod> <key> <nn> ;
        <17> = ('Shift' | 'control-key' | 'Alt' ) <mod> <mod> <key> <nn> ;
        <9> = 'Save and Close' ;
        <10> = 'Dont Save and Close' ;
        <11> = 'Save File' ;
        <12> = 'Save As' ;
        <13> = 'Replace' ;
        <any> = <1>|<2>|<3>|<4>|<5>|<14>|<6>|<15>|<7>|<16>|<8>|<17>|<9>|<10>|<11>|<12>|<13>;
        <sequence> exported = <any>;
    """
    
    def initialize(self):
        self.load(self.gramSpec)
        self.currentModule = ("","",0)
        self.ruleSet1 = ['sequence']

    def gotBegin(self,moduleInfo):
        # Return if wrong application
        window = matchWindow(moduleInfo,'notepad','')
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
                except natlink.BadWindow:
                    pass

    def convert_number_word(self, word):
        if   word == '0':
            return '0'
        else:
            return word

    def get_key(self, list_buffer, functional, word):
        if word == 'alpha':
            list_buffer += 'a'
        elif word == 'bravo':
            list_buffer += 'b'
        elif word == 'charlie':
            list_buffer += 'c'
        elif word == 'delta':
            list_buffer += 'd'
        elif word == 'echo':
            list_buffer += 'e'
        elif word == 'foxtrot':
            list_buffer += 'f'
        elif word == 'golf':
            list_buffer += 'g'
        elif word == 'hotel':
            list_buffer += 'h'
        elif word == 'india':
            list_buffer += 'i'
        elif word == 'juliett':
            list_buffer += 'j'
        elif word == 'kilo':
            list_buffer += 'k'
        elif word == 'lima':
            list_buffer += 'l'
        elif word == 'mike':
            list_buffer += 'm'
        elif word == 'november':
            list_buffer += 'n'
        elif word == 'oscar':
            list_buffer += 'o'
        elif word == 'papa':
            list_buffer += 'p'
        elif word == 'quebec':
            list_buffer += 'q'
        elif word == 'romeo':
            list_buffer += 'r'
        elif word == 'sierra':
            list_buffer += 's'
        elif word == 'tango':
            list_buffer += 't'
        elif word == 'uniform':
            list_buffer += 'u'
        elif word == 'victor':
            list_buffer += 'v'
        elif word == 'whiskey':
            list_buffer += 'w'
        elif word == 'xray':
            list_buffer += 'x'
        elif word == 'yankee':
            list_buffer += 'y'
        elif word == 'zulu':
            list_buffer += 'z'
        elif word == '0':
            list_buffer += '0'
        elif word == '1':
            list_buffer += '1'
        elif word == '2':
            list_buffer += '2'
        elif word == '3':
            list_buffer += '3'
        elif word == '4':
            list_buffer += '4'
        elif word == '5':
            list_buffer += '5'
        elif word == '6':
            list_buffer += '6'
        elif word == '7':
            list_buffer += '7'
        elif word == '8':
            list_buffer += '8'
        elif word == '9':
            list_buffer += '9'
        elif word == '!':
            list_buffer += '!'
        elif word == '@':
            list_buffer += '@'
        elif word == '#':
            list_buffer += '#'
        elif word == '$':
            list_buffer += '$'
        elif word == '%':
            list_buffer += '%'
        elif word == '^':
            list_buffer += '^'
        elif word == '&':
            list_buffer += '&'
        elif word == '*':
            list_buffer += '*'
        elif word == '(':
            list_buffer += '('
        elif word == ')':
            list_buffer += ')'
        elif word == '`':
            list_buffer += '`'
        elif word == '~':
            list_buffer += '~'
        elif word == '-':
            list_buffer += '-'
        elif word == '_':
            list_buffer += '_'
        elif word == '=':
            list_buffer += '='
        elif word == '+':
            list_buffer += '+'
        elif word == '\\':
            list_buffer += '\\'
        elif word == '|':
            list_buffer += '|'
        elif word == '[':
            list_buffer += '['
        elif word == '{':
            list_buffer += '{'
        elif word == ']':
            list_buffer += ']'
        elif word == '}':
            list_buffer += '}'
        elif word == ';':
            list_buffer += ';'
        elif word == ':':
            list_buffer += ':'
        elif word == '\'':
            list_buffer += '\''
        elif word == '"':
            list_buffer += '"'
        elif word == ',':
            list_buffer += ','
        elif word == '<':
            list_buffer += '<'
        elif word == '.':
            list_buffer += '.'
        elif word == '>':
            list_buffer += '>'
        elif word == '/':
            list_buffer += '/'
        elif word == '?':
            list_buffer += '?'
        elif word == 'Left':
            list_buffer += 'Left'
        elif word == 'Right':
            list_buffer += 'Right'
        elif word == 'Up':
            list_buffer += 'Up'
        elif word == 'Down':
            list_buffer += 'Down'
        elif word == 'space-bar':
            list_buffer += ' '
        elif word == 'tab-key':
            list_buffer += 'Tab'
        elif word == 'Enter':
            list_buffer += 'Enter'
        elif word == 'page-up':
            list_buffer += 'PgUp'
        elif word == 'page-down':
            list_buffer += 'PgDn'
        elif word == 'Backspace':
            list_buffer += 'Backspace'
        elif word == 'delete':
            list_buffer += 'Del'
        elif word == 'Escape':
            list_buffer += 'Esc'
        elif word == 'Home':
            list_buffer += 'Home'
        elif word == 'End':
            list_buffer += 'End'
        return list_buffer

    # 'Press' <key>
    def gotResults_1(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{'
            word = fullResults[1 + self.firstWord][0]
            top_buffer = self.get_key(top_buffer, False, word)
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
        except Exception, e:
            handle_error('keys.vch', 20, '\'Press\' <key>', e)
            self.firstWord = -1

    # <key> 'Here'
    def gotResults_2(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer = do_flush(False, top_buffer);
            call_Dragon('ButtonClick', 'ii', [])
            top_buffer += '{'
            word = fullResults[0 + self.firstWord][0]
            top_buffer = self.get_key(top_buffer, False, word)
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
        except Exception, e:
            handle_error('keys.vch', 21, '<key> \'Here\'', e)
            self.firstWord = -1

    # 'Space Bar'
    def gotResults_3(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += ' '
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_3(words[1:], fullResults)
        except Exception, e:
            handle_error('keys.vch', 23, '\'Space Bar\'', e)
            self.firstWord = -1

    # 'Tab Key'
    def gotResults_4(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Tab}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_4(words[1:], fullResults)
        except Exception, e:
            handle_error('keys.vch', 24, '\'Tab Key\'', e)
            self.firstWord = -1

    def get_special(self, list_buffer, functional, word):
        if word == 'Left':
            list_buffer += 'Left'
        elif word == 'Right':
            list_buffer += 'Right'
        elif word == 'Up':
            list_buffer += 'Up'
        elif word == 'Down':
            list_buffer += 'Down'
        elif word == 'space-bar':
            list_buffer += ' '
        elif word == 'tab-key':
            list_buffer += 'Tab'
        elif word == 'Enter':
            list_buffer += 'Enter'
        elif word == 'page-up':
            list_buffer += 'PgUp'
        elif word == 'page-down':
            list_buffer += 'PgDn'
        elif word == 'Backspace':
            list_buffer += 'Backspace'
        elif word == 'delete':
            list_buffer += 'Del'
        elif word == 'escape':
            list_buffer += 'Esc'
        return list_buffer

    def get_mod(self, list_buffer, functional, word):
        if word == 'Shift':
            list_buffer += 'Shift'
        elif word == 'control-key':
            list_buffer += 'Ctrl'
        elif word == 'Alt':
            list_buffer += 'Alt'
        return list_buffer

    def get_nn(self, list_buffer, functional, word):
        list_buffer += self.convert_number_word(word)
        return list_buffer

    # 'Press' <special> <nn>
    def gotResults_5(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{'
            word = fullResults[1 + self.firstWord][0]
            top_buffer = self.get_special(top_buffer, False, word)
            top_buffer += '_'
            word = fullResults[2 + self.firstWord][0]
            top_buffer = self.get_nn(top_buffer, False, word)
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 3
        except Exception, e:
            handle_error('keys.vch', 36, '\'Press\' <special> <nn>', e)
            self.firstWord = -1

    # (('Left' | 'Right' | 'Up' | 'Down' | 'space-bar' | 'tab-key' | 'Enter' | 'page-up' | 'page-down' | 'Backspace' | 'delete' | 'escape')) <nn>
    def gotResults_14(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{'
            word = fullResults[0 + self.firstWord][0]
            if word == 'Left':
                top_buffer += 'Left'
            elif word == 'Right':
                top_buffer += 'Right'
            elif word == 'Up':
                top_buffer += 'Up'
            elif word == 'Down':
                top_buffer += 'Down'
            elif word == 'space-bar':
                top_buffer += ' '
            elif word == 'tab-key':
                top_buffer += 'Tab'
            elif word == 'Enter':
                top_buffer += 'Enter'
            elif word == 'page-up':
                top_buffer += 'PgUp'
            elif word == 'page-down':
                top_buffer += 'PgDn'
            elif word == 'Backspace':
                top_buffer += 'Backspace'
            elif word == 'delete':
                top_buffer += 'Del'
            elif word == 'escape':
                top_buffer += 'Esc'
            top_buffer += '_'
            word = fullResults[1 + self.firstWord][0]
            top_buffer = self.get_nn(top_buffer, False, word)
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
        except Exception, e:
            handle_error('keys.vch', 36, '((\'Left\' | \'Right\' | \'Up\' | \'Down\' | \'space-bar\' | \'tab-key\' | \'Enter\' | \'page-up\' | \'page-down\' | \'Backspace\' | \'delete\' | \'escape\')) <nn>', e)
            self.firstWord = -1

    # 'Press' <mod> <key> <nn>
    def gotResults_6(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{'
            word = fullResults[1 + self.firstWord][0]
            top_buffer = self.get_mod(top_buffer, False, word)
            top_buffer += '+'
            word = fullResults[2 + self.firstWord][0]
            top_buffer = self.get_key(top_buffer, False, word)
            top_buffer += '_'
            word = fullResults[3 + self.firstWord][0]
            top_buffer = self.get_nn(top_buffer, False, word)
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 4
        except Exception, e:
            handle_error('keys.vch', 37, '\'Press\' <mod> <key> <nn>', e)
            self.firstWord = -1

    # ('Shift' | 'control-key' | 'Alt') <key> <nn>
    def gotResults_15(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{'
            word = fullResults[0 + self.firstWord][0]
            if word == 'Shift':
                top_buffer += 'Shift'
            elif word == 'control-key':
                top_buffer += 'Ctrl'
            elif word == 'Alt':
                top_buffer += 'Alt'
            top_buffer += '+'
            word = fullResults[1 + self.firstWord][0]
            top_buffer = self.get_key(top_buffer, False, word)
            top_buffer += '_'
            word = fullResults[2 + self.firstWord][0]
            top_buffer = self.get_nn(top_buffer, False, word)
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 3
        except Exception, e:
            handle_error('keys.vch', 37, '(\'Shift\' | \'control-key\' | \'Alt\') <key> <nn>', e)
            self.firstWord = -1

    # 'Press' <mod> <mod> <key> <nn>
    def gotResults_7(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{'
            word = fullResults[1 + self.firstWord][0]
            top_buffer = self.get_mod(top_buffer, False, word)
            top_buffer += '+'
            word = fullResults[2 + self.firstWord][0]
            top_buffer = self.get_mod(top_buffer, False, word)
            top_buffer += '+'
            word = fullResults[3 + self.firstWord][0]
            top_buffer = self.get_key(top_buffer, False, word)
            top_buffer += '_'
            word = fullResults[4 + self.firstWord][0]
            top_buffer = self.get_nn(top_buffer, False, word)
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 5
        except Exception, e:
            handle_error('keys.vch', 38, '\'Press\' <mod> <mod> <key> <nn>', e)
            self.firstWord = -1

    # ('Shift' | 'control-key' | 'Alt') <mod> <key> <nn>
    def gotResults_16(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{'
            word = fullResults[0 + self.firstWord][0]
            if word == 'Shift':
                top_buffer += 'Shift'
            elif word == 'control-key':
                top_buffer += 'Ctrl'
            elif word == 'Alt':
                top_buffer += 'Alt'
            top_buffer += '+'
            word = fullResults[1 + self.firstWord][0]
            top_buffer = self.get_mod(top_buffer, False, word)
            top_buffer += '+'
            word = fullResults[2 + self.firstWord][0]
            top_buffer = self.get_key(top_buffer, False, word)
            top_buffer += '_'
            word = fullResults[3 + self.firstWord][0]
            top_buffer = self.get_nn(top_buffer, False, word)
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 4
        except Exception, e:
            handle_error('keys.vch', 38, '(\'Shift\' | \'control-key\' | \'Alt\') <mod> <key> <nn>', e)
            self.firstWord = -1

    # 'Press' <mod> <mod> <mod> <key> <nn>
    def gotResults_8(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{'
            word = fullResults[1 + self.firstWord][0]
            top_buffer = self.get_mod(top_buffer, False, word)
            top_buffer += '+'
            word = fullResults[2 + self.firstWord][0]
            top_buffer = self.get_mod(top_buffer, False, word)
            top_buffer += '+'
            word = fullResults[3 + self.firstWord][0]
            top_buffer = self.get_mod(top_buffer, False, word)
            top_buffer += '+'
            word = fullResults[4 + self.firstWord][0]
            top_buffer = self.get_key(top_buffer, False, word)
            top_buffer += '_'
            word = fullResults[5 + self.firstWord][0]
            top_buffer = self.get_nn(top_buffer, False, word)
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 6
        except Exception, e:
            handle_error('keys.vch', 39, '\'Press\' <mod> <mod> <mod> <key> <nn>', e)
            self.firstWord = -1

    # ('Shift' | 'control-key' | 'Alt') <mod> <mod> <key> <nn>
    def gotResults_17(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{'
            word = fullResults[0 + self.firstWord][0]
            if word == 'Shift':
                top_buffer += 'Shift'
            elif word == 'control-key':
                top_buffer += 'Ctrl'
            elif word == 'Alt':
                top_buffer += 'Alt'
            top_buffer += '+'
            word = fullResults[1 + self.firstWord][0]
            top_buffer = self.get_mod(top_buffer, False, word)
            top_buffer += '+'
            word = fullResults[2 + self.firstWord][0]
            top_buffer = self.get_mod(top_buffer, False, word)
            top_buffer += '+'
            word = fullResults[3 + self.firstWord][0]
            top_buffer = self.get_key(top_buffer, False, word)
            top_buffer += '_'
            word = fullResults[4 + self.firstWord][0]
            top_buffer = self.get_nn(top_buffer, False, word)
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 5
        except Exception, e:
            handle_error('keys.vch', 39, '(\'Shift\' | \'control-key\' | \'Alt\') <mod> <mod> <key> <nn>', e)
            self.firstWord = -1

    # 'Save and Close'
    def gotResults_9(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}s{Alt+f}x'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_9(words[1:], fullResults)
        except Exception, e:
            handle_error('notepad.vcl', 4, '\'Save and Close\'', e)
            self.firstWord = -1

    # 'Dont Save and Close'
    def gotResults_10(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}xn'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_10(words[1:], fullResults)
        except Exception, e:
            handle_error('notepad.vcl', 5, '\'Dont Save and Close\'', e)
            self.firstWord = -1

    # 'Save File'
    def gotResults_11(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+s}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_11(words[1:], fullResults)
        except Exception, e:
            handle_error('notepad.vcl', 7, '\'Save File\'', e)
            self.firstWord = -1

    # 'Save As'
    def gotResults_12(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}A'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_12(words[1:], fullResults)
        except Exception, e:
            handle_error('notepad.vcl', 8, '\'Save As\'', e)
            self.firstWord = -1

    # 'Replace'
    def gotResults_13(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+h}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_13(words[1:], fullResults)
        except Exception, e:
            handle_error('notepad.vcl', 10, '\'Replace\'', e)
            self.firstWord = -1

thisGrammar = ThisGrammar()
thisGrammar.initialize()

def unload():
    global thisGrammar
    if thisGrammar: thisGrammar.unload()
    thisGrammar = None
