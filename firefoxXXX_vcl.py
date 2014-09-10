# NatLink macro definitions for NaturallySpeaking
# coding: latin-1
# Generated by vcl2py 2.8.1I+, Wed Sep 10 23:17:36 2014

import natlink
from natlinkutils import *
from VocolaUtils import *


class ThisGrammar(GrammarBase):

    gramSpec = """
        <dgndictation> imported;
        <n> = ('zero' | 'one' | 'two' | 'three' | 'four' | 'five' | 'six' | 'seven' | 'eight' | 'nine') ;
        <1> = 'zoom in' ;
        <2> = 'zoom out' ;
        <3> = 'save' ;
        <4> = 'new' ;
        <5> = 'new window' ;
        <6> = 'next' <n> ;
        <7> = 'previous' <n> ;
        <8> = 'switch tab' <n> ;
        <9> = 'private' ;
        <10> = 'close' ;
        <11> = 'bookmark' ;
        <12> = 'show bookmarks' ;
        <13> = 'tools' ;
        <14> = 'options' ;
        <15> = 'reload' ;
        <16> = 'back page' ;
        <17> = ('Copy' | 'Paste' | 'Go' ) ('Address' | 'URL' ) ;
        <pick> = ('pick' | 'drop pick' | 'go pick' | 'push pick' | 'tab pick' | 'window pick' | 'menu pick' | 'save pick' | 'copy pick' ) ;
        <18> = <pick> ('zero' | 'one' | 'two' | 'three' | 'four' | 'five' | 'six' | 'seven' | 'eight' | 'nine') ;
        <19> = <pick> ('zero' | 'one' | 'two' | 'three' | 'four' | 'five' | 'six' | 'seven' | 'eight' | 'nine') ('zero' | 'one' | 'two' | 'three' | 'four' | 'five' | 'six' | 'seven' | 'eight' | 'nine') ;
        <20> = <pick> ('zero' | 'one' | 'two' | 'three' | 'four' | 'five' | 'six' | 'seven' | 'eight' | 'nine') ('zero' | 'one' | 'two' | 'three' | 'four' | 'five' | 'six' | 'seven' | 'eight' | 'nine') ('zero' | 'one' | 'two' | 'three' | 'four' | 'five' | 'six' | 'seven' | 'eight' | 'nine') ;
        <21> = <pick> ('zero' | 'one' | 'two' | 'three' | 'four' | 'five' | 'six' | 'seven' | 'eight' | 'nine') ('zero' | 'one' | 'two' | 'three' | 'four' | 'five' | 'six' | 'seven' | 'eight' | 'nine') ('zero' | 'one' | 'two' | 'three' | 'four' | 'five' | 'six' | 'seven' | 'eight' | 'nine') ('zero' | 'one' | 'two' | 'three' | 'four' | 'five' | 'six' | 'seven' | 'eight' | 'nine') ;
        <22> = 'show numbers' ;
        <23> = 'refresh numbers' ;
        <24> = 'blur me' ;
        <25> = 'link' <dgndictation> ;
        <26> = 'new link' <dgndictation> ;
        <27> = 'window link' <dgndictation> ;
        <any> = <1>|<2>|<3>|<4>|<5>|<6>|<7>|<8>|<9>|<10>|<11>|<12>|<13>|<14>|<15>|<16>|<17>|<18>|<19>|<20>|<21>|<22>|<23>|<24>|<25>|<26>|<27>;
        <sequence> exported = <any>;
    """
    
    def initialize(self):
        self.load(self.gramSpec)
        self.currentModule = ("","",0)
        self.ruleSet1 = ['sequence']

    def gotBegin(self,moduleInfo):
        # Return if wrong application
        window = matchWindow(moduleInfo,'firefoxxxx','')
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

    # 'zoom in'
    def gotResults_1(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+plus}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_1(words[1:], fullResults)
        except Exception, e:
            handle_error('firefoxXXX.vcl', 6, '\'zoom in\'', e)
            self.firstWord = -1

    # 'zoom out'
    def gotResults_2(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+minus}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_2(words[1:], fullResults)
        except Exception, e:
            handle_error('firefoxXXX.vcl', 7, '\'zoom out\'', e)
            self.firstWord = -1

    # 'save'
    def gotResults_3(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+s}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_3(words[1:], fullResults)
        except Exception, e:
            handle_error('firefoxXXX.vcl', 8, '\'save\'', e)
            self.firstWord = -1

    # 'new'
    def gotResults_4(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+t}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_4(words[1:], fullResults)
        except Exception, e:
            handle_error('firefoxXXX.vcl', 9, '\'new\'', e)
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
            handle_error('firefoxXXX.vcl', 11, '\'new window\'', e)
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
            handle_error('firefoxXXX.vcl', 12, '\'next\' <n>', e)
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
            handle_error('firefoxXXX.vcl', 13, '\'previous\' <n>', e)
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
            handle_error('firefoxXXX.vcl', 14, '\'switch tab\' <n>', e)
            self.firstWord = -1

    # 'private'
    def gotResults_9(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+Shift+p}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_9(words[1:], fullResults)
        except Exception, e:
            handle_error('firefoxXXX.vcl', 15, '\'private\'', e)
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
            handle_error('firefoxXXX.vcl', 17, '\'close\'', e)
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
            handle_error('firefoxXXX.vcl', 18, '\'bookmark\'', e)
            self.firstWord = -1

    # 'show bookmarks'
    def gotResults_12(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+Shift+b}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_12(words[1:], fullResults)
        except Exception, e:
            handle_error('firefoxXXX.vcl', 19, '\'show bookmarks\'', e)
            self.firstWord = -1

    # 'tools'
    def gotResults_13(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+e}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_13(words[1:], fullResults)
        except Exception, e:
            handle_error('firefoxXXX.vcl', 20, '\'tools\'', e)
            self.firstWord = -1

    # 'options'
    def gotResults_14(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+t}'
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '100'
            call_Dragon('Wait', 'i', [dragon_arg1])
            top_buffer += '{o}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_14(words[1:], fullResults)
        except Exception, e:
            handle_error('firefoxXXX.vcl', 21, '\'options\'', e)
            self.firstWord = -1

    # 'reload'
    def gotResults_15(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{f5}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_15(words[1:], fullResults)
        except Exception, e:
            handle_error('firefoxXXX.vcl', 22, '\'reload\'', e)
            self.firstWord = -1

    # 'back page'
    def gotResults_16(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{backspace}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_16(words[1:], fullResults)
        except Exception, e:
            handle_error('firefoxXXX.vcl', 23, '\'back page\'', e)
            self.firstWord = -1

    # ('Copy' | 'Paste' | 'Go') ('Address' | 'URL')
    def gotResults_17(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+d}'
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '200'
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
            if len(words) > 2: self.gotResults_17(words[2:], fullResults)
        except Exception, e:
            handle_error('firefoxXXX.vcl', 24, '(\'Copy\' | \'Paste\' | \'Go\') (\'Address\' | \'URL\')', e)
            self.firstWord = -1

    def get_pick(self, list_buffer, functional, word):
        if word == 'pick':
            list_buffer += '{shift}{enter}'
        elif word == 'drop pick':
            list_buffer += '{shift}{enter}{alt+down}'
        elif word == 'go pick':
            list_buffer += '{shift}'
        elif word == 'push pick':
            list_buffer += '{shift}{ctrl+enter}'
        elif word == 'tab pick':
            list_buffer += '{shift}{ctrl+shift+enter}'
        elif word == 'window pick':
            list_buffer += '{shift}{shift+enter}'
        elif word == 'menu pick':
            list_buffer += '{shift}{shift+f10}'
        elif word == 'save pick':
            list_buffer += '{shift}{shift+f10}'
            list_buffer = do_flush(functional, list_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '100'
            call_Dragon('Wait', 'i', [dragon_arg1])
            list_buffer += 'k'
        elif word == 'copy pick':
            list_buffer += '{shift}{shift+f10}'
            list_buffer = do_flush(functional, list_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '100'
            call_Dragon('Wait', 'i', [dragon_arg1])
            list_buffer += 'a'
        return list_buffer

    # <pick> 0..9
    def gotResults_18(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{ctrl+NumKey/}'
            top_buffer += '{alt+ctrl+'
            word = fullResults[1 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer += '}'
            word = fullResults[0 + self.firstWord][0]
            top_buffer = self.get_pick(top_buffer, False, word)
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
        except Exception, e:
            handle_error('firefoxXXX.vcl', 74, '<pick> 0..9', e)
            self.firstWord = -1

    # <pick> 0..9 0..9
    def gotResults_19(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{ctrl+NumKey/}'
            top_buffer += '{alt+ctrl+'
            word = fullResults[1 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer += '}{alt+ctrl+'
            word = fullResults[2 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer += '}'
            word = fullResults[0 + self.firstWord][0]
            top_buffer = self.get_pick(top_buffer, False, word)
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 3
        except Exception, e:
            handle_error('firefoxXXX.vcl', 75, '<pick> 0..9 0..9', e)
            self.firstWord = -1

    # <pick> 0..9 0..9 0..9
    def gotResults_20(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{ctrl+NumKey/}'
            top_buffer += '{alt+ctrl+'
            word = fullResults[1 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer += '}{alt+ctrl+'
            word = fullResults[2 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer += '}{alt+ctrl+'
            word = fullResults[3 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer += '}'
            word = fullResults[0 + self.firstWord][0]
            top_buffer = self.get_pick(top_buffer, False, word)
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 4
        except Exception, e:
            handle_error('firefoxXXX.vcl', 76, '<pick> 0..9 0..9 0..9', e)
            self.firstWord = -1

    # <pick> 0..9 0..9 0..9 0..9
    def gotResults_21(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{ctrl+NumKey/}'
            top_buffer += '{alt+ctrl+'
            word = fullResults[1 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer += '}{alt+ctrl+'
            word = fullResults[2 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer += '}{alt+ctrl+'
            word = fullResults[3 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer += '}'
            top_buffer += '{alt+ctrl+'
            word = fullResults[4 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer += '}'
            word = fullResults[0 + self.firstWord][0]
            top_buffer = self.get_pick(top_buffer, False, word)
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 5
        except Exception, e:
            handle_error('firefoxXXX.vcl', 78, '<pick> 0..9 0..9 0..9 0..9', e)
            self.firstWord = -1

    # 'show numbers'
    def gotResults_22(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{ctrl+NumKey/}'
            top_buffer += '{NumKey.}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_22(words[1:], fullResults)
        except Exception, e:
            handle_error('firefoxXXX.vcl', 80, '\'show numbers\'', e)
            self.firstWord = -1

    # 'refresh numbers'
    def gotResults_23(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{ctrl+NumKey/}'
            top_buffer += '{NumKey.}'
            top_buffer += '{NumKey.}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_23(words[1:], fullResults)
        except Exception, e:
            handle_error('firefoxXXX.vcl', 81, '\'refresh numbers\'', e)
            self.firstWord = -1

    # 'blur me'
    def gotResults_24(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{ctrl+NumKey/}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_24(words[1:], fullResults)
        except Exception, e:
            handle_error('firefoxXXX.vcl', 82, '\'blur me\'', e)
            self.firstWord = -1

    # 'link' <_anything>
    def gotResults_25(self, words, fullResults):
        if self.firstWord<0:
            return
        fullResults = combineDictationWords(fullResults)
        opt = 1 + self.firstWord
        if opt >= len(fullResults) or fullResults[opt][1] != 'converted dgndictation':
            fullResults.insert(opt, ['', 'converted dgndictation'])
        try:
            top_buffer = ''
            top_buffer += '{ctrl+NumKey/}'
            top_buffer += '\''
            word = fullResults[1 + self.firstWord][0]
            top_buffer += word
            top_buffer += '{enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
        except Exception, e:
            handle_error('firefoxXXX.vcl', 89, '\'link\' <_anything>', e)
            self.firstWord = -1

    # 'new link' <_anything>
    def gotResults_26(self, words, fullResults):
        if self.firstWord<0:
            return
        fullResults = combineDictationWords(fullResults)
        opt = 1 + self.firstWord
        if opt >= len(fullResults) or fullResults[opt][1] != 'converted dgndictation':
            fullResults.insert(opt, ['', 'converted dgndictation'])
        try:
            top_buffer = ''
            top_buffer += '{ctrl+NumKey/}'
            top_buffer += '\''
            word = fullResults[1 + self.firstWord][0]
            top_buffer += word
            top_buffer += '{ctrl+enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
        except Exception, e:
            handle_error('firefoxXXX.vcl', 90, '\'new link\' <_anything>', e)
            self.firstWord = -1

    # 'window link' <_anything>
    def gotResults_27(self, words, fullResults):
        if self.firstWord<0:
            return
        fullResults = combineDictationWords(fullResults)
        opt = 1 + self.firstWord
        if opt >= len(fullResults) or fullResults[opt][1] != 'converted dgndictation':
            fullResults.insert(opt, ['', 'converted dgndictation'])
        try:
            top_buffer = ''
            top_buffer += '{ctrl+NumKey/}'
            top_buffer += '\''
            word = fullResults[1 + self.firstWord][0]
            top_buffer += word
            top_buffer += '{shift+enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
        except Exception, e:
            handle_error('firefoxXXX.vcl', 91, '\'window link\' <_anything>', e)
            self.firstWord = -1

thisGrammar = ThisGrammar()
thisGrammar.initialize()

def unload():
    global thisGrammar
    if thisGrammar: thisGrammar.unload()
    thisGrammar = None
