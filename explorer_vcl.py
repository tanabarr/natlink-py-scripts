# NatLink macro definitions for NaturallySpeaking
# coding: latin-1
# Generated by vcl2py 2.8, Mon Nov 18 17:39:31 2013

import natlink
from natlinkutils import *
from VocolaUtils import *


class ThisGrammar(GrammarBase):

    gramSpec = """
        <1> = 'Refresh' [ 'View' ] ;
        <2> = ('Show' | 'View' ) ('Details' | 'List' ) ;
        <3> = 'Search' ;
        <4> = 'Address' ;
        <5> = 'Left Side' ;
        <6> = 'Right Side' ;
        <7> = 'Go' ('Back' | 'Forward' ) ;
        <8> = 'Go' ('Back' | 'Forward' ) ('one' | 'two' | 'three' | 'four' | 'five' | 'six' | 'seven' | 'eight' | 'nine' | 10) ;
        <9> = ('Copy' | 'Paste' | 'Go' ) ('Address' | 'URL' ) ;
        <folder> = ('Temp' | 'Downloads' | 'Start Menu' | 'Vocola' | 'NatLink' | 'NatSpeak' ) ;
        <10> = 'Folder' <folder> ;
        <11> = 'Search' <folder> ;
        <12> = 'New Folder' ;
        <13> = 'Folders' ;
        <14> = 'Open Folder' ;
        <15> = 'Expand That' ;
        <16> = 'Collapse That' ;
        <17> = 'Share That' ;
        <18> = 'Copy Filename' ;
        <19> = 'Copy Folder Name' ;
        <20> = 'Copy Leaf Name' ;
        <21> = 'Duplicate That' ;
        <22> = 'Rename That' ;
        <23> = 'Paste Here' ;
        <24> = ('Show' | 'Edit' ) 'Properties' ;
        <25> = [ 'Toggle' ] 'Read Only' ;
        <any> = <1>|<2>|<3>|<4>|<5>|<6>|<7>|<8>|<9>|<10>|<11>|<12>|<13>|<14>|<15>|<16>|<17>|<18>|<19>|<20>|<21>|<22>|<23>|<24>|<25>;
        <sequence> exported = <any>;
    """
    
    def initialize(self):
        self.load(self.gramSpec)
        self.currentModule = ("","",0)
        self.ruleSet1 = ['sequence']

    def gotBegin(self,moduleInfo):
        # Return if wrong application
        window = matchWindow(moduleInfo,'explorer','')
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

    # Refresh [View]
    def gotResults_1(self, words, fullResults):
        if self.firstWord<0:
            return
        opt = 1 + self.firstWord
        if opt >= len(fullResults) or fullResults[opt][0] != 'View':
            fullResults.insert(opt, 'dummy')
        try:
            top_buffer = ''
            top_buffer += '{Alt+v}r'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_1(words[2:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 5, 'Refresh [View]', e)
            self.firstWord = -1

    # (Show | View) (Details | List)
    def gotResults_2(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+v}'
            word = fullResults[1 + self.firstWord][0]
            if word == 'Details':
                top_buffer += 'd'
            elif word == 'List':
                top_buffer += 'l'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_2(words[2:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 6, '(Show | View) (Details | List)', e)
            self.firstWord = -1

    # Search
    def gotResults_3(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+e}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_3(words[1:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 7, 'Search', e)
            self.firstWord = -1

    # Address
    def gotResults_4(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+d}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_4(words[1:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 8, 'Address', e)
            self.firstWord = -1

    # Left Side
    def gotResults_5(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+d}{Tab_2}{Left}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_5(words[1:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 9, 'Left Side', e)
            self.firstWord = -1

    # Right Side
    def gotResults_6(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+d}{Shift+Tab}{Left}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_6(words[1:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 10, 'Right Side', e)
            self.firstWord = -1

    # Go (Back | Forward)
    def gotResults_7(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '{Alt+'
            word = fullResults[1 + self.firstWord][0]
            if word == 'Back':
                dragon_arg1 += 'Left'
            elif word == 'Forward':
                dragon_arg1 += 'Right'
            dragon_arg1 += '}'
            call_Dragon('SendSystemKeys', 'si', [dragon_arg1])
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_7(words[2:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 11, 'Go (Back | Forward)', e)
            self.firstWord = -1

    # Go (Back | Forward) 1..10
    def gotResults_8(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '{Alt+'
            word = fullResults[1 + self.firstWord][0]
            if word == 'Back':
                dragon_arg1 += 'Left'
            elif word == 'Forward':
                dragon_arg1 += 'Right'
            dragon_arg1 += '_'
            word = fullResults[2 + self.firstWord][0]
            dragon_arg1 += self.convert_number_word(word)
            dragon_arg1 += '}'
            call_Dragon('SendSystemKeys', 'si', [dragon_arg1])
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 3
            if len(words) > 3: self.gotResults_8(words[3:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 12, 'Go (Back | Forward) 1..10', e)
            self.firstWord = -1

    # (Copy | Paste | Go) (Address | URL)
    def gotResults_9(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+d}'
            word = fullResults[0 + self.firstWord][0]
            if word == 'Copy':
                top_buffer += '{Ctrl+c}'
            elif word == 'Paste':
                top_buffer += '{Ctrl+v}'
            elif word == 'Go':
                top_buffer += ''
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_9(words[2:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 13, '(Copy | Paste | Go) (Address | URL)', e)
            self.firstWord = -1

    def get_folder(self, list_buffer, functional, word):
        if word == 'Temp':
            list_buffer += 'C:\\Temp'
        elif word == 'Downloads':
            list_buffer += 'C:\\Programs\\Downloads'
        elif word == 'Start Menu':
            list_buffer += 'C:\\Documents and Settings\\Rick Mohr\\Start Menu'
        elif word == 'Vocola':
            list_buffer += 'C:\\Programs\\NatLink\\Vocola'
        elif word == 'NatLink':
            list_buffer += 'C:\\Programs\\NatLink\\MacroSystem'
        elif word == 'NatSpeak':
            list_buffer += 'C:\\Programs\\NatSpeak\\Users\\Rick\\current'
        return list_buffer

    # Folder <folder>
    def gotResults_10(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+d}'
            word = fullResults[1 + self.firstWord][0]
            top_buffer = self.get_folder(top_buffer, False, word)
            top_buffer += '{Enter}{Tab_2}'
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '{Alt+NumKey+}'
            call_Dragon('SendSystemKeys', 'si', [dragon_arg1])
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
        except Exception, e:
            handle_error('explorer.vcl', 19, 'Folder <folder>', e)
            self.firstWord = -1

    # Search <folder>
    def gotResults_11(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+e}{Alt+l}'
            word = fullResults[1 + self.firstWord][0]
            top_buffer = self.get_folder(top_buffer, False, word)
            top_buffer += '{Alt+m}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
        except Exception, e:
            handle_error('explorer.vcl', 20, 'Search <folder>', e)
            self.firstWord = -1

    # New Folder
    def gotResults_12(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}wf'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_12(words[1:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 22, 'New Folder', e)
            self.firstWord = -1

    # Folders
    def gotResults_13(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+v}eo'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_13(words[1:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 23, 'Folders', e)
            self.firstWord = -1

    # Open Folder
    def gotResults_14(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}{Enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_14(words[1:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 24, 'Open Folder', e)
            self.firstWord = -1

    # Expand That
    def gotResults_15(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '{Alt+NumKey+}'
            call_Dragon('SendSystemKeys', 'si', [dragon_arg1])
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_15(words[1:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 25, 'Expand That', e)
            self.firstWord = -1

    # Collapse That
    def gotResults_16(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '{Alt+NumKey-}'
            call_Dragon('SendSystemKeys', 'si', [dragon_arg1])
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_16(words[1:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 26, 'Collapse That', e)
            self.firstWord = -1

    # Share That
    def gotResults_17(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}r'
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '1000'
            call_Dragon('Wait', 'i', [dragon_arg1])
            top_buffer += '{Tab_5}{Right_2}{Alt+s}{Enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_17(words[1:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 27, 'Share That', e)
            self.firstWord = -1

    # Copy Filename
    def gotResults_18(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}m{Ctrl+c}{Alt+d}{Right}\\{Ctrl+v}'
            top_buffer += '{Home}{Shift+End}{Ctrl+c}{Esc}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_18(words[1:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 32, 'Copy Filename', e)
            self.firstWord = -1

    # Copy Folder Name
    def gotResults_19(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+d}{Ctrl+c}{Esc}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_19(words[1:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 33, 'Copy Folder Name', e)
            self.firstWord = -1

    # Copy Leaf Name
    def gotResults_20(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}m{Ctrl+c}{Esc}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_20(words[1:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 34, 'Copy Leaf Name', e)
            self.firstWord = -1

    # Duplicate That
    def gotResults_21(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+c}{Left}{Ctrl+v}c'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_21(words[1:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 36, 'Duplicate That', e)
            self.firstWord = -1

    # Rename That
    def gotResults_22(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{F2}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_22(words[1:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 37, 'Rename That', e)
            self.firstWord = -1

    # Paste Here
    def gotResults_23(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '1'
            dragon_arg2 = ''
            dragon_arg2 += '1'
            call_Dragon('ButtonClick', 'ii', [dragon_arg1, dragon_arg2])
            top_buffer += '{Ctrl+v}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_23(words[1:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 38, 'Paste Here', e)
            self.firstWord = -1

    # (Show | Edit) Properties
    def gotResults_24(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}r'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_24(words[2:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 40, '(Show | Edit) Properties', e)
            self.firstWord = -1

    # [Toggle] Read Only
    def gotResults_25(self, words, fullResults):
        if self.firstWord<0:
            return
        opt = 0 + self.firstWord
        if opt >= len(fullResults) or fullResults[opt][0] != 'Toggle':
            fullResults.insert(opt, 'dummy')
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}r'
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '2000'
            call_Dragon('Wait', 'i', [dragon_arg1])
            top_buffer += '{Alt+r}{Enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_25(words[2:], fullResults)
        except Exception, e:
            handle_error('explorer.vcl', 41, '[Toggle] Read Only', e)
            self.firstWord = -1

thisGrammar = ThisGrammar()
thisGrammar.initialize()

def unload():
    global thisGrammar
    if thisGrammar: thisGrammar.unload()
    thisGrammar = None
