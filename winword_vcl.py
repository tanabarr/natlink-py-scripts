# NatLink macro definitions for NaturallySpeaking
# coding: latin-1
# Generated by vcl2py 2.8.1I+, Fri Oct 31 17:34:50 2014

import natlink
from natlinkutils import *
from VocolaUtils import *


class ThisGrammar(GrammarBase):

    gramSpec = """
        <1> = 'Save File' ;
        <2> = 'Save As' ;
        <3> = 'Recent Files' ;
        <4> = 'Recent File' (1 | 2 | 3 | 4) ;
        <5> = ('Outline' | 'Normal' | 'Print Layout' | 'Web Layout' ) 'View' ;
        <6> = 'New Item' ;
        <7> = 'New Item Promote' ;
        <8> = 'New Item Demote' ;
        <9> = 'Promote That' ;
        <10> = 'Demote That' ;
        <11> = 'Expand That' ;
        <12> = 'Collapse That' ;
        <13> = 'Item' ('Up' | 'Down' ) ;
        <14> = 'Item' ('Up' | 'Down' ) (1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 | 11 | 12 | 13 | 14 | 15 | 16 | 17 | 18 | 19 | 20) ;
        <15> = 'Style' ('Normal' | 'None' ) ;
        <16> = 'Heading' (1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9) ;
        <17> = 'Edit Style' ;
        <18> = 'Insert' ('Row Below' | 'Row Above' | 'Column Left' | 'Column Right' ) ;
        <19> = 'Merge Cells' ;
        <20> = 'Merge Row' ;
        <21> = 'Align' ('Top' | 'Center' | 'Bottom' ) ('Left' | 'Center' | 'Right' ) ;
        <22> = 'Find That' ;
        <23> = 'Replace Text' ;
        <24> = 'Keep with Next' ;
        <25> = ('Subscript' | 'Superscript' ) 'That' ;
        <26> = 'En Dash' ;
        <27> = 'Em Dash' ;
        <28> = 'Kill Here' ;
        <29> = 'Kill Back Here' ;
        <30> = 'Rosy Find' ;
        <31> = 'Rosy Do' ;
        <32> = 'Accept' ;
        <33> = 'Reject' ;
        <34> = 'Find' ;
        <35> = 'Replace' ;
        <36> = 'Replace All' ;
        <37> = 'Find Next' ;
        <38> = 'Find' ;
        <any> = <1>|<2>|<3>|<4>|<5>|<6>|<7>|<8>|<9>|<10>|<11>|<12>|<13>|<14>|<15>|<16>|<17>|<18>|<19>|<20>|<21>|<22>|<23>|<24>|<25>|<26>|<27>|<28>|<29>|<30>|<31>;
        <sequence> exported = <any>;
        <any_set2> = <any>|<32>|<33>|<34>;
        <sequence_set2> exported = <any_set2>;
        <any_set3> = <any>|<35>|<36>|<37>|<38>;
        <sequence_set3> exported = <any_set3>;
    """
    
    def initialize(self):
        self.load(self.gramSpec)
        self.currentModule = ("","",0)
        self.ruleSet1 = ['sequence']
        self.ruleSet2 = ['sequence_set2']
        self.ruleSet3 = ['sequence_set3']

    def gotBegin(self,moduleInfo):
        # Return if wrong application
        window = matchWindow(moduleInfo,'winword','')
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
        if string.find(title,'accept or reject changes') >= 0:
            for rule in self.ruleSet2:
                try:
                    self.activate(rule,window)
                except natlink.BadWindow:
                    pass
        if string.find(title,'find and replace') >= 0:
            for rule in self.ruleSet3:
                try:
                    self.activate(rule,window)
                except natlink.BadWindow:
                    pass

    def convert_number_word(self, word):
        if   word == '0':
            return '0'
        else:
            return word

    # 'Save File'
    def gotResults_1(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+s}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_1(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 3, '\'Save File\'', e)
            self.firstWord = -1

    # 'Save As'
    def gotResults_2(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}A'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_2(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 4, '\'Save As\'', e)
            self.firstWord = -1

    # 'Recent Files'
    def gotResults_3(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_3(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 5, '\'Recent Files\'', e)
            self.firstWord = -1

    # 'Recent File' 1..4
    def gotResults_4(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}'
            word = fullResults[1 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_4(words[2:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 6, '\'Recent File\' 1..4', e)
            self.firstWord = -1

    # ('Outline' | 'Normal' | 'Print Layout' | 'Web Layout') 'View'
    def gotResults_5(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+v}'
            word = fullResults[0 + self.firstWord][0]
            if word == 'Outline':
                top_buffer += 'o'
            elif word == 'Normal':
                top_buffer += 'n'
            elif word == 'Print Layout':
                top_buffer += 'p'
            elif word == 'Web Layout':
                top_buffer += 'w'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_5(words[2:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 8, '(\'Outline\' | \'Normal\' | \'Print Layout\' | \'Web Layout\') \'View\'', e)
            self.firstWord = -1

    # 'New Item'
    def gotResults_6(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+Down}{Left_2}{End}{Enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_6(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 13, '\'New Item\'', e)
            self.firstWord = -1

    # 'New Item Promote'
    def gotResults_7(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+Down}{Left_2}{End}{Enter}{Alt+Shift+Left}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_7(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 14, '\'New Item Promote\'', e)
            self.firstWord = -1

    # 'New Item Demote'
    def gotResults_8(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+Down}{Left_2}{End}{Enter}{Alt+Shift+Right}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_8(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 15, '\'New Item Demote\'', e)
            self.firstWord = -1

    # 'Promote That'
    def gotResults_9(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+Shift+Left}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_9(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 17, '\'Promote That\'', e)
            self.firstWord = -1

    # 'Demote That'
    def gotResults_10(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+Shift+Right}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_10(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 18, '\'Demote That\'', e)
            self.firstWord = -1

    # 'Expand That'
    def gotResults_11(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+Shift+Plus}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_11(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 19, '\'Expand That\'', e)
            self.firstWord = -1

    # 'Collapse That'
    def gotResults_12(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+Shift+Minus}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_12(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 20, '\'Collapse That\'', e)
            self.firstWord = -1

    # 'Item' ('Up' | 'Down')
    def gotResults_13(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+Shift+'
            word = fullResults[1 + self.firstWord][0]
            top_buffer += word
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_13(words[2:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 22, '\'Item\' (\'Up\' | \'Down\')', e)
            self.firstWord = -1

    # 'Item' ('Up' | 'Down') 1..20
    def gotResults_14(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+Shift+'
            word = fullResults[1 + self.firstWord][0]
            top_buffer += word
            top_buffer += ' '
            word = fullResults[2 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 3
            if len(words) > 3: self.gotResults_14(words[3:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 23, '\'Item\' (\'Up\' | \'Down\') 1..20', e)
            self.firstWord = -1

    # 'Style' ('Normal' | 'None')
    def gotResults_15(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+S}'
            word = fullResults[1 + self.firstWord][0]
            if word == 'Normal':
                top_buffer += 'Normal'
            elif word == 'None':
                top_buffer += 'Default Paragraph Font'
            top_buffer += '{Enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_15(words[2:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 30, '\'Style\' (\'Normal\' | \'None\')', e)
            self.firstWord = -1

    # 'Heading' 1..9
    def gotResults_16(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '{Ctrl+S}'
            call_Dragon('SendSystemKeys', 'si', [dragon_arg1])
            top_buffer += 'Heading '
            word = fullResults[1 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer += '{Enter}{Down}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_16(words[2:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 32, '\'Heading\' 1..9', e)
            self.firstWord = -1

    # 'Edit Style'
    def gotResults_17(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+o}'
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '100'
            call_Dragon('Wait', 'i', [dragon_arg1])
            top_buffer += 's{Alt+m}{Alt+o}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_17(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 34, '\'Edit Style\'', e)
            self.firstWord = -1

    # 'Insert' ('Row Below' | 'Row Above' | 'Column Left' | 'Column Right')
    def gotResults_18(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+a}i'
            word = fullResults[1 + self.firstWord][0]
            if word == 'Row Below':
                top_buffer += 'b'
            elif word == 'Row Above':
                top_buffer += 'a'
            elif word == 'Column Left':
                top_buffer += 'l'
            elif word == 'Column Right':
                top_buffer += 'r'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_18(words[2:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 39, '\'Insert\' (\'Row Below\' | \'Row Above\' | \'Column Left\' | \'Column Right\')', e)
            self.firstWord = -1

    # 'Merge Cells'
    def gotResults_19(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+a}m'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_19(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 41, '\'Merge Cells\'', e)
            self.firstWord = -1

    # 'Merge Row'
    def gotResults_20(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+Home}{Shift+Alt+End}{Alt+a}m'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_20(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 42, '\'Merge Row\'', e)
            self.firstWord = -1

    # 'Align' ('Top' | 'Center' | 'Bottom') ('Left' | 'Center' | 'Right')
    def gotResults_21(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+a}r{Alt+e}{Alt+'
            word = fullResults[1 + self.firstWord][0]
            if word == 'Top':
                top_buffer += 'p'
            elif word == 'Center':
                top_buffer += 'c'
            elif word == 'Bottom':
                top_buffer += 'b'
            top_buffer += '}{Enter}{Alt+o}p{Alt+i}{Alt+g}'
            word = fullResults[2 + self.firstWord][0]
            if word == 'Left':
                top_buffer += 'l'
            elif word == 'Center':
                top_buffer += 'c'
            elif word == 'Right':
                top_buffer += 'r'
            top_buffer += '{Enter}{Enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 3
            if len(words) > 3: self.gotResults_21(words[3:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 45, '\'Align\' (\'Top\' | \'Center\' | \'Bottom\') (\'Left\' | \'Center\' | \'Right\')', e)
            self.firstWord = -1

    # 'Find That'
    def gotResults_22(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+c}{Ctrl+f}{Ctrl+v}{Enter}{Esc}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_22(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 50, '\'Find That\'', e)
            self.firstWord = -1

    # 'Replace Text'
    def gotResults_23(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+e}e'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_23(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 51, '\'Replace Text\'', e)
            self.firstWord = -1

    # 'Keep with Next'
    def gotResults_24(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+o}p{Alt+p}{Alt+x}{Enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_24(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 53, '\'Keep with Next\'', e)
            self.firstWord = -1

    # ('Subscript' | 'Superscript') 'That'
    def gotResults_25(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+o}f{Alt+'
            word = fullResults[0 + self.firstWord][0]
            if word == 'Subscript':
                top_buffer += 'b'
            elif word == 'Superscript':
                top_buffer += 'p'
            top_buffer += '}{Enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_25(words[2:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 55, '(\'Subscript\' | \'Superscript\') \'That\'', e)
            self.firstWord = -1

    # 'En Dash'
    def gotResults_26(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+NumKey-}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_26(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 56, '\'En Dash\'', e)
            self.firstWord = -1

    # 'Em Dash'
    def gotResults_27(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+Alt+NumKey-}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_27(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 57, '\'Em Dash\'', e)
            self.firstWord = -1

    # 'Kill Here'
    def gotResults_28(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Shift+End}{Shift+Left}{Del}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_28(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 59, '\'Kill Here\'', e)
            self.firstWord = -1

    # 'Kill Back Here'
    def gotResults_29(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Shift+Home}{Shift+Right}{Del}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_29(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 60, '\'Kill Back Here\'', e)
            self.firstWord = -1

    # 'Rosy Find'
    def gotResults_30(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+f}{Enter}{Esc}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_30(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 62, '\'Rosy Find\'', e)
            self.firstWord = -1

    # 'Rosy Do'
    def gotResults_31(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Home}{Shift+Down_2}{Del}{Ctrl+S}CodeHeading{Enter}{Down}{Del_13}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_31(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 63, '\'Rosy Do\'', e)
            self.firstWord = -1

    # 'Accept'
    def gotResults_32(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+a}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_32(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 66, '\'Accept\'', e)
            self.firstWord = -1

    # 'Reject'
    def gotResults_33(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+r}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_33(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 67, '\'Reject\'', e)
            self.firstWord = -1

    # 'Find'
    def gotResults_34(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_34(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 68, '\'Find\'', e)
            self.firstWord = -1

    # 'Replace'
    def gotResults_35(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+r}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_35(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 71, '\'Replace\'', e)
            self.firstWord = -1

    # 'Replace All'
    def gotResults_36(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+a}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_36(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 72, '\'Replace All\'', e)
            self.firstWord = -1

    # 'Find Next'
    def gotResults_37(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_37(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 73, '\'Find Next\'', e)
            self.firstWord = -1

    # 'Find'
    def gotResults_38(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_38(words[1:], fullResults)
        except Exception, e:
            handle_error('winword.vcl', 73, '\'Find\'', e)
            self.firstWord = -1

thisGrammar = ThisGrammar()
thisGrammar.initialize()

def unload():
    global thisGrammar
    if thisGrammar: thisGrammar.unload()
    thisGrammar = None
