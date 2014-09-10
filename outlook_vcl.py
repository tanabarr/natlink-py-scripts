# NatLink macro definitions for NaturallySpeaking
# coding: latin-1
# Generated by vcl2py 2.8.1I+, Wed Sep 10 21:27:47 2014

import natlink
from natlinkutils import *
from VocolaUtils import *


class ThisGrammar(GrammarBase):

    gramSpec = """
        <folder> = 'Inbox' | 'Drafts' | 'Sent' | 'Calendar' | 'Contacts' | 'Tasks' ;
        <1> = 'Folder' <folder> ;
        <2> = 'Move To' <folder> ;
        <3> = 'New Message' ;
        <4> = 'New Message to Contact' ;
        <5> = 'Reply to' ('Message' | 'All' ) ;
        <6> = 'Forward Message' ;
        <7> = 'Flag That' ;
        <8> = 'Unflag That' ;
        <9> = 'Get Mail' ;
        <10> = 'Sort by Date' ;
        <11> = 'Sort by Sender' ;
        <12> = 'Final Message' ;
        <13> = 'Search Messages' ;
        <14> = 'Mark All Read' ;
        <15> = 'Kill Line' ;
        <16> = 'Save Attachment' (1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10) ;
        <17> = 'Save Attachment All' ;
        <18> = 'Kill Message' ;
        <19> = 'Reply Here' ;
        <20> = 'Send That' ;
        <21> = 'Plain Text' ;
        <22> = 'Fix Return' ;
        <23> = 'Plain Text' ;
        <24> = ('To' | 'See See' | 'Be See See' ) 'That' ;
        <25> = 'Calendar' ;
        <26> = 'View' ('Month' | 'Work Week' ) ;
        <27> = 'Dismiss' ;
        <28> = 'Edit That' ;
        <29> = 'Print Preview' ;
        <30> = ('Design' | 'Edit' ) 'View' ;
        <31> = 'Edit Email Address' ;
        <32> = 'Update Email Address' ;
        <33> = 'Delete Email Address' ;
        <34> = 'Copy Email Address' ;
        <35> = 'New Contact' ;
        <36> = ('Use That' | 'Save and Close' ) ;
        <37> = 'Field' ('Categories' | 'File As' | 'Business' | 'Home' | 'Mobile' | 'Address' | 'Email' ) ;
        <38> = 'Home Address' ;
        <39> = 'Swap E-mail' (2 | 3) ;
        <40> = 'Use E-mail' (2 | 3) ;
        <41> = 'Close That' ;
        <42> = 'Close' ;
        <any> = <1>|<2>;
        <sequence> exported = <any>;
        <any_set2> = <any>|<3>|<4>|<5>|<6>|<7>|<8>|<9>|<10>|<11>|<12>|<13>|<14>|<15>|<16>|<17>;
        <sequence_set2> exported = <any_set2>;
        <any_set3> = <any>|<18>;
        <sequence_set3> exported = <any_set3>;
        <any_set4> = <any>|<19>|<20>;
        <sequence_set4> exported = <any_set4>;
        <any_set5> = <any>|<21>|<22>;
        <sequence_set5> exported = <any_set5>;
        <any_set6> = <any>|<23>;
        <sequence_set6> exported = <any_set6>;
        <any_set7> = <any>|<24>;
        <sequence_set7> exported = <any_set7>;
        <any_set8> = <any>|<25>;
        <sequence_set8> exported = <any_set8>;
        <any_set9> = <any>|<26>;
        <sequence_set9> exported = <any_set9>;
        <any_set10> = <any>|<27>;
        <sequence_set10> exported = <any_set10>;
        <any_set11> = <any>|<28>|<29>|<30>|<31>|<32>|<33>|<34>|<35>;
        <sequence_set11> exported = <any_set11>;
        <any_set12> = <any>|<36>|<37>|<38>|<39>|<40>;
        <sequence_set12> exported = <any_set12>;
        <any_set13> = <any>|<41>|<42>;
        <sequence_set13> exported = <any_set13>;
    """
    
    def initialize(self):
        self.load(self.gramSpec)
        self.currentModule = ("","",0)
        self.ruleSet1 = ['sequence']
        self.ruleSet2 = ['sequence_set2']
        self.ruleSet3 = ['sequence_set3']
        self.ruleSet4 = ['sequence_set4']
        self.ruleSet5 = ['sequence_set5']
        self.ruleSet6 = ['sequence_set6']
        self.ruleSet7 = ['sequence_set7']
        self.ruleSet8 = ['sequence_set8']
        self.ruleSet9 = ['sequence_set9']
        self.ruleSet10 = ['sequence_set10']
        self.ruleSet11 = ['sequence_set11']
        self.ruleSet12 = ['sequence_set12']
        self.ruleSet13 = ['sequence_set13']

    def gotBegin(self,moduleInfo):
        # Return if wrong application
        window = matchWindow(moduleInfo,'outlook','')
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
        if string.find(title,'microsoft outlook') >= 0:
            for rule in self.ruleSet2:
                try:
                    self.activate(rule,window)
                except natlink.BadWindow:
                    pass
        if string.find(title,'message') >= 0:
            for rule in self.ruleSet3:
                try:
                    self.activate(rule,window)
                except natlink.BadWindow:
                    pass
        if string.find(title,'message (plain text)') >= 0:
            for rule in self.ruleSet4:
                try:
                    self.activate(rule,window)
                except natlink.BadWindow:
                    pass
        if string.find(title,'message (html)') >= 0:
            for rule in self.ruleSet5:
                try:
                    self.activate(rule,window)
                except natlink.BadWindow:
                    pass
        if string.find(title,'message (rich text)') >= 0:
            for rule in self.ruleSet6:
                try:
                    self.activate(rule,window)
                except natlink.BadWindow:
                    pass
        if string.find(title,'select names') >= 0:
            for rule in self.ruleSet7:
                try:
                    self.activate(rule,window)
                except natlink.BadWindow:
                    pass
        if string.find(title,'personal folders') >= 0:
            for rule in self.ruleSet8:
                try:
                    self.activate(rule,window)
                except natlink.BadWindow:
                    pass
        if string.find(title,'calendar') >= 0:
            for rule in self.ruleSet9:
                try:
                    self.activate(rule,window)
                except natlink.BadWindow:
                    pass
        if string.find(title,'overdue') >= 0:
            for rule in self.ruleSet10:
                try:
                    self.activate(rule,window)
                except natlink.BadWindow:
                    pass
        if string.find(title,'contacts') >= 0:
            for rule in self.ruleSet11:
                try:
                    self.activate(rule,window)
                except natlink.BadWindow:
                    pass
        if string.find(title,'- contact') >= 0:
            for rule in self.ruleSet12:
                try:
                    self.activate(rule,window)
                except natlink.BadWindow:
                    pass
        if string.find(title,'print preview') >= 0:
            for rule in self.ruleSet13:
                try:
                    self.activate(rule,window)
                except natlink.BadWindow:
                    pass

    def convert_number_word(self, word):
        if   word == '0':
            return '0'
        else:
            return word

    def get_folder(self, list_buffer, functional, word):
        list_buffer += word
        return list_buffer

    # 'Folder' <folder>
    def gotResults_1(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+y}{Up_40}{Left}'
            top_buffer += 'Inbox{Right}'
            top_buffer += '{Up_40}'
            word = fullResults[1 + self.firstWord][0]
            top_buffer = self.get_folder(top_buffer, False, word)
            top_buffer += '{Enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
        except Exception, e:
            handle_error('outlook.vcl', 8, '\'Folder\' <folder>', e)
            self.firstWord = -1

    # 'Move To' <folder>
    def gotResults_2(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+x}'
            top_buffer += '{Ctrl+y}{Up_40}{Left}'
            top_buffer += 'Inbox{Right}'
            top_buffer += '{Up_40}'
            word = fullResults[1 + self.firstWord][0]
            top_buffer = self.get_folder(top_buffer, False, word)
            top_buffer += '{Enter}'
            top_buffer += '{Ctrl+v}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
        except Exception, e:
            handle_error('outlook.vcl', 9, '\'Move To\' <folder>', e)
            self.firstWord = -1

    # 'New Message'
    def gotResults_3(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+n}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_3(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 14, '\'New Message\'', e)
            self.firstWord = -1

    # 'New Message to Contact'
    def gotResults_4(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+a}m'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_4(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 15, '\'New Message to Contact\'', e)
            self.firstWord = -1

    # 'Reply to' ('Message' | 'All')
    def gotResults_5(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            word = fullResults[1 + self.firstWord][0]
            if word == 'Message':
                top_buffer += '{Ctrl+r}'
            elif word == 'All':
                top_buffer += '{Ctrl+Shift+r}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_5(words[2:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 16, '\'Reply to\' (\'Message\' | \'All\')', e)
            self.firstWord = -1

    # 'Forward Message'
    def gotResults_6(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+f}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_6(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 17, '\'Forward Message\'', e)
            self.firstWord = -1

    # 'Flag That'
    def gotResults_7(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+G}{Enter}{Ctrl+q}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_7(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 18, '\'Flag That\'', e)
            self.firstWord = -1

    # 'Unflag That'
    def gotResults_8(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+G}{Alt+c}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_8(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 19, '\'Unflag That\'', e)
            self.firstWord = -1

    # 'Get Mail'
    def gotResults_9(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+t}e{Enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_9(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 20, '\'Get Mail\'', e)
            self.firstWord = -1

    # 'Sort by Date'
    def gotResults_10(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '4'
            dragon_arg2 = ''
            eval_template2_arg1 = ''
            eval_template2_arg1 += '{"n":1, "e":2, "s":3, "w":4, "ne":5, "se":6, "sw":7, "nw":8}[%a]'
            eval_template2_arg2 = ''
            eval_template2_arg2 += 'ne'
            dragon_arg2 += eval_template(eval_template2_arg1, eval_template2_arg2)
            call_Dragon('SetMousePosition', 'iii', [dragon_arg1, dragon_arg2])
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '2'
            dragon_arg2 = ''
            eval_template2_arg1 = ''
            eval_template2_arg1 += '15*%a'
            eval_template2_arg2 = ''
            eval_template2_arg2 += '-3'
            dragon_arg2 += eval_template(eval_template2_arg1, eval_template2_arg2)
            dragon_arg3 = ''
            eval_template2_arg1 = ''
            eval_template2_arg1 += '15*%a'
            eval_template2_arg2 = ''
            eval_template2_arg2 += '8'
            dragon_arg3 += eval_template(eval_template2_arg1, eval_template2_arg2)
            call_Dragon('SetMousePosition', 'iii', [dragon_arg1, dragon_arg2, dragon_arg3])
            top_buffer = do_flush(False, top_buffer);
            call_Dragon('ButtonClick', 'ii', [])
            top_buffer = do_flush(False, top_buffer);
            call_Dragon('ButtonClick', 'ii', [])
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_10(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 21, '\'Sort by Date\'', e)
            self.firstWord = -1

    # 'Sort by Sender'
    def gotResults_11(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '4'
            dragon_arg2 = ''
            eval_template2_arg1 = ''
            eval_template2_arg1 += '{"n":1, "e":2, "s":3, "w":4, "ne":5, "se":6, "sw":7, "nw":8}[%a]'
            eval_template2_arg2 = ''
            eval_template2_arg2 += 'nw'
            dragon_arg2 += eval_template(eval_template2_arg1, eval_template2_arg2)
            call_Dragon('SetMousePosition', 'iii', [dragon_arg1, dragon_arg2])
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '2'
            dragon_arg2 = ''
            eval_template2_arg1 = ''
            eval_template2_arg1 += '15*%a'
            eval_template2_arg2 = ''
            eval_template2_arg2 += '23'
            dragon_arg2 += eval_template(eval_template2_arg1, eval_template2_arg2)
            dragon_arg3 = ''
            eval_template2_arg1 = ''
            eval_template2_arg1 += '15*%a'
            eval_template2_arg2 = ''
            eval_template2_arg2 += '8'
            dragon_arg3 += eval_template(eval_template2_arg1, eval_template2_arg2)
            call_Dragon('SetMousePosition', 'iii', [dragon_arg1, dragon_arg2, dragon_arg3])
            top_buffer = do_flush(False, top_buffer);
            call_Dragon('ButtonClick', 'ii', [])
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_11(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 22, '\'Sort by Sender\'', e)
            self.firstWord = -1

    # 'Final Message'
    def gotResults_12(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{End}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_12(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 23, '\'Final Message\'', e)
            self.firstWord = -1

    # 'Search Messages'
    def gotResults_13(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+Shift+f}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_13(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 24, '\'Search Messages\'', e)
            self.firstWord = -1

    # 'Mark All Read'
    def gotResults_14(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+e}e'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_14(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 25, '\'Mark All Read\'', e)
            self.firstWord = -1

    # 'Kill Line'
    def gotResults_15(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Del}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_15(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 26, '\'Kill Line\'', e)
            self.firstWord = -1

    # 'Save Attachment' 1..10
    def gotResults_16(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}n{Down_'
            word = fullResults[1 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer += '}{Up}{Enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_16(words[2:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 27, '\'Save Attachment\' 1..10', e)
            self.firstWord = -1

    # 'Save Attachment All'
    def gotResults_17(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}n{Up}{Enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_17(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 28, '\'Save Attachment All\'', e)
            self.firstWord = -1

    # 'Kill Message'
    def gotResults_18(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+F4}'
            top_buffer += 'n'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_18(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 31, '\'Kill Message\'', e)
            self.firstWord = -1

    # 'Reply Here'
    def gotResults_19(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Home}{Shift+End}{Del}{Enter_3}{Left_2}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_19(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 33, '\'Reply Here\'', e)
            self.firstWord = -1

    # 'Send That'
    def gotResults_20(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}e{Enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_20(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 34, '\'Send That\'', e)
            self.firstWord = -1

    # 'Plain Text'
    def gotResults_21(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+o}ty'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_21(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 36, '\'Plain Text\'', e)
            self.firstWord = -1

    # 'Fix Return'
    def gotResults_22(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{End}{Del}{Shift+Enter}{Right}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_22(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 37, '\'Fix Return\'', e)
            self.firstWord = -1

    # 'Plain Text'
    def gotResults_23(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+o}ty'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_23(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 39, '\'Plain Text\'', e)
            self.firstWord = -1

    # ('To' | 'See See' | 'Be See See') 'That'
    def gotResults_24(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+'
            word = fullResults[0 + self.firstWord][0]
            if word == 'To':
                top_buffer += 'o'
            elif word == 'See See':
                top_buffer += 'c'
            elif word == 'Be See See':
                top_buffer += 'b'
            top_buffer += '}{Down}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_24(words[2:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 41, '(\'To\' | \'See See\' | \'Be See See\') \'That\'', e)
            self.firstWord = -1

    # 'Calendar'
    def gotResults_25(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+v}gc'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_25(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 44, '\'Calendar\'', e)
            self.firstWord = -1

    # 'View' ('Month' | 'Work Week')
    def gotResults_26(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+'
            word = fullResults[1 + self.firstWord][0]
            if word == 'Month':
                top_buffer += 'm'
            elif word == 'Work Week':
                top_buffer += 'r'
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_26(words[2:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 46, '\'View\' (\'Month\' | \'Work Week\')', e)
            self.firstWord = -1

    # 'Dismiss'
    def gotResults_27(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+d}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_27(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 48, '\'Dismiss\'', e)
            self.firstWord = -1

    # 'Edit That'
    def gotResults_28(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Enter}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_28(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 51, '\'Edit That\'', e)
            self.firstWord = -1

    # 'Print Preview'
    def gotResults_29(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+f}v'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_29(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 52, '\'Print Preview\'', e)
            self.firstWord = -1

    # ('Design' | 'Edit') 'View'
    def gotResults_30(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+v}vc'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_30(words[2:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 53, '(\'Design\' | \'Edit\') \'View\'', e)
            self.firstWord = -1

    # 'Edit Email Address'
    def gotResults_31(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+o}{Tab_17}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_31(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 54, '\'Edit Email Address\'', e)
            self.firstWord = -1

    # 'Update Email Address'
    def gotResults_32(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+o}{Tab_17}{Ctrl+v}{Alt+s}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_32(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 55, '\'Update Email Address\'', e)
            self.firstWord = -1

    # 'Delete Email Address'
    def gotResults_33(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+o}{Tab_17}{Del}{Alt+s}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_33(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 56, '\'Delete Email Address\'', e)
            self.firstWord = -1

    # 'Copy Email Address'
    def gotResults_34(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+o}{Tab_17}{Ctrl+c}{Esc}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_34(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 57, '\'Copy Email Address\'', e)
            self.firstWord = -1

    # 'New Contact'
    def gotResults_35(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Ctrl+n}'
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += 'Category'
            dragon_arg2 = ''
            dragon_arg2 += 'Main'
            call_Dragon('HeardWord', 'ssssssss', [dragon_arg1, dragon_arg2])
            top_buffer += '{Tab_3}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_35(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 58, '\'New Contact\'', e)
            self.firstWord = -1

    # ('Use That' | 'Save and Close')
    def gotResults_36(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+s}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_36(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 61, '(\'Use That\' | \'Save and Close\')', e)
            self.firstWord = -1

    # 'Field' ('Categories' | 'File As' | 'Business' | 'Home' | 'Mobile' | 'Address' | 'Email')
    def gotResults_37(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '4'
            dragon_arg2 = ''
            eval_template2_arg1 = ''
            eval_template2_arg1 += '{"n":1, "e":2, "s":3, "w":4, "ne":5, "se":6, "sw":7, "nw":8}[%a]'
            eval_template2_arg2 = ''
            eval_template2_arg2 += 'sw'
            dragon_arg2 += eval_template(eval_template2_arg1, eval_template2_arg2)
            call_Dragon('SetMousePosition', 'iii', [dragon_arg1, dragon_arg2])
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += '2'
            dragon_arg2 = ''
            eval_template2_arg1 = ''
            eval_template2_arg1 += '15*%a'
            eval_template2_arg2 = ''
            eval_template2_arg2 += '5'
            dragon_arg2 += eval_template(eval_template2_arg1, eval_template2_arg2)
            dragon_arg3 = ''
            eval_template2_arg1 = ''
            eval_template2_arg1 += '15*%a'
            eval_template2_arg2 = ''
            eval_template2_arg2 += '-5'
            dragon_arg3 += eval_template(eval_template2_arg1, eval_template2_arg2)
            call_Dragon('SetMousePosition', 'iii', [dragon_arg1, dragon_arg2, dragon_arg3])
            top_buffer = do_flush(False, top_buffer);
            call_Dragon('ButtonClick', 'ii', [])
            top_buffer += '{Shift+Tab_'
            word = fullResults[1 + self.firstWord][0]
            if word == 'Categories':
                top_buffer += '24'
            elif word == 'File As':
                top_buffer += '18'
            elif word == 'Business':
                top_buffer += '16'
            elif word == 'Home':
                top_buffer += '14'
            elif word == 'Mobile':
                top_buffer += '10'
            elif word == 'Address':
                top_buffer += '8'
            elif word == 'Email':
                top_buffer += '4'
            top_buffer += '}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_37(words[2:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 65, '\'Field\' (\'Categories\' | \'File As\' | \'Business\' | \'Home\' | \'Mobile\' | \'Address\' | \'Email\')', e)
            self.firstWord = -1

    # 'Home Address'
    def gotResults_38(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer = do_flush(False, top_buffer);
            dragon_arg1 = ''
            dragon_arg1 += 'Field'
            dragon_arg2 = ''
            dragon_arg2 += 'Address'
            call_Dragon('HeardWord', 'ssssssss', [dragon_arg1, dragon_arg2])
            top_buffer += '{Tab} {Down_2}{Enter}{Tab} {Shift+Tab_2}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_38(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 67, '\'Home Address\'', e)
            self.firstWord = -1

    # 'Swap E-mail' 2..3
    def gotResults_39(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Shift+Tab} {Down '
            word = fullResults[1 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer += '}{Enter}{Tab}{Ctrl+c}'
            top_buffer += '{Shift+Tab} {Down}{Enter}{Tab}{Left}{Ctrl+v}{Backspace 2}'
            top_buffer += '{Shift+Right}{Ctrl+x}'
            top_buffer += '{Shift+Tab} {Down '
            word = fullResults[1 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer += '}{Enter}{Tab}{Ctrl+v}'
            top_buffer += '{Shift+Tab} {Down}{Enter}{Tab}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_39(words[2:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 72, '\'Swap E-mail\' 2..3', e)
            self.firstWord = -1

    # 'Use E-mail' 2..3
    def gotResults_40(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Shift+Tab} {Down '
            word = fullResults[1 + self.firstWord][0]
            top_buffer += self.convert_number_word(word)
            top_buffer += '}{Enter}{Tab}{Ctrl+x}'
            top_buffer += '{Shift+Tab} {Down}{Enter}{Tab}{Ctrl+v}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 2
            if len(words) > 2: self.gotResults_40(words[2:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 74, '\'Use E-mail\' 2..3', e)
            self.firstWord = -1

    # 'Close That'
    def gotResults_41(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+c}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_41(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 77, '\'Close That\'', e)
            self.firstWord = -1

    # 'Close'
    def gotResults_42(self, words, fullResults):
        if self.firstWord<0:
            return
        try:
            top_buffer = ''
            top_buffer += '{Alt+c}'
            top_buffer = do_flush(False, top_buffer);
            self.firstWord += 1
            if len(words) > 1: self.gotResults_42(words[1:], fullResults)
        except Exception, e:
            handle_error('outlook.vcl', 77, '\'Close\'', e)
            self.firstWord = -1

thisGrammar = ThisGrammar()
thisGrammar.initialize()

def unload():
    global thisGrammar
    if thisGrammar: thisGrammar.unload()
    thisGrammar = None
