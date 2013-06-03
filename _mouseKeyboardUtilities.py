#
# Python Macro Language for Dragon NaturallySpeaking
# nextApp
#

import natlink
from natlinkutils import *
import win32gui as wg

class ThisGrammar(GrammarBase):

    # Macros can be repeated with recognitionmimic function which takes list of
    # words as parameter
    abrvMap = {'normal': 'switch to normal mode', 'spell':
               'switch to spell mode', 'escape': 'press escape',
               'insert': 'press insert','hash': 'press hash'}

    # playstring function of nat link uses format:
    # playstring(<keystring>, modifier flags: <ctrl,alt,shift>(bitwise 3 LSBs)
    # modifier applies to first character in string. More information in 'natlink.txt'
    kbMacros = {# Global commands
                'downshift': ('{down}',0x01), 'rightshift': ('{right}',0x01),
                'leftshift': ('{left}',0x01), 'upshift': ('{up}',0x01),
                #'select': ('{ctrl+shift}',0x00),#c-s-click requires playevents
                # Google Chrome commands
                'next': ('{ctrl+tab}',0x00),
                'previous': ('{ctrl+shift+tab}',0x00),
                'private': ('N',0x05), 'new': ('n',0x04),
                'close': ('w',0x04), 'flag': ('{alt}aa',0x00),
                'save': ('s',0x04), 'bookmark': ('b',0x04),
                'prompt': ('{space}c',0x02),# 'prompt 'closes command prompt
                'history': ('r',0x04), 'zoom in': ('+',0x04),
                'zoom out': ('-',0x04),
                # vim commands
                'vim format': ('Q',0x00), 'vim undo': ('u',0x00),
                'vim redo': ('{ctrl+r}',0x00), 'vim next': (':bn',0x00),
                'vim previous': (':bp',0x00), 'vim save': (':w',0x00),
                'vim quit': (':q',0x00), 'vim taglist': ('{ctrl+p}',0x00),
                'vim update': (':!ctags -a .',0x00),
                'vim folds': ('{ctrl+f}',0x00),
                'vim remove': (':bd',0x00), #remove buffer
                'vim jump back': ('g,',0x00),
                'vim command last': (':{up},',0x00),
                #screen commands
                'attach screen ': ('screen -R{enter}',0x00),
                'screen previous': ('p',0x00), 'screen next': ('n',0x00),
                'screen help': ('?',0x00), 'screen new': ('c',0x00),
                'screen detach': ('d',0x00), 'screen list': ('"',0x00),
                'screen kill': ('k',0x00), 'screen title': ('A',0x00),
                # window split related
                'screen switch': ('{tab}',0x00), 'screen split': ('S',0x00),
                'screen vertical': ('|',0x00), 'screen crop': ('Q',0x00),
                'screen remove': ('X',0x00),
                }

    # Todo: embed this list of strings within grammar to save space
    # mapping of keyboard keys to virtual key code to send as key input
    # VK_SPACE,VK_UP,VK_DOWN,VK_LEFT,VK_RIGHT,VK_RETURN,VK_BACK
    kmap = {'space': 0x20, 'up': 0x26, 'down': 0x28, 'left': 0x25, 'right': 0x27,
            'enter': 0x0d, 'backspace': 0x08, 'delete': 0x2e, 'leftclick': 0x201,
            'rightclick': 0x204, 'doubleclick': 0x202}

    nullTitles = ['Default IME', 'MSCTFIME UI', 'Engine Window',
                  'VDct Notifier Window', 'Program Manager',
                  'Spelling Window', 'Start']

    gramSpec = """
        <quickStart> exported = QuickStart (left|right|double) row ({3}) column ({3});
        <repeatKey> exported = repeat key ({2}) ({4}|{5});
        <windowFocus> exported = focus [on] window ({4}) [from bottom];
        <androidSC> exported =  show coordinates and screen size;
        <abrvPhrase> exported = ({1}) [mode];
        <kbMacro> exported = ({0});
    """.format('|'.join(kbMacros.keys()),'|'.join(abrvMap.keys()),
               '|'.join(kmap.keys()),
                str(range(6)).strip('[]').replace(', ','|'),
                str(range(20)).strip('[]').replace(', ','|'),
                str(range(20,50,10)).strip('[]').replace(', ','|'))

    def gotResults_kbMacro(self, words, fullResults):
        lenWords = len(words)
        # global macros
        if lenWords == 1:
            macro=self.kbMacros[words[0]]
            playString(macro[0],macro[1])
        # terminal application-specific (when ms OS running Dragon is not aware
        # of application context)
        elif lenWords > 1:
            macro, flags=self.kbMacros[' '.join(words)]
            # process application specific macro
            newmacro = macro
            newflags = flags
            if words[0] == 'screen':
                # screen command prefix is c-a, 0x04 modifier
                newflags = 0x00 # specify explicitly instead {ctrl}
                newmacro = ''.join(['{ctrl+a}',str(newmacro)])
            elif words[0] == 'vim':
                # vim command mode entered with 'esc 'key, command line
                # commands entered with : prefix and require 'enter' to
                # complete. No modifier on initial command character.
                newflags = 0x00
                if macro.startswith(':'):
                    newmacro=''.join([str(newmacro),'{enter}'])
                newmacro = ''.join(['{esc}',str(newmacro)])
            playString(newmacro,newflags)

    def gotResults_abrvPhrase(self, words, fullResults):
        phrase=self.abrvMap[words[0]]
        recognitionMimic(phrase.split())

    def gotResults_quickStart(self, words, fullResults):
        """Bottom QuickStart item can be approximated by 56 pixels above the bottom
        of the screen (because this depends on the fixed size text date which
        doesn't scale much with screen resolution). Bottom QuickStart item = y
        obtained from screen dimension-getScreenSize())-56. when selecting an
        icon to operate on the QuickStart corneriterate from bottom left (row
        1:1) """

        # screen dimensions (excluding taskbar)
        x, y = getScreenSize()
        # number of pixels between bottom of screen and bottom row of QuickStart icons
        row_initial = 56
        # number of pixels between left side of taskbar and first column of icons
        col_initial = 14
        # separation between subsequent QuickStart rows andcolumns
        col_sep = row_sep = 25
        # coordinate calculated using row and column numbers ofpress QuickStart icon
        x, y = x + col_initial, y - row_initial
        x, y = x + (col_sep * (int(words[5]) - 1)), y - (row_sep * (int(words[3]) - 1))
        # get the equivalent event code of the type of mouse event to perform
        # leftclick, rightclick, rightdouble-click
        event = self.kmap[str(str(words[1]) + 'click')]
        # play events down click and then release (for left double click
        # increment from left button up event which produces no action
        # then when incremented, performs the double-click)
        natlink.playEvents([(wm_mousemove, x, y), (event, x, y), (event + 1, x, y)])

    def callBack_popWin(self, hwnd, args):
        """ this callback function is called with handle of each top-level
        window. Window handles are used to check the of window in question is
        visible and if so it's title strings checked to see if it is a standard
        application (e.g. not the start button or natlink voice command itself).
        Populate dictionary of window title keys to window handle values. """
        if wg.IsWindowVisible(hwnd):
            winText = wg.GetWindowText(hwnd).strip()
            if winText and winText not in self.nullTitles and\
               winText not in args[1].values():
#                print [self.nullTitles + args[1].values()]
##               and args[0] != winText.split():
#               and filter(lambda x: x in args[0], winText.split()):
                # key on unique handle, not text of window
                args[1].update({hwnd: winText})
#            elif winText:
#                print("Skipping duplicate handle ({0}) for window \
#'{1}'".format(str(hwnd),winText))

    def gotResults_windowFocus(self, words, fullResults):
        """ Vertical taskbar window titles are spreadstarting from 150 pixels
        from the taskbar. Each subsequent icon is roughly 25 pixels
        between.this is assumed the same between subsequent rows.  If 'from
        bottom'modifier supplied to command string, calculate the offset from
        the first window title. This technique is an alternative to be able to
        determine a phrase that would be recognised as a window title (e.g.
        explicitly setting each programs window title)
        TODO: need to fix, inaccurate when counting from bottom. Windows not
        filtered properly? (Extra "Windows" which do not have physical
        component)"""

        # Detect the optional word to specify offset for variable component
        if words[1] == 'on':
            repeat_modifier_offset = 3
        else:
            repeat_modifier_offset = 2
        # screen dimensions (excluding taskbar)
        x, y = getScreenSize()
        # number of pixels between top of screen and top row of taskbar icons
        row_initial = 30 #75
        # number of pixels between left side of taskbar and first column of icons
        col_initial = 14
        # separation between subsequent rows
        row_sep = 25 #32
        # coordinate calculated, vertical offset is from top, horizontal offset
        # from the left edge of the taskbar (maximum horizontal value of the
        # screen visible size (which excludes taskbar)) calculated earlier.
        x, y = x + col_initial, row_initial  # initial position
        # argument to pass to callback contains words used in voice command
        # (this is also a recognised top-level window?) And dictionary of
        # strings: handles. (Window title: window handle)
        wins = (words, {})
        # selecting index from bottom window title on taskbar
        # enumerate all top-level windows and send handles to callback
        wg.EnumWindows(self.callBack_popWin,wins)
        # after visible taskbar application windows have been added to
        # dictionary (second element of wins tuple), we can calculate
        # relative offset from last taskbar title.
        total_windows = len(wins[1])
        # print('Number of taskbar applications: {0};'.format( total_windows))
        # print wins[1].keys()
        # enumerate child windows of visible desktop top-level windows.
        # we want to use the dictionary component of wins and create a map of
        # parent to child Window handles.
        win_map= {}
#        for hwin in wins[1].iterkeys():
        ch_wins= []
        hwin=wins[1].keys()[0]
        #hwin=wins[1][wins[1].keys()[0]]
        #print wg.GetWindowRect(hwin)
        #wg.EnumChildWindows(hwin,self.callBack_popChWin,ch_wins)
        win_map[hwin]=ch_wins
        if words[4:5]:
            # get desired index "from bottom" (negative index)
            from_bottom_modifier = int(words[ repeat_modifier_offset])
            # maximum number of increments is total_Windows -1
            relative_offset = total_windows - from_bottom_modifier - 1
        else:
            # get the index of window title required, add x vertical offsets
            # to get right vertical coordinate(0-based)
            relative_offset = int(words[repeat_modifier_offset])
        if 0 > relative_offset or relative_offset >= total_windows:
            print('Specified taskbar index out of range. '
                  '{0} window titles listed.'.format(total_windows))
            return 1
        y = y + (row_sep *  relative_offset)
        event = self.kmap['leftclick']
        # move mouse to 00 first, avoids occasional click failure
        natlink.playEvents([(wm_mousemove, 0,0),(wm_mousemove, x, y), (event, x, y), (event + 1, x, y)])

    def gotResults_repeatKey(self, words, fullResults):
        num = int(words[3])
        key = self.kmap[str(words[2])]
        # assumed only standard keys (not system keys) will want to be repeated
        #  this requires wm_key(down|up)
        [natlink.playEvents([(wm_keydown, key, 1), (wm_keyup, key, 1)]) for i
         in range(num)]

    def gotResults_androidSC(self, words, fullResults):
        print 'Screen dimensions: ' + str(getScreenSize())
        print 'Mouse cursor position: ' + str(getCursorPos())
        print 'Entire recognition result: ' + str(fullResults)
        print 'Partial recognition result: ' + str(words)

    def initialize(self):
        self.load(self.gramSpec)
        self.activateAll()

thisGrammar = ThisGrammar()
thisGrammar.initialize()


def unload():
    global thisGrammar
    if thisGrammar:
        thisGrammar.unload()
    thisGrammar = None
