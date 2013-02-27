#
# Python Macro Language for Dragon NaturallySpeaking
# nextApp
#

import natlink
from natlinkutils import *
import win32gui as wg

class AppWindow():
    def __init__(self, name, rect):
        self.winName = name
        self.winRect = rect
        self.buttons = ['home', 'menu', 'back', 'search', 'call', 'endcall']
        self.mimicCmds = {}.fromkeys(self.buttons)

class ThisGrammar(GrammarBase, AppWindow):
    """ Class uses application window objects to store grid reference of window
    buttons on Windows without "say what you see" Dragon NaturallySpeaking
    functionality. To be developed to use the either grid reference or
    MouseGrid coordinate utterances. There should be functionality to add
    buttons in real-time, need to be backed up in persistent database."""

    appDict = {}
    appDict.update({"iphoneWin": AppWindow("tans-iPhone", None)})
    appDict.update({"xbmcChromeWin": AppWindow("XBMC - Google Chrome", None)})
    appSelectionStr = '(' + str(appDict.keys()).strip('][').replace(',','|') +\
    ')'

    appButtonStr = '(' + str(appDict["iphoneWin"].buttons).strip('][').replace(',','|') +\
    ')'

    print(appSelectionStr,appButtonStr)

    gramSpec = """
        <winclick> exported = click {0} {1};
        <iphonetap> exported = tap iphone (home|menu|back|search|call|endcall);
    """.format(appSelectionStr,appButtonStr)

    # Todo: embed this list of strings within grammar to save space
    # list of android screencast buttons
    iphoneCmdDict = appDict["iphoneWin"].mimicCmds
    print appDict["iphoneWin"].mimicCmds
    iphoneCmdDict.update({'home': ['1','5','8']})
    print iphoneCmdDict

    nullTitles = ['Default IME', 'MSCTFIME UI', 'Engine Window',
                  'VDct Notifier Window', 'Program Manager',
                  'Spelling Window', 'Start']

    def callBack_popWin(self, hwin, args):
        """ this callback function is called with handle of each top-level
        window. Window handles are used to check the of window in question is
        visible and if so it's title strings checked to see if it is a standard
        application (e.g. not the start button or natlink voice command itself).
        Populate dictionary of window title keys to window handle values. """
        if wg.IsWindowVisible(hwin):
            winText = wg.GetWindowText(hwin).strip()
            if winText and winText not in self.nullTitles and\
               winText not in args[1].values():
                args[1].update({hwin: winText})

    def gotResults_iphonetap(self, words, fullResults):
    #    print 'Screen dimensions: ' + str(getScreenSize())
    #    print 'Mouse cursor position: ' + str(getCursorPos())
    #    print 'Entire recognition result: ' + str(fullResults)
    #    print 'Partial recognition result: ' + str(words)

        self.winDiscovery(words, 'tans-iPhone')
#        # Detect the optional word to specify offset for variable component
#        if words[1] == 'on':
#            repeat_modifier_offset = 3
#        else:
#            repeat_modifier_offset = 2
#
#        # screen dimensions (excluding taskbar)
#        x, y = getScreenSize()
#
#        # number of pixels between bottom of screen and bottom row of QuickStart icons
#        row_initial = 75
#
#        # number of pixels between left side of taskbar and first column of icons
#        col_initial = 14
#
#        # separation between subsequent rows
#        row_sep = 32
#
#        # coordinate calculated, vertical offset is from top
#        x, y = x + col_initial, row_initial  # initial position
#
    def winDiscovery(self, words, wintitle):
        # argument to pass to callback contains words used in voice command
        # (this is also a recognised top-level window?) And dictionary of
        # strings: handles. (Window title: window handle)
        wins = (words, {})
        hwin = None
        # selecting index from bottom window title on taskbar
        # enumerate all top-level windows and send handles to callback
        wg.EnumWindows(self.callBack_popWin,wins)

        # after visible taskbar application windows have been added to
        # dictionary (second element of wins tuple), we can calculate
        # relative offset from last taskbar title.
        total_windows = len(wins[1])
        print('Number of taskbar applications: {0};'.format( total_windows))
        print wins[1]

        #print wins
        try:
            index = wins[1].values().index(wintitle)
        except:
            index = None

        if index is not None:
            print(index)
            hwin = (wins[1].keys())[index]
            print("Name: {0}, Handle: {1}".format(wins[1][hwin], str(hwin)))
            #wg.EnableWindow(hwin,True)
            # wg.SetFocus(hwin)
            wg.SetForegroundWindow(hwin)
            print wg.GetWindowRect(hwin)
            return wg.GetWindowRect(hwin)
            print mimicCmds
        print str(hwin)
        return str(hwin)


#        if words[4:5]:
#            # get desired index "from bottom" (negative index)
#            from_bottom_modifier = int(words[ repeat_modifier_offset])
#            # maximum number of increments is total_Windows -1
#            relative_offset = total_windows - from_bottom_modifier - 1
#        else:
#            # get the index of window title required, add x vertical offsets
#            # to get right vertical coordinate(0-based)
#            relative_offset = int(words[repeat_modifier_offset])
#
#        if 0 > relative_offset or relative_offset >= total_windows:
#            print('Specified taskbar index out of range. '
#                  '{0} window titles listed.'.format(total_windows))
#            return 1
#        y = y + (row_sep *  relative_offset)
#
#        event = self.kmap['leftclick']
#        # move mouse to 00 first, avoids occasional click failure
#        natlink.playEvents([(wm_mousemove, 0,0),(wm_mousemove, x, y), (event, x, y), (event + 1, x, y)])
    '''
        natlink.playEvents([ (wm_syskeydown,0x12,1),
                              (wm_keydown,0x09,1),
                              (wm_keyup,0x09,1),#(wm_lbuttondown,x,y),
                              (wm_keydown,0x09,1),
                              (wm_keyup,0x09,1),#(wm_lbuttondown,x,y),
                              #(wm_mousemove,x,y+30),
                              #(wm_lbuttonup,x,y+30),
                               (wm_syskeyup,0x12,1)
                            ])
    '''


#    def gotResults_androidSC(self, words, fullResults):
#        print 'Screen dimensions: ' + str(getScreenSize())
#        print 'Mouse cursor position: ' + str(getCursorPos())
#        print 'Entire recognition result: ' + str(fullResults)
#        print 'Partial recognition result: ' + str(words)
#
#        """  buttons along the bottom of the android screencast window.
#        to get the coordinates of the buttons, get the bottom left,
#        increment offsets up and increments of 25 right for each button.
#        doesn't scale much with screen resolution). Bottom QuickStart item = y
#        (obtained from screen dimension-getScreenSize())-56."""
#        """hack to get the bottom left corner is to mimic mousegrid seven
#        seven, then get cursor position"""
#
#        # assuming the correct window is in focus
#        recognitionMimic(["MouseGrid", "window"] + ["seven" for i in range(4)])
#        # escape from the MouseGrid, sets cursor position
#        recognitionMimic(["right", "one"])
#        #recognitionMimic(["mouse", "click"])
#        x, y = getCursorPos()
#        print 'Mouse cursor position: ' + str(getCursorPos())
#
#        # Todo: don't like magic numbers, find way to avoid hard coding.
#        # Note: doesn't work in maximised mode, window edges required
#        # number of pixels between bottom of android screencast window and
#        # control buttons
#        row_initial = 15
#        # number of pixels between left side of application window and centre
#        # of the first button
#        col_initial = 45
#        # separation in pixels between adjacent buttons (average)
#        col_sep = 30
#
#        # coordinate calculated using initial position of first button
#        x, y = x + col_initial, y - row_initial
#
#        # words array contains a keyword somewhere, we need the index of this
#        # keyword in our buttons array to work out the horizontal offset
#        ret = self.buttons.index(filter(lambda x: x in self.buttons, words)[0])
#        print(ret)
#        # Only needs increment horizontal value by the index of the button
#        x += col_sep * ret
#        self.click('leftclick',x,y)
#
    def click(self, clickType, x, y):
        # get the equivalent event code of the type of mouse event to perform
        # leftclick, rightclick, rightdouble-click (see kmap)
        event = self.kmap[clickType]

        # play events down click and then release (for left double click
        # increment from left button up event which produces no action
        # then when incremented, performs the double-click)
        natlink.playEvents([(wm_mousemove, x, y), (event, x, y), (event + 1, x, y)])

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
