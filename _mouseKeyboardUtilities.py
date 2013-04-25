#
# Python Macro Language for Dragon NaturallySpeaking
# nextApp
#

import natlink
from natlinkutils import *
import win32gui as wg

class ThisGrammar(GrammarBase):

    gramSpec = """
        <start> exported = QuickStart (left|right|double) row (1|2|3|4|5)
            column (1|2|3|4|5);
        <repeatKey> exported = repeat key (space|up|down|left|
            right|enter|backspace|delete)  [ (1|2|3|4|5|6|7|8|9|10|
            11|12|13|14|15|16|17|18|19|20|30|40|50|100) ];
        <windowFocus> exported = focus [on] window (0|1|2|3|4|5|6|7|8|9|10|
            11|12|13|14|15|16|17|18|19) [from bottom];
        <androidSC> exported =  show coordinates and screen size;
        <abrvPhrase> exported = (normal|spell|insert|escape) [mode];
        <kbMacro> exported = private|next|previous;
    """
##        <androidSC> exported =press [the] button (home|menu|back|search|call|endcall);

    # Todo: embed this list of strings within grammar to save space
    # mapping of keyboard keys to virtual key code to send as key input
    # VK_SPACE,VK_UP,VK_DOWN,VK_LEFT,VK_RIGHT,VK_RETURN,VK_BACK
    kmap = {'space': 0x20, 'up': 0x26, 'down': 0x28, 'left': 0x25, 'right': 0x27,
            'enter': 0x0d, 'backspace': 0x08, 'delete': 0x2e, 'leftclick': 0x201,
            'rightclick': 0x204, 'doubleclick': 0x202}

    # Todo: embed this list of strings within grammar to save space
    # list of android screencast buttons
    buttons = ['home', 'menu', 'back', 'search', 'call', 'endcall']

    nullTitles = ['Default IME', 'MSCTFIME UI', 'Engine Window',
                  'VDct Notifier Window', 'Program Manager',
                  'Spelling Window', 'Start']

    abrvMap = {'normal': 'switch to normal mode', 'spell': \
               'switch to spell mode', 'escape': 'press escape',
               'insert': 'press insert','hash': 'press hash'}

    def gotResults_kbMacro(self, words, fullResults):
        if words[0] == 'private':
            playString( 'N',0x05)
        else:
            playString({Tab},0x06)
        #playString( 'a',0x04)

    def gotResults_abrvPhrase(self, words, fullResults):
        phrase=self.abrvMap[words[0]]
#        print phrase
        recognitionMimic(phrase.split())

#        phrase=['switch','to',words[0],'mode']
#        print phrase

    def gotResults_start(self, words, fullResults):
    #    print 'Screen dimensions: ' + str(getScreenSize())
    #    print 'Mouse cursor position: ' + str(getCursorPos())
    #    print 'Entire recognition result: ' + str(fullResults)
    #    print 'Partial recognition result: ' + str(words)

        """ icons are spreadstarting from 14 pixels from the taskbar. Each subsequent
        icon is roughly 25 pixels between.this is assumed the same between subsequent rows.
        Bottom QuickStart item can be approximated by 56 pixels above the bottom
        of the screen (because this depends on the fixed size text date which
        doesn't scale much with screen resolution). Bottom QuickStart item = y
        (obtained from screen dimension-getScreenSize())-56."""

        """ when selecting an icon to operate on the QuickStart corneriterate from
        bottom left (row 1:1) """

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
     #   print event

        # play events down click and then release (for left double click
        # increment from left button up event which produces no action
        # then when incremented, performs the double-click)
        natlink.playEvents([(wm_mousemove, x, y), (event, x, y), (event + 1, x, y)])

        #execute a control-left drag down 30 pixels
        #x,y = natlink.getCursorPos()
#:w         natlink.playString('{Tab}',0x02)

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

    def callBack_popChWin(self, chwnd, args):
        """ this callback function is called with handle of each child
        window. Window handles are used to check the of window in question is
        visible and if so it's title strings checked to see if it is a standard
        application (e.g. not the start button or natlink voice command itself).
        Populate dictionary of window title keys to window handle values. """
#        if wg.IsWindowVisible(hwnd):
#            winText = wg.GetWindowText(hwnd).strip()
#            if winText and winText not in self.nullTitles\
#               and args[0] != winText.split():
##               and filter(lambda x: x in args[0], winText.split())
        #winText='s'
        #args.update({hwnd:''})
#        print "... cw: {0} list: {1}".format(chwnd,args)
#        print "rect: {0}".format(wg.GetWindowRect(chwnd))
#        args.append(hwnd)
         #            args.appendh_wins[win]=children

    def gotResults_windowFocus(self, words, fullResults):
    #    print 'Screen dimensions: ' + str(getScreenSize())
    #    print 'Mouse cursor position: ' + str(getCursorPos())
    #    print 'Entire recognition result: ' + str(fullResults)
    #    print 'Partial recognition result: ' + str(words)

        """ Vertical taskbar window titles are spreadstarting from 150 pixels
        from the taskbar. Each subsequent icon is roughly 25 pixels
        between.this is assumed the same between subsequent rows.  If 'from
        bottom'modifier supplied to command string, calculate the offset from
        the first window title. This technique is an alternative to be able to
        determine a phrase that would be recognised as a window title (e.g.
        explicitly setting each programs window title)"""

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

        #print wins
        # enumerate child windows of visible desktop top-level windows.
        # we want to use the dictionary component of wins and create a map of
        # parent to child Window handles.
        win_map= {}
#        for hwin in wins[1].iterkeys():
        ch_wins= []
        hwin=wins[1].keys()[0]
        #hwin=wins[1][wins[1].keys()[0]]
        #print str(hwin)
        #print wg.GetWindowRect(hwin)
    #print
    #print str(hwin)

        #wg.EnumChildWindows(hwin,self.callBack_popChWin,ch_wins)
        win_map[hwin]=ch_wins

        #print win_map

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

    def gotResults_repeatKey(self, words, fullResults):
    #    print 'Entire recognition result: ' + str(fullResults)
    #    print 'Partial recognition result: ' + str(words)
    #    print 'Repeating "%s" key %d times!' % (str(words[2]), int(words[3]))
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

##        """  buttons along the bottom of the android screencast window.
##        to get the coordinates of the buttons, get the bottom left,
##        increment offsets up and increments of 25 right for each button.
##        doesn't scale much with screen resolution). Bottom QuickStart item = y
##        (obtained from screen dimension-getScreenSize())-56."""
##        """hack to get the bottom left corner is to mimic mousegrid seven
##        seven, then get cursor position"""
##
##        # assuming the correct window is in focus
##        recognitionMimic(["MouseGrid", "window"] + ["seven" for i in range(4)])
##        # escape from the MouseGrid, sets cursor position
##        recognitionMimic(["right", "one"])
##        #recognitionMimic(["mouse", "click"])
##        x, y = getCursorPos()
##        print 'Mouse cursor position: ' + str(getCursorPos())
##
##        # Todo: don't like magic numbers, find way to avoid hard coding.
##        # Note: doesn't work in maximised mode, window edges required
##        # number of pixels between bottom of android screencast window and
##        # control buttons
##        row_initial = 15
##        # number of pixels between left side of application window and centre
##        # of the first button
##        col_initial = 45
##        # separation in pixels between adjacent buttons (average)
##        col_sep = 30
##
##        # coordinate calculated using initial position of first button
##        x, y = x + col_initial, y - row_initial
##
##        # words array contains a keyword somewhere, we need the index of this
##        # keyword in our buttons array to work out the horizontal offset
###        ret = (self.buttons and words)
###        print(ret)
###        if ret is not None:
###            ret = self.buttons.index(ret)
####        ret = None
####        for word in words:
####            try:
####                ret =  self.buttons.index(word)
####                if ret: break
####            except: pass
###        [ if ret = buttons.index(i) is not None:  break for i in words ]
###        ret = filter(lambda x: x in self.buttons, words).next()
###       ret =  self.buttons.index(set(self.buttons).intersect(words))
##        ret = self.buttons.index(filter(lambda x: x in self.buttons, words)[0])
##        print(ret)
##        # Only needs increment horizontal value by the index of the button
##        x += col_sep * ret
##
##        # get the equivalent event code of the type of mouse event to perform
##        # leftclick, rightclick, rightdouble-click
##        event = self.kmap['leftclick']
##
##        # play events down click and then release (for left double click
##        # increment from left button up event which produces no action
##        # then when incremented, performs the double-click)
##        natlink.playEvents([(wm_mousemove, x, y), (event, x, y), (event + 1, x, y)])

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
