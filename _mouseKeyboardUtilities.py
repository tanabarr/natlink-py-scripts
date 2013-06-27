#
# Python Macro Language for Dragon NaturallySpeaking
# nextApp
#

import natlink
from natlinkutils import *
import win32gui as wg
import logging as log

log.basicConfig(level=log.DEBUG)

class MacroObj():
    def __init__(self,string='',flags=0):
        self.string=string
        self.flags=flags

## unfinished
#class FileStore():
#    def __init__(self,filename='defaultkbm.txt',preDict={},delim="'"):
#        self.f=open(filename,'a')
#        self.postDict=preDict
#        self.readfile()
#        self.writefile()
#        return postDict
#
#    def readfile(self):
#        for line in self.f.readlines():
#            if (len(line) == 3):
#                gram, macro, flags = line.split("'") #delim)
#                self.postDict[gram] = MacroObj(macro, flags)
#
#    def writefile(self):
#        for gram, macroobj in self.postDict.iteritems():
#            try:
#                self.f.write("'".join([gram, macroobj.string, macroobj.flags]))
#            except:
#                #parse object instead of just string values
#                pass
#

class ThisGrammar(GrammarBase):

    # Macros can be repeated with recognitionmimic function which takes list of
    # words as parameter
    abrvMap = {'norm': 'switch to normal mode', 'spell':
               'switch to spell mode', 'escape': 'press escape',
               'insert': 'press insert','hash': 'press hash',
               'sleep': 'go to sleep',
                }

    # playstring function of nat link uses format:
    # playstring(<keystring>, modifier flags: <ctrl,alt,shift>(bitwise 3 LSBs)
    # modifier applies to first character in string. More information in 'natlink.txt'
    kbMacros = {# Global commands
                'downshift': ('{down}',0x01), 'rightshift': ('{right}',0x01),
                'leftshift': ('{left}',0x01), 'upshift': ('{up}',0x01),
                #'select': ('{ctrl+shift}',0x00),#c-s-click requires playevents
                # Generic application commands
                'back tab': ('{shift+tab}',0x00),
                'save': ('s',0x04),
                'new': ('n',0x04),
                'zoom in': ('+',0x04),
                'zoom out': ('-',0x04),
                # Google Chrome commands
                'next': ('{ctrl+tab}',0x00),
                'previous': ('{ctrl+shift+tab}',0x01),
                'private': ('N',0x05),
                'close': ('w',0x04), 'flag': ('{alt}aa',0x00),
                'bookmark': ('b',0x04),
                'tools': ('e',0x02),
                # Shell related commands
                'close prompt': ('{space}c',0x02),# 'prompt 'closes command prompt
                'bash history': ('r',0x04),
                # c-style programming abbreviations
                'vim begin comment': ('i/* ',0),
                'vim end comment': ('i */{enter}',0),
                'vim begin long comment': ('i#{esc}ib{space}',0),
                'vim end long comment': ('i#{esc}ie{enter}',0),
                'vim line comment': ('i#{esc}il{enter}',0),
                'vim define': ('i#{esc}id{space}',0),
                'vim include': ('i#{esc}ii{space}',0),
                'vim equals': ('i = ',0),
                # vim commands
                'vim format': ('Q',0x00), 'vim undo': ('u',0x00),
                'vim redo': ('{ctrl+r}',0x00), 'vim next': (':bn',0x00),
                'vim remove buffer': (':bd',0x00),
                'vim start macro': ('qz',0),
                'vim repeat macro': ('@z',0),
                'vim previous': (':bp',0x00), 'vim save': (':w',0x00),
                'vim close': (':q',0x00), 'vim taglist': ('{ctrl+p}',0x00),
                'vim update': (':!ctags -a .',0x01),
                'vim last change': ('g,',0x00),
                'vim return': ("''",0x00),
                'vim matching': ('%',0x00),
                'vim undo jump': ('``',0x00),
                'vim insert space': ('i{space}{esc}',0),
                'vim hash': ('i#{esc}',0),
                'vim insert blank line next': ('o{up}',0),
                'vim insert blank line previous': ('O{down}',0),
                'vim last command': (':{up}',0xff),
                'vim copy previous line': (':-1y',0),
                'vim copy next line': (':+1y',0),
                'vim remove previous line': (':-1d',0),
                'vim remove next line': (':+1d',0),
                'vim set mark': ('mz',0),
                'vim goto mark': ("'zi",0),
                'vim scroll to top': ('zt',0),
                'vim scroll to bottom': ('zb',0),
                'vim edit another': (':edit ',0xff),
                'vim folds': ('{ctrl+f}',0x00),
                'vim window up': ('{ctrl+k}',0x00),
                'vim window down': ('{ctrl+j}',0x00),
                'vim window left': ('{ctrl+h}',0x00),
                'vim window right': ('{ctrl+l}',0x00),
                'vim split vertical': (':vsp',0x00),
                #screen commands
                'attach screen ': ('screen -R{enter}',0x20000),
                'attach screen existing': ('screen -x{enter}',0x20000),
                'screen scrollback mode': ('[',0x00),
                'screen scrollback paste': (']',0x00),
                'screen previous': ('p',0x00), 'screen next': ('n',0x00),
                'screen help': ('?',0x00), 'screen new': ('c',0x00),
                'screen detach': ('d',0x00), 'screen list': ('"',0x00),
                'screen kill': ('k',0x00), 'screen title': ('A',0x00),
                # window split related
                'screen switch': ('{tab}',0x00), 'screen split': ('S',0x00),
                'screen vertical': ('|',0x00), 'screen crop': ('Q',0x00),
                'screen remove': ('X',0x00),
                # Windows live mail shortcuts
                'live moved to folder': ('{ctrl+shift+v}',0),
                'live sort by date': ('{alt}vb{down}{enter}',0),
                'live sort by flag': ('{alt}vb' + 6*'{down}' + '{enter}',0),
                # todo: fix
                'live emails': ('{esc}{tab}',0x100),
                }

    #kbMacros = {k: MacroObj(v[0],v[1]) for k, v in self.kbMacros.iteritems()}
    #kbMacros = dict([(k, MacroObj(v[0],v[1])) for k, v in
    #    kbMacros.iteritems()])
    # we want to be able to reference the macro string as an attribute
    # dictionary comprehension not available pre python 2.7
    for k, v in kbMacros.iteritems():
        kbMacros[k] = MacroObj(v[0],v[1])

    ##result=FileStore()
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

    gramSpec = """
        <quickStart> exported = QuickStart (left|right|double) row ({3}) column ({3});
        <repeatKey> exported = repeat key ({2}) ({4}|{5});
        <windowFocus> exported = focus [on] window ({4}) [from bottom];
        <androidSC> exported =  show coordinates and screen size;
        <abrvPhrase> exported = ({1}) [mode];
        <kbMacro> exported = ({0});
        <kbMacroPrint> exported = show macros;
        <reloadEverything> exported = reload everything;
    """.format(' | '.join(kbMacros.keys()),'|'.join(abrvMap.keys()),
               '|'.join(kmap.keys()),
                str(range(6)).strip('[]').replace(', ','|'),
                str(range(20)).strip('[]').replace(', ','|'),
                str(range(20,50,10)).strip('[]').replace(', ','|'))

    def gotResults_reloadEverything(self, words, fullResults):
        # todo: bring window to front
        #[log.info(k) for k in self.kbMacros.keys()]
        pass

    def gotResults_kbMacroPrint(self, words, fullResults):
        # todo: bring window to front
        [log.info(k) for k in self.kbMacros.keys()]

    def gotResults_kbMacro(self, words, fullResults):
        lenWords = len(words)
        # global macros
        if lenWords == 1:
            macro=self.kbMacros[words[0]]
            playString(macro.string,macro.flags)
        # terminal application-specific (when ms OS running Dragon is not aware
        # of application context)
        elif lenWords > 1:
            macro=self.kbMacros[' '.join(words)]
            # process application specific macro
            newmacro = macro.string
            newflags = macro.flags
            if words[0] == 'screen':
                # screen command prefix is c-a, 0x04 modifier
                newflags = 0x04
                newmacro = ''.join(['a',str(newmacro)])
            elif words[0] == 'vim':
                # vim command mode entered with 'esc 'key, command line
                # commands entered with : prefix and require 'enter' to
                # complete. No modifier on initial command character.
                # 0xff signifies partial command, don't append "enter"
                newflags = 0x00
                if macro.string.startswith(':') and macro.flags != 0xff:
                    newmacro=''.join([str(newmacro),'{enter}'])
                newmacro = ''.join(['{esc}',str(newmacro)])
            log.debug('vim resultant macro: %s'% newmacro)
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

        self.nullTitles.append(' '.join(words))

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
            try:
                winText = wg.GetWindowText(hwnd).strip()
                if winText and winText not in self.nullTitles and\
                 winText not in args[1].values():
                    args[1].update({hwnd: winText})
            except:
                log.error('cannot retrieve window title')
#               and filter(lambda x: x in args[0], winText.split()):

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
        try:
            wg.EnumWindows(self.callBack_popWin,wins)
            log.debug('enumerate Windows: %s'% wins[1])
        except:
            log.error('cannot enumerate Windows')

        # after visible taskbar application windows have been added to
        # dictionary (second element of wins tuple), we can calculate
        # relative offset from last taskbar title.
        total_windows = len(wins[1])
        # print('Number of taskbar applications: {0};'.format( total_windows))
        # print wins[1].keys()
        # enumerate child windows of visible desktop top-level windows.
        # we want to use the dictionary component of wins and create a map of
        # parent to child Window handles.
        #/win_map= {}
#       #/ for hwin in wins[1].iterkeys():
        #/ch_wins= []
        #/hwin=wins[1].keys()[0]
        #/#hwin=wins[1][wins[1].keys()[0]]
        #/#print wg.GetWindowRect(hwin)
        #/#wg.EnumChildWindows(hwin,self.callBack_popChWin,ch_wins)
        #/win_map[hwin]=ch_wins
        if words[4:5]:
            # get desired index "from bottom" (negative index)
            from_bottom_modifier = int(words[ repeat_modifier_offset])
            # maximum number of increments is total_Windows -1
            relative_offset = total_windows - from_bottom_modifier - 1
        else:
            # get the index of window title required, add x vertical offsets
            # to get right vertical coordinate(0-based)
            relative_offset = int(words[repeat_modifier_offset]) - 1
        if 0 > relative_offset or relative_offset >= total_windows:
            print('Specified taskbar index out of range. '
                  '{0} window titles listed.'.format(total_windows))
            return 1
        y = y + (row_sep *  relative_offset)
        event = self.kmap['leftclick']
        # move mouse to 00 first then separate mouse movement from click events
        # this seems to avoid occasional click failure
        natlink.playEvents([(wm_mousemove, x, y),])
        #natlink.playEvents([(wm_mousemove, 0,0),(wm_mousemove, x, y)])
        natlink.playEvents([(event, x, y), (event + 1, x, y)])

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

    #def winDiscovery(self, words, appName=None):
    #    # argument to pass to callback contains words used in voice command
    #    # (this is also a recognised top-level window?) And title of window
    #    # to find (optional). Return tuple (handle of appName, window dict).
    #    wins = (words, {})
    #    hwin = None
    #    # selecting index from bottom window title on taskbar
    #    # enumerate all top-level windows and send handles to callback
    #    wg.EnumWindows(self.callBack_popWin, wins)
    #    # after visible taskbar application windows have been added to
    #    # dictionary (second element of wins tuple), we can calculate
    #    total_windows = len(wins[1])
    #    # print('Number of taskbar applications: {0};'.format( total_windows))
    #    # the function called without selecting window to find, just return
    #    # window dictionary.
    #    if not appName:
    #        return (None, wins[1])
    #    #log.debug("discover window %s" % appName)
    #    # trying to find window title of selected application within window
    #    # dictionary( local application context). Checking that the window
    #    # exists and it has a supportive local application context.
    #    try:
    #        app = self.appDict[str(appName)]
    #        #log.debug("supported application window: %s" % app.winName)
    #        index = wins[1].values().index(app.winName)
    #        # need to find a list of Windows again (to refresh)
    #    except:
    #        index = None

    #    if index is not None:
    #        #log.debug("index of application window: %d" % index)
    #        hwin = (wins[1].keys())[index]
    #        log.debug(
    #            "Name: {0}, Handle: {1}".format(wins[1][hwin], str(hwin)))
    #        app.winHandle = hwin
    #        wg.SetForegroundWindow(hwin)
    #        # print wg.GetWindowRect(hwin)
    #        app.winRect = wg.GetWindowRect(hwin)
    #        # print str(hwin)
    #        return (str(hwin), wins[1])

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
