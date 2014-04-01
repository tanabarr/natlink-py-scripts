#
# Python Macro Language for Dragon NaturallySpeaking
# 031213: significant amounts of functionality transferred to vocola/unimacro
#

import natlink
## 060114: import directly from main, otherwise not found (vocola 2.8.1 update)
#from natlinkmain import recognitionMimic
from natlinkutils import *
import win32gui as wg
import logging
import ioutils as iou
import traceback as tb
import time

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger()

# Windows GUI\ parameters
QS_ROW_INITIAL =56
# number of pixels between left side of taskbar and first column of icons
QS_COL_INITIAL =14
# separation between subsequent QuickStart rows andcolumns
QS_COL_ROW_SEP =25
# number of pixels between top of screen and top row of taskbar icons
TB_ROW_INITIAL =35 #75
# number of pixels between left side of taskbar and first column of icons
TB_COL_INITIAL =13
# separation between subsequent rows
TB_ROW_SEP =24 #32

class ThisGrammar(GrammarBase):


    # Macros can be repeated with recognitionmimic function which takes list of
    # words as parameter
    abrvMap = {'norm': 'switch to normal mode', 'spelling':
               'switch to spell mode',
               #'get out': 'press escape',
               #'insert': 'press insert', #'hash': 'press hash',
               #'sleep': 'go to sleep',
               #'window left': 'press windows left',
               #'window right': 'press windows right',
               #'page': 'page down',
                }

    # Todo: embed this list of strings within grammar to save space
    # mapping of keyboard keys to virtual key code to send as key input
    # VK_SPACE,VK_UP,VK_DOWN,VK_LEFT,VK_RIGHT,VK_RETURN,VK_BACK
    kmap = {'space': 0x20, 'up': 0x26, 'down': 0x28, 'left': 0x25, 'right': 0x27,
            'enter': 0x0d, 'backspace': 0x08, 'delete': 0x2e, 'leftclick': 0x201,
            'rightclick': 0x204, 'doubleclick': 0x202}

    # WinSCP seems to display extra window
    nullTitles = ['Default IME', 'MSCTFIME UI', 'Engine Window',
                  'VDct Notifier Window', 'Program Manager',
                  'Spelling Window', 'Start', 'WinSCP']

    # window handler
    windows = iou.Windows(nullTitles=nullTitles)

    # load default macros from file, tries db file first
    ### EDIT: 031213
    ###schema="name text, string text, flags text"
    ###fs = iou.FileStore(schema=schema) #preDict=kbMacros)
    ###kbMacros = fs.postDict

    ## writing to files/db and reading updates now performed in fs
    #initialisation
    #fs.writedb(schema,db_filename='tmp.db')

    gramSpec = """
        <quickStart> exported = QuickStart (left|right|double) row ({3}) column ({3});
        <bumpMouse> exported = ({0}) ({0}) bump mouse [({4})];
        <repeatKey> exported = repeat key ({2}) ({4}|{5});
        <windowFocus> exported = (focus|close) [on] window ({4}) [from bottom];
        <androidSC> exported =  show coordinates and screen size;
        <abrvPhrase> exported = ({1}) [mode];
        <reloadEverything> exported = reload everything;
        <resetMic> exported = reset microphone;
    """.format( 
                str(range(10)).strip('[]').replace(', ','|'),
                '|'.join(abrvMap.keys()),
                '|'.join(kmap.keys()),
                str(range(7)).strip('[]').replace(', ','|'),
                str(range(21)).strip('[]').replace(', ','|'),
                str(range(20,50,10)).strip('[]').replace(', ','|'))
        ### EDIT: 031213      removed
        ###' | '.join(kbMacros.keys()),
        ###<kbMacro> exported = ({0});
        ###<kbMacroPrint> exported = show macros [(vim|screen)];

    #msgPy = 'New PYD Test'
    msgPy = 'Messages from Python Macros'
    def gotResults_reloadEverything(self, words, fullResults):
        # bring window to front
        #[logger.info(k) for k in self.kbMacros.keys)]
        self.windows.winDiscovery(beginTitle=self.msgPy)
        natlink.playString('{alt}{down}',0)
        # seems to close unexpectedlywhen issuing the following
        #natlink.playString('{enter}',0)

    def gotResults_kbMacroPrint(self, words, fullResults):
        # todo: bring window to front
        for k,v in sorted(self.kbMacros.items()):
            try:
                if not k.startswith(words[2]):
                    continue
            except:
                pass
            logger.info("%s -> %s" % (k, self.kbMacros[k].string))
            #[logger.info(k) for k in self.kbMacros.keys().sort()]
        self.windows.winDiscovery(beginTitle=self.msgPy)

    def gotResults_kbMacro(self, words, fullResults):
        lenWords = len(words)
        # global macros
        if lenWords == 1:
            macro=self.kbMacros[words[0]]
            natlink.playString(macro.string,macro.flags)
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
            #logger.debug('vim resultant macro: %s'% newmacro)
            natlink.playString(newmacro,newflags)

    def gotResults_abrvPhrase(self, words, fullResults):
        phrase=self.abrvMap[' '.join(words)]
        natlink.recognitionMimic(phrase.split())

#    def sanitise_movement(func):
#        def checker(*args,**kwargs):
#            print args
#            print kwargs
#            ret = func(*args,**kwargs)
#            return ret
#        return checker

    def gotResults_bumpMouse(self, words, fullResults):
        """ usage: <x-axis increment> <y-axis increment> bump <integer
        multiplier> """

        # register required displacement integer 's\s
        x_disp = int(words[0])
        y_disp = int(words[1])
        # do we have a multiplier?
        mult = 1
        if len(words) > 4:
            mult = int(words[4])
        logger.error( "mult:%d" % (mult))
        # get current position
        x, y = natlink.getCursorPos()
        logger.error("x:%d y:%d" % (x,y))
        # coordinate calculated using row and column pixel offset * multiplier 
        if int(words[0]) != 0:
            x = x + x_disp*mult
        if int(words[1]) != 0:
            y = y + y_disp*mult

        logger.error("x:%d y:%d" % (x,y))
        # get the equivalent event code of the type of mouse event to perform
        # leftclick, rightclick, rightdouble-click (currently left click)
        event = self.kmap['leftclick']
        # play events down click and then release (for left double click
        # increment from left button up event which produces no action
        # then when incremented, performs the double-click)
        natlink.playEvents([(wm_mousemove, x, y), (event, x, y), (event + 1, x, y)])


    def gotResults_quickStart(self, words, fullResults):
        """Bottom QuickStart item can be approximated by 56 pixels above the bottom
        of the screen (because this depends on the fixed size text date which
        doesn't scale much with screen resolution). Bottom QuickStart item = y
        obtained from screen dimension-getScreenSize())-56. when selecting an
        icon to operate on the QuickStart corner iterate from bottom left (row
        1:1) """

        self.windows.nullTitles.append(' '.join(words))

        # screen dimensions (excluding taskbar)
        x, y = natlink.getScreenSize()
        # number of pixels between bottom of screen and bottom row of QuickStart icons
        row_initial = QS_ROW_INITIAL #56g
        # number of pixels between left side of taskbar and first column of icons
        col_initial = QS_COL_INITIAL #14
        # separation between subsequent QuickStart rows andcolumns
        col_sep = row_sep = QS_COL_ROW_SEP #25
        # coordinate calculated using row and column numbers ofpress QuickStart icon
        x, y = x + col_initial, y - row_initial
        x, y = x + (col_sep * (int(words[5]) - 1)), y - (row_sep * (int(words[3]) - 1))
        # get the equivalent event code of the type of mouse event to perform
        # leftclick, rightclick, rightdouble-click
        event = self.kmap[str(str(words[1]) + 'click')]
        # play events down click and then release (for left double click
        # increment from left button up event which produces no action
        # then when incremented, performs the double-click)

        playEvents = iou.sanitise_movement(natlink.playEvents)
#    @sanitise_movement
        playEvents([(wm_mousemove, x, y), (event, x, y), (event + 1, x, y)])


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
        component). EDIT: 291013 - modified to perform close"""

        # Detect the optional word to specify offset for variable component
        if words[1] == 'on':
            repeat_modifier_offset = 3
        else:
            repeat_modifier_offset = 2
        # screen dimensions (excluding taskbar)
        x, y = natlink.getScreenSize()
        # number of pixels between top of screen and top row of taskbar icons
        row_initial = TB_ROW_INITIAL #35 #75
        # number of pixels between left side of taskbar and first column of icons
        col_initial = TB_COL_INITIAL #14
        # separation between subsequent rows
        row_sep = TB_ROW_SEP #23 #32
        # coordinate calculated, vertical offset is from top, horizontal offset
        # from the left edge of the taskbar (maximum horizontal value of the
        # screen visible size (which excludes taskbar)) calculated earlier.
        x, y = x + col_initial, row_initial  # initial position
        # argument to pass to callback contains words used in voice command
        # (this is also a recognised top-level window?) And dictionary of
        # strings: handles. (Window title: window handle)
        # selecting index from bottom window title on taskbar
        # enumerate all top-level windows and send handles to callback
        wins={}
        try:
            # Windows dictionary returned assecond element of tuple
            wins=self.windows.winDiscovery(skipTitle=' '.join(words))[1]
            #logger.debug('enumerate Windows: %s'% wins)
        except:
            logger.error('cannot enumerate Windows')
            logger.error(tb.format_exc())
            return

        # after visible taskbar application windows have been added to
        # dictionary (second element of wins tuple), we can calculate
        # relative offset from last taskbar title.
        total_windows = len(wins)
        # print('Number of taskbar applications: {0};'.format( total_windows))
        # print wins.keys()
        # enumerate child windows of visible desktop top-level windows.
        # we want to use the dictionary component of wins and create a map of
        # parent to child Window handles.
        if words[4:5]:
            # get desired index "from bottom" (negative index)
            from_bottom_modifier = int(words[ repeat_modifier_offset])
            # maximum number of increments is total_Windows -1
            relative_offset = total_windows - from_bottom_modifier #- 1
        else:
            # get the index of window title required, add x vertical offsets
            # to get right vertical coordinate(0-based)
            relative_offset = int(words[repeat_modifier_offset]) - 1
        if 0 > relative_offset or relative_offset >= total_windows:
            print('Specified taskbar index out of range. '
                  '{0} window titles listed.'.format(total_windows))
            return 1
        y = y + (row_sep *  relative_offset)
        # move mouse to 00 first then separate mouse movement from click events
        # this seems to avoid occasional click failure
        natlink.playEvents([(wm_mousemove, x, y),])
        if words[0] == 'close':
            event = self.kmap['rightclick']
            natlink.playEvents([(event, x, y), (event + 1, x, y)])
            natlink.playString('{up}{enter}')
        else:
            event = self.kmap['leftclick']
            natlink.playEvents([(event, x, y), (event + 1, x, y)])

    def gotResults_repeatKey(self, words, fullResults):
        num = int(words[3])
        key = self.kmap[str(words[2])]
        # assumed only standard keys (not system keys) will want to be repeated
        #  this requires wm_key(down|up)
        [natlink.playEvents([(wm_keydown, key, 1), (wm_keyup, key, 1)]) for i
         in range(num)]

    def gotResults_androidSC(self, words, fullResults):
        print 'Screen dimensions: ' + str(natlink.getScreenSize())
        print 'Mouse cursor position: ' + str(natlink.getCursorPos())
        print 'Entire recognition result: ' + str(fullResults)
        print 'Partial recognition result: ' + str(words)

    def gotResults_resetMic(self, words, fullResults):
        natlink.setMicState("off")
        time.sleep(2)
        natlink.setMicState("on")

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

# requests
##copy putty output to Windows clipboard: 'click system''copy all to clipboard'
##screen copy mode/paste macros
