#
# Python Macro Language for Dragon NaturallySpeaking
# nextApp
#

import natlink
from natlinkutils import *
import win32gui as wg
import os
from subprocess import Popen
import macroutils as mu
import logging
import time
import wmi

logging.basicConfig(level=logging.INFO)

class ThisGrammar(GrammarBase):

    """ Class uses application window objects to store grid reference of window
    buttons on Windows without "say what you see" Dragon NaturallySpeaking
    functionality. To be developed to use the either grid reference or
    MouseGrid coordinate utterances. There should be functionality to add
    buttons in real-time, need to be backed up in persistent database."""

    # Todo: embed this list of strings within grammar to save space
    # mapping of keyboard keys to virtual key code to send as key input
    # VK_SPACE,VK_UP,VK_DOWN,VK_LEFT,VK_RIGHT,VK_RETURN,VK_BACK
    kmap = {
        'space': 0x20, 'up': 0x26, 'down': 0x28, 'left': 0x25, 'right': 0x27,
        'enter': 0x0d, 'backspace': 0x08, 'delete': 0x2e, 'leftclick': 0x201,
        'rightclick': 0x204, 'doubleclick': 0x202}

    nullTitles = ['Default IME', 'MSCTFIME UI', 'Engine Window',
                  'VDct Notifier Window', 'Program Manager',
                  'Spelling Window', 'Start']

    # dictionary of application objects, preferably read from file
    appDict = {}
    appDict.update({"iphoneWin": mu.AppWindow(["tans-iPhone",
                                            "host210.msm.che.vodafone"], None)})
    appDict.update({"xbmcChromeWin": mu.AppWindow(["XBMC - Google Chrome",], None)})
    # appSelectionStr = '(' + str(appDict.keys()).strip('][').replace(',','|') +\
    #')'
    appSelectionStr = None

    windows = mu.Windows(appDict=appDict, nullTitles=nullTitles)

    # Todo: embed this list of strings within grammar to save space
    # list of android screencast buttons
    # InitialiseMouseGrid coordination commands
    windows.appDict["iphoneWin"].mimicCmds.update(
    #logging.debug(appDict["iphoneWin"].mimicCmds)
        {'back': ['one'],
         'cancel': ['three'],
         'personal hotspot toggle': [],
         'show': [],
         'home': [],
         'end': ['seven', 'nine', 'two'],
         'answer': ['nine', 'seven', 'two'],
         'call': ['seven', 'five', 'eight'],
         'messages': ['one', 'five', 'eight'],
         'settings': ['six', 'eight'],
         'maps': ['five','one', 'two'],
         'maps search': ['two'],
         'message ok': ['five', 'eight'],
         'dismiss': ['five', 'eight'],
         'bluetooth on': ['three', 'eight', 'two', 'eight'],
         # drag context
         'drag up': ['eight','two','eight'],
         'drag down': ['two','eight', 'two'],
         # call context
         'contacts': ['eight', 'eight'],
         'recent': ['seven', 'nine'],
         'keypad': ['nine', 'seven'],
         # messages context
         'new message': ['three', 'six'],
         'send message middle': ['six', 'eight'],
         'send message bottom': ['nine', 'eight'],
         'message text middle': ['five', 'eight'],
         'message text bottom': ['eight', 'eight'], #five', 'eight'],
         # recent context
         'view all': ['two', 'four'],
         'view missed': ['two', 'six'],
         # contacts context
         'search text': ['two','eight','two'],
         'select entry': [],
         'first number': ['five','two'],
         'delete contact': ['eight',],
         # keypad context
         'keypad call': ['eight',],
         # in call keypad context
         'incall keypad': ['five',],
         'keypad zero': [ 'eight', 'two',],
         'keypad one': [ 'four', 'three', 'eight',],
         'keypad two': ['five', 'two', 'eight',],
         'keypad tree': ['six', 'one', 'eight',],
         'keypad four': ['four', 'six',],
         'keypad five': ['five', 'five',],
         'keypad six': ['six', 'four',],
         'keypad seven': ['four', 'nine',],
         'keypad eight': ['five', 'eight',],
         'keypad nine': ['six', 'seven',],
         'keypad star': [ 'seven', 'three',],
         'keypad hash': ['nine', 'one',],
         })
    ## Note: recognition seems to be dependent on the numbers being spelt out in
    # words. Button location macros/strings should be persisted in file or
    # database.
    ## Todo: button presses can be executed using assistive touch (transparent
    #white circle, for example device >>  lock screen long press can be used to
    #turn off (needs slide motion macro as well)

    # List of buttons
    appButtonStr = '|'.join(appDict["iphoneWin"].mimicCmds.keys())

    gramSpec = """
        <iphoneselect> exported = iphone select entry ({0});
        <iphonetap> exported = iphone ({1});
    """.format(str(range(20)).strip('[]').replace(', ','|'),appButtonStr)

    def gotResults_iphonetap(self, words, fullResults):
        appName = 'iphoneWin'
        retries = 3
        for i in xrange(retries):
            # return index ofapplication window title
            if self.windows.winDiscovery(appName=appName)[0]:
                # supplied the key of the intended window name
                return self.winAction(words[1:], appName)
            else:
                logging.debug("iphone window not found")
                # window doesn't exist, might need to start USB tunnel application
                # as well as vnc
                if 'itunnel_mux.exe' not in [c.Name for c in wmi.WMI().Win32_Process()]:
                    itun_p = Popen(["C:\win scripts\iphone usb.bat", "&"])
                # need to supply executable string (so-can locate the Windows
                # executable, it's not a Python executable) and then the
                # configuration file which has the password stored (doesn't seem to
                # support command line supplied password).
                # vnc_p=Popen('C:\\Program Files (x86)\\TightVNC\\vncviewer.exe' +\
                #            ' localhost:5904 -password test')
                vnc_p = Popen([r'C:\Program Files (x86)\TightVNC\vncviewer.exe',
                            '-config', r'C:\win scripts\localhost-5904.vnc'])
                # wait for creation
                logging.debug("waiting for process creation")
                time.sleep(2)
                # window should now exist, discover again
        logging.info('could not connect with phone ')

    def gotResults_iphoneselect(self, words, fullResults):
        """ Gives coordinates of an entry in a list on the iPhone. Receives the
        number of entries in the list (actually how many entries, given the
        size on the screen, would fit into the screen dimensions), the offset
        index of the first usable entry from top (how many entries Could fit
        above the first usable entry, give the index of the first usable entry)
        and the index of the desired entry. """
        cmdwords= words[:3]
        self.gotResults_iphonetap(cmdwords, fullResults)
        # now target window should be in focus and ready
        # static variables:
        appName = "iphoneWin"; num_entries = 14; offset_index = 3; select_int = 1
        select_int=(int(words[3]))
        # can use global variables populated by iphonetap
        hwin = self.windows.appDict[appName].winHandle
        if hwin:
            x,y,x1,y1 = wg.GetWindowRect(hwin)
            logging.debug('window Rect: %d,%d,%d,%d'% (x,y,x1,y1))
            x_ofs = x + (x1 - x)/2
            y_inc = (y1 - y)/num_entries
            y_ofs = y + y_inc/2 + (select_int + offset_index - 1)*y_inc
            logging.debug('horizontal: %d, vertical: %d, vertical increments: %d'%
                      (x_ofs,y_ofs,y_inc))
            if (select_int + offset_index - num_entries - 1) <= 0:
                # if entering search text,hide keypad
                playString('{enter}',0)
                self.click('leftclick',x=x_ofs,y=y_ofs,appName=appName)
        return

# use playstring instead
#    def press(self, key='space'):
#        event = self.kmap[key]
#        natlink.playEvents([(wm_keydown, event, 0),(wm_keyup, event, 0)])

    def drag(self, dragDirection='up', startPos=None, dist=None):
        recognitionMimic(['mouse', 'drag', dragDirection])
        # let the user stop as normal with voice...

    def click(self, clickType='leftclick', x=None, y=None, appName='iphoneWin'):
        # get the equivalent event code of the type of mouse event to perform
        # leftclick, rightclick, rightdouble-click (see kmap)
        event = self.kmap[clickType]
        # play events down click and then release (for left double click
        # increment from left button up event which produces no action
        # then when incremented, performs the double-click)
        # if coordinates are not supplied, just click
        if getattr(event, 'conjugate'):
            if not (x or y):
                x, y = getCursorPos()
            # apply vertical offset dependent on presence of "personal hotspot"
            # bar across the top of the screen
            y += self.windows.appDict[appName].vert_offset
            logging.debug('clicking at: %d, %d'% (x,y))
            natlink.playEvents(
                [(wm_mousemove, x, y), (event, x, y), (event + 1, x, y)])
        else:
            logging.error(' incorrect click look up for the event %s'% str(clickType))
            # default to
            recognitionMimic(['mouse', 'click'])

    def winAction(self, actionKey='', appName='iphoneWin'):
        # concatenate actionKey
        if getattr(actionKey, 'insert'):
            actionKey = ' '.join(actionKey)
            logging.debug("action Key of command concatenated: %s"% actionKey)
        # assuming the correct window is in focus
        # wake. Recognition mimic doesn't seem to be a good model. Something to
        # do with speed ofplayback etc. Grammar not always recognised as a
        # command.
        playString('{space}', 0x00)
        app = self.windows.appDict[str(appName)]
        gramList = []
        if str(actionKey) in app.mimicCmds:
            # we want to get out of grid mode aftermouse positioning
            # special cases first.
            if str(actionKey) == 'home':
                recognitionMimic(['mouse', 'window'])
                recognitionMimic(['go'])
                self.click('rightclick',appName=appName)
            elif str(actionKey) == 'personal hotspot toggle':
                if app.vert_offset:
                    app.vert_offset = 0
                else:
                    app.vert_offset = app.TOGGLE_VOFFSET
                logging.info("Toggled vertical offset, before: %d, after: %d"%
                            (old, app.vert_offset))
            elif str(actionKey).startswith("select"):
                pass # function continued in its own handler
            elif str(actionKey).startswith("show"):
                pass
            elif str(actionKey).startswith("drag"):
                recognitionMimic(['mouse', 'window'])
                gramList = app.mimicCmds[actionKey]
                logging.info("Grammer list for action '{0}': {1}".format(
                    actionKey, gramList))
                recognitionMimic(gramList)
                recognitionMimic(['go'])
                self.drag(dragDirection=actionKey.split()[1])
            else:
                recognitionMimic(['mouse', 'window'])
                gramList = app.mimicCmds[actionKey]
                logging.info("Grammer list for action '{0}': {1}".format(
                    actionKey, gramList))
                recognitionMimic(gramList)
                recognitionMimic(['go'])
                self.click('leftclick',appName=appName)
            return 0
        else:
            logging.error('unknown actionKey')
            return 1
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
