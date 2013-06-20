#
# Python Macro Language for Dragon NaturallySpeaking
# nextApp
#

import natlink
from natlinkutils import *
import win32gui as wg
import os
from subprocess import Popen
import logging as log
import time
import wmi

log.basicConfig(level=log.INFO)

class AppWindow():

    def __init__(self, name, rect=None, hwin=None):
        self.winName = name
        self.winRect = rect
        self.winHandle = hwin
        buttons = ['home', 'menu', 'back', 'search', 'call', 'end']
        self.mimicCmds = {}.fromkeys(buttons)


class ThisGrammar(GrammarBase, AppWindow):

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

    # dictionary of application objects, preferably read from file
    appDict = {}
    appDict.update({"iphoneWin": AppWindow("tans-iPhone", None)})
    appDict.update({"xbmcChromeWin": AppWindow("XBMC - Google Chrome", None)})
    # appSelectionStr = '(' + str(appDict.keys()).strip('][').replace(',','|') +\
    #')'
    appSelectionStr = None

    def selectEntry(self, num_entries, offset_index, select_index):
        """ Gives coordinates of an entry in a list on the iPhone. Receives the
        number of entries in the list (actually how many entries, given the
        size on the screen, would fit into the screen dimensions), the offset
        index of the first usable entry from top (how many entries Could fit
        above the first usable entry, give the index of the first usable entry)
        and the index of the desired entry. """
        x,y = getWindowSize()
        #res = setCursorPos()
        return

    # Todo: embed this list of strings within grammar to save space
    # list of android screencast buttons
    # InitialiseMouseGrid coordination commands
    appDict["iphoneWin"].mimicCmds.update(
    #log.debug(appDict["iphoneWin"].mimicCmds)
        {'back': ['one', 'five', 'eight'],
         'home': [],
         'end': ['seven', 'nine'],
         'answer': ['nine', 'seven'],
         'call': ['seven', 'five', 'eight'],
         'messages': ['one', 'five', 'eight'],
         'settings': ['six', 'two'],
         # call context
         'contacts': ['eight', 'eight'],
         'recent': ['seven', 'nine'],
         'keypad': ['nine', 'seven'],
         # messages context
         'new message': ['three', 'six'],
         'send message': ['nine', 'eight'],
         'message text': ['five', 'eight'],
         # recent context
         'view all': ['two', 'four'],
         'view missed': ['two', 'six'],
         # contacts context
         'search': ['two','eight','two'],
         'first number': ['five','two'],
         })
    ## Note: recognition seems to be dependent on the numbers being spelt out in
    # words. Button location macros/strings should be persisted in file or
    # database.

    # log.debug(appDict["iphoneWin"].mimicCmds) #iphoneCmdDict)
    # List of buttons
    appButtonStr = '|'.join(appDict["iphoneWin"].mimicCmds.keys())

    gramSpec = """
        <winclick> exported = click ({0}) ({1});
        <iphonetap> exported = tap iphone ({1});
    """.format(appSelectionStr, appButtonStr)
    print gramSpec

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

        appWin = 'iphoneWin'
        retries = 3
        for i in xrange(retries):
            if self.winDiscovery(words, appWin)[0]:
                # we now want to call the window action function with the
                # "tapRelative"
                #log.debug(words)
                # log.debug(self.appDict[appWin].mimicCmds[words[2]],'tapRelative')
                # supplied the key of the intended window action
                self.winAction(words[2:]) #, 'tapRelative')
                return 0
            else:
                log.debug("iphone window not found")
        log.info('could not connect with phone ')

    def winDiscovery(self, words, appWin=None):
        # argument to pass to callback contains words used in voice command
        # (this is also a recognised top-level window?) And title of window
        # to find (optional). Return tuple (handle of appWin, window dict).
        wins = (words, {})
        hwin = None
        # selecting index from bottom window title on taskbar
        # enumerate all top-level windows and send handles to callback
        wg.EnumWindows(self.callBack_popWin, wins)
        # after visible taskbar application windows have been added to
        # dictionary (second element of wins tuple), we can calculate
        total_windows = len(wins[1])
        # print('Number of taskbar applications: {0};'.format( total_windows))
        # the function called without selecting window to find, just return
        # window dictionary.
        if not appWin:
            return (None, wins[1])
        #log.debug("discover window %s" % appWin)
        # trying to find window title of selected application within window
        # dictionary( local application context). Checking that the window
        # exists and it has a supportive local application context.
        try:
            app = self.appDict[str(appWin)]
            #log.debug("supported application window: %s" % app.winName)
            index = wins[1].values().index(app.winName)
            # need to find a list of Windows again (to refresh)
        except:
            index = None
        if index is not None:
            #log.debug("index of application window: %d" % index)
            hwin = (wins[1].keys())[index]
            log.debug(
                "Name: {0}, Handle: {1}".format(wins[1][hwin], str(hwin)))
            app.winHandle = hwin
            wg.SetForegroundWindow(hwin)
            # print wg.GetWindowRect(hwin)
            app.winRect = wg.GetWindowRect(hwin)
            # print str(hwin)
            return (str(hwin), wins[1])
        else:
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
            # vncwas crashing due to budget dodgy cable
            # wait for creation
            log.debug("waiting for process creation")
            time.sleep(2)
            # window should now exist, discover again
            return (False, wins[1])

# use playstring instead
#    def press(self, key='space'):
#        event = self.kmap[key]
#        natlink.playEvents([(wm_keydown, event, 0),(wm_keyup, event, 0)])

    def click(self, clickType='leftclick', x=None, y=None):
        # get the equivalent event code of the type of mouse event to perform
        # leftclick, rightclick, rightdouble-click (see kmap)
        event = self.kmap[clickType]
        # play events down click and then release (for left double click
        # increment from left button up event which produces no action
        # then when incremented, performs the double-click)
        # if coordinates are not supplied, just click
        #log.debug(dir(event))
        if getattr(event, 'conjugate'):
            if not (x or y):
                x, y = getCursorPos()
            log.debug('clicking at: %d, %d'% (x,y))
            natlink.playEvents(
                [(wm_mousemove, x, y), (event, x, y), (event + 1, x, y)])
        else:
            log.error(' incorrect click look up for the event %s'% str(clickType))
            # default to
            recognitionMimic(['mouse', 'click'])

    def winAction(self, actionKey='', actionType='tapRelative'):
        # concatenate actionKey
        if getattr(actionKey, 'insert'):
            actionKey = ' '.join(actionKey)
            log.debug("action Key of command concatenated: %s"% actionKey)
        # assuming the correct window is in focus
        # wake. Recognition mimic doesn't seem to be a good model. Something to
        # do with speed ofplayback etc. Grammar not always recognised as a
        # command.
        playString('{space}', 0x00)
        # todo: how to reset state machine, start from home screen without
        # feedback?
        gramList = newgramList = []
        if str(actionKey) in self.appDict["iphoneWin"].mimicCmds:
            recognitionMimic(['mouse', 'window'])
            gramList = self.appDict["iphoneWin"].mimicCmds[actionKey]
            # we want to get out of grid mode aftermouse positioning
            if gramList is not None:
                # remember empty list is not evaluated as "None"
                if str(actionKey) == 'home':
                    # special case
                    #recognitionMimic(['mouse', 'right', 'click'])
                    recognitionMimic(['go'])
                    self.click('rightclick')
                else:
                    newgramList = gramList
                    log.info("Grammer list for action '{0}': {1}".format(
                        actionKey, newgramList))
                    recognitionMimic(newgramList)
                    recognitionMimic(['go'])
                    #time.sleep(1)
                    #recognitionMimic(['mouse', 'click'])
                    self.click('leftclick')
            else:
                log.error('grammar list missing')
        else:
            log.error('unknown actionKey')

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
