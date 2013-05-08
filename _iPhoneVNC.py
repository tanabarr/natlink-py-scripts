#
# Python Macro Language for Dragon NaturallySpeaking
# nextApp
#

import natlink
from natlinkutils import *
import win32gui as wg
import os
from subprocess import Popen
import logging

logging.basicConfig(level=logging.DEBUG)

class AppWindow():
    def __init__(self, name, rect=None, hwin=None):
        self.winName = name
        self.winRect = rect
        self.winHandle = hwin
        buttons = ['home', 'menu', 'back', 'search', 'call', 'endcall']
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
    kmap = {'space': 0x20, 'up': 0x26, 'down': 0x28, 'left': 0x25, 'right': 0x27,
            'enter': 0x0d, 'backspace': 0x08, 'delete': 0x2e, 'leftclick': 0x201,
            'rightclick': 0x204, 'doubleclick': 0x202}

    # dictionary of application objects, preferably read from file
    appDict = {}
    appDict.update({"iphoneWin": AppWindow("tans-iPhone", None)})
    appDict.update({"xbmcChromeWin": AppWindow("XBMC - Google Chrome", None)})
    appSelectionStr = '(' + str(appDict.keys()).strip('][').replace(',','|') +\
    ')'
    # Todo: embed this list of strings within grammar to save space
    # list of android screencast buttons
    # InitialiseMouseGrid coordination commands
    iphoneCmdDict = appDict["iphoneWin"].mimicCmds
    logging.debug(appDict["iphoneWin"].mimicCmds)
    iphoneCmdDict.update({'back': ['1','5','8']})
    iphoneCmdDict.update({'call': ['7','5','8']})
    iphoneCmdDict.update({'contacts': ['8','8']})
    logging.debug(appDict["iphoneWin"].mimicCmds) #iphoneCmdDict)
    # List of buttons
    appButtonStr = '(' + str(appDict["iphoneWin"].mimicCmds.keys()).strip(']['
                    ).replace(',','|') + ')'

    logging.debug(appSelectionStr,appButtonStr)

    gramSpec = """
        <winclick> exported = click {0} {1};
        <iphonetap> exported = tap iphone {1};
    """.format(appSelectionStr,appButtonStr)
#        <iphonetap> exported = tap iphone (home|menu|back|search|call|endcall);


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

        appWin = 'iphoneWin'
        retries = 1
        for i in xrange(retries):
            if self.winDiscovery(words, appWin)[0]:
                # we now want to call the window action function with the "tapRelative"
                logging.debug(words)
                logging.debug(self.appDict[appWin].mimicCmds[words[2]],'tapRelative')
                self.winAction(self.iphoneCmdDict[words[2]],'tapRelative')
            else:
                logging.debug("iphone window not found")
                #os.

    def winDiscovery(self, words, appWin=None):
        # argument to pass to callback contains words used in voice command
        # (this is also a recognised top-level window?) And title of window
        # to find (optional). Return tuple (handle of appWin, window dict).
        wins = (words, {})
        hwin = None
        # selecting index from bottom window title on taskbar
        # enumerate all top-level windows and send handles to callback
        wg.EnumWindows(self.callBack_popWin,wins)

        # after visible taskbar application windows have been added to
        # dictionary (second element of wins tuple), we can calculate
        # relative offset from last taskbar title.
        total_windows = len(wins[1])
        #print('Number of taskbar applications: {0};'.format( total_windows))
        #print wins[1]

        # the function called without selecting window to find, just return
        # window dictionary.
        if not appWin: return (None, wins[1])

        logging.debug(appWin)
        # trying to find window title of selected application within window
        # dictionary
        try:
            app = self.appDict[str(appWin)]
            logging.debug(app.winName)
            index = wins[1].values().index(app.winName)
        except:
            index = None

        if index is not None:
            logging.debug(index)
            hwin = (wins[1].keys())[index]
            logging.debug("Name: {0}, Handle: {1}".format(wins[1][hwin], str(hwin)))
            app.winHandle = hwin
            wg.SetForegroundWindow(hwin)
            #print wg.GetWindowRect(hwin)
            app.winRect = wg.GetWindowRect(hwin)
            #print str(hwin)
            return (str(hwin), wins[1])
        else:
            # window doesn't exist, might need to start USB tunnel application
            # as well as vnc
            itun_p=Popen(["C:\win scripts\iphone usb.bat", "&"])
            try:
                stdout, stderr=vnc_p.communicate(timeout=2)
            except: # TimeoutError:
                logging.debug("error")
            # need to supply executable string (so-can locate the Windows
            # executable, it's not a Python executable) and then the
            # configuration file which has the password stored (doesn't seem to
            # support command line supplied password).
            #vnc_p=Popen('C:\\Program Files (x86)\\TightVNC\\vncviewer.exe' +\
            #            ' localhost:5904 -password test')
            vnc_p=Popen('C:\\Program Files (x86)\\TightVNC\\vncviewer.exe' +\
                        ' -config \"C:\\win scripts\\Mobile screen.vnc\"')
            try:
                stdout, stderr=vnc_p.communicate(timeout=2)
            except: # TimeoutError:
                logging.debug("error vnc")
            # window should now exist, discover again
            self.winDiscovery(words, appwin)

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
        if x and y:
            natlink.playEvents([(wm_mousemove, x, y), (event, x, y), (event + 1, x, y)])
        else:
            recognitionMimic(["MouseGrid", "window"])
            recognitionMimic(["mouse", str(clickType)])

    def winAction(self, gramList, actionType):
        # assuming the correct window is in focus
        # wake
        playString('{space}',0x00)
        #todo: how to reset state machine, start from home screen without
        #feedback?
        #self.click(clickType='rightclick')
        gramList=['MouseGrid', 'window'] + gramList
        logging.info("Grammer list {0} ".format(gramList))
        #recognitionMimic(['MouseGrid', 'window'] + gramList)
        recognitionMimic(gramList)
        #recognitionMimic(["MouseGrid", "window", "8", "5", "8"]) # ] + gramList)

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
