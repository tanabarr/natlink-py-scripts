#
# Python Macro Language for Dragon NaturallySpeaking
# nextApp
#

import natlink
from natlinkutils import *
import win32gui as wg
import os
from subprocess import Popen
import ioutils as iou
import logging
import time
import wmi
from VocolaUtils import *

logging.basicConfig(level=logging.INFO)

class ThisGrammar(GrammarBase):

    """ Class uses application window objects to store grid reference of window
    buttons on Windows without "say what you see" Dragon NaturallySpeaking
    functionality. To be developed to use the either grid reference or
    MouseGrid coordinate utterances. There should be functionality to add
    buttons in real-time, need to be backed up in persistent database.
    with iOS seven use the "say" functionality to dictate, avoids key
    latching over the VNC connection """

    # function mappings to process coordinates for drag action
    dragDirMapx = {
        'right': lambda x,d: x+d, 'left': lambda x,d: x-d,
        'up': lambda x,d: x, 'down': lambda x,d: x,
    }
    dragDirMapy = {
        'right': lambda y,d: y, 'left': lambda y,d: y,
        'up': lambda y,d: y-d, 'down': lambda y,d: y+d
    }
    dispMap = {
        'right': 190,'left': 190,'up': 250,'down':250
    }

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
    appDict.update({"iphoneWin": iou.AppWindow(["tans-iPhone",
                                               "tans-iphone.local",
                                               "host210.msm.che.vodafone"],
                                               None)})
    appDict.update({"xbmcChromeWin": iou.AppWindow(["XBMC - Google Chrome",], None)})
    # appSelectionStr = '(' + str(appDict.keys()).strip('][').replace(',','|') +\
    #')'
    appSelectionStr = None

    windows = iou.Windows(appDict=appDict, nullTitles=nullTitles)

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
         'wake': [],
         'trust': ['four', 'nine'],
         'switch screens': ['eight'],
         'end': ['seven', 'nine', 'two', 'two',],
         'answer': ['nine', 'seven', 'two', 'two'],
         'call': ['seven', 'five', 'eight'],
         'messages': ['one', 'five', 'eight'],
         'app store': ['five', 'seven'],
         'settings': ['six', 'eight'],
         'maps': ['five','one', 'two'],
         'maps text': ['two'],
         'message ok': ['five', 'eight'],
         'dismiss': ['five', 'eight','eight', 'eight'],
         'bluetooth on': ['three', 'eight', 'five', 'eight'],
         # drag context
         'drag up': ['eight','two','eight'],
         'drag down': ['two','eight', 'two'],
         'drag left': ['nine','eight','three'],
         'drag right': ['seven','eight','three'],
         # call context
         'contacts': ['eight', 'eight'],
         'recent': ['seven', 'nine'],
         'favourites': ['seven', 'eight', 'four'],
         'keypad': ['nine', 'seven'],
         # messages context
         'new message': ['three', 'six'],
         'send message middle': ['six', 'eight', 'six'],
         'send message bottom': ['nine', 'eight'],
         'message text middle': ['five', 'eight'],
         'message text bottom': ['eight', 'eight'], #five', 'eight'],
         # recent context
         'view all': ['two', 'four'],
         'view missed': ['two', 'six'],
         # contacts list context
         'search text': ['two','eight','two'],
         ## separate grammar for the below two
         #'select entry': [],
         #'select entry details': [],
         # contact context
         'first number': ['five','two'],
         'delete contact': ['eight',],
         'contact send message': ['four' ,'eight'],
         # keypad context
         'keypad call': ['eight',],
         # in call keypad context
         'incall keypad': ['five','two','two'],
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
         '3478 contact call': [],
         '858 contact send message': [],
         '852 contact voice call': [],
         '98 call voicemail': [],
         })
    ## Note: recognition seems to be dependent on the numbers being spelt out in
    # words. Button location macros/strings should be persisted in file or
    # database.
    ## Todo: button presses can be executed using assistive touch (transparent
    #white circle, for example device >>  lock screen long press can be used to
    #turn off (needs slide motion macro as well)

    num_to_word =\
    ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine"]

    # translate shorthand keys into proper dictionary entry
    # requires mapping numbers to words
    cDict = appDict["iphoneWin"].mimicCmds
    for k in cDict.keys():
      if str(k)[0].isdigit():
        coordinates,command=str(k).split(' ',1)
        coordinates_list= [num_to_word[int(i)] for i in coordinates]
        print command
        print coordinates_list
        del cDict[k]
        cDict.update({command: coordinates_list})

    def grid(self, coordinates):
        """ replace the MouseGrid recognition mimic calls with function
        calculate a destination coordinate """
        pass

    # List of buttons
    appButtonStr = '|'.join(cDict.keys())

    gramSpec = """
        <iphoneselect> exported = iphone select entry ({0}) [details];
        <iphonetap> exported = iphone ({1});
    """.format(str(range(20)).strip('[]').replace(', ','|'),appButtonStr)

    def gotResults_iphonetap(self, words, fullResults):
        appName = 'iphoneWin'
        retries = 2 #3
        for i in xrange(retries):
            # return index ofapplication window title
            if self.windows.winDiscovery(appName=appName)[0]:
                print "return from function discovery"
                # supplied the key of the intended window name
                return self.winAction(words[1:], appName)
            else:
                print(str(self.__module__) +  "debug: iphone window not found")
                # window doesn't exist, might need to start USB tunnel application
                # as well as vnc
                if 'itunnel_mux.exe' not in [c.Name for c in wmi.WMI().Win32_Process()]:
                    print(str(self.__module__) +  "debug: itunnel process not found, starting...")
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
                print(str(self.__module__) +  "debug: waiting for process creation")
                time.sleep(2)
                # window should now exist, discover again
        print(str(self.__module__) +  "info:" 'could not connect with phone ')

    def gotResults_iphoneselect(self, words, fullResults):
        """ Gives coordinates of an entry in a list on the iPhone. Receives the
        number of entries in the list (actually how many entries, given the
        size on the screen, would fit into the screen dimensions), the offset
        index of the first usable entry from top (how many entries Could fit
        above the first usable entry, give the index of the first usable entry)
        and the index of the desired entry. TODO: click on contact buttons on
        the right of entries """
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
            if len(words) > 4 and str(words[4]) == 'details':
                # selectable blue Arrow on the right side of the contact
                x_ofs = x + 17*(x1 - x)/18
            else:
                # otherwise click in the centre (horizontal)
                x_ofs = x + (x1 - x)/2
            y_inc = (y1 - y)/num_entries
            y_ofs = y + y_inc/2 + (select_int + offset_index - 1)*y_inc
            logging.debug('horizontal: %d, vertical: %d, vertical increments: %d'%
                      (x_ofs,y_ofs,y_inc))
            if (select_int + offset_index - num_entries - 1) <= 0:
                # if entering search text,hide keypad
                natlink.playString('{enter}',0)
                self.click('leftclick',x=x_ofs,y=y_ofs,appName=appName)
        return

# use playstring instead
#    def press(self, key='space'):
#        event = self.kmap[key]
#        natlink.playEvents([(wm_keydown, event, 0),(wm_keyup, event, 0)])
    def sanitise_movement(func):
        def checker(*args,**kwargs):
            print args
            print kwargs
            ret = func(*args,**kwargs)
            return ret
        return checker

    def getDragPoints(self,x,y,displacement,dragDirection):
        """ take current cursor position, displacement and direction,
        receives starting coordinates and returns end coordinates.
        axis displacement is in direction of "dragDirection". e.g. to Drag right;
        Place mouse at beginning of desired Drag action, add displacement
        from x coordinate and return the new coordinates """

        return (self.dragDirMapx[dragDirection](x,displacement),
                   self.dragDirMapy[dragDirection](y,displacement))

    #@sanitise_movement
    def drag(self, dragDirection='up', startPos=None, dist=4):
        displacement = self.dispMap[dragDirection]
        call_Dragon('RememberPoint', '', [])
        x, y = natlink.getCursorPos()
        newx,newy = self.getDragPoints(x,y,displacement,dragDirection)
        natlink.playEvents([(wm_mousemove, newx, newy)])
        call_Dragon('DragToPoint', 'i', [])

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
                x, y = natlink.getCursorPos()
            # apply vertical offset dependent on presence of "personal hotspot"
            # bar across the top of the screen
            y += self.windows.appDict[appName].vert_offset
            logging.debug('clicking at: %d, %d'% (x,y))
            natlink.playEvents(
                [(wm_mousemove, x, y), (event, x, y), (event + 1, x, y)])
        else:
            logging.error(' incorrect click look up for the event %s'% str(clickType))
            # default to
            natlink.recognitionMimic(['mouse', 'click'])

    def winAction(self, actionKey='', appName='iphoneWin'):
        print "action"
        # concatenate actionKey
        if getattr(actionKey, 'insert'):
            actionKey = ' '.join(actionKey)
            print(str(self.__module__) +  "debug: action Key of command concatenated: %s"% actionKey)
        # assuming the correct window is in focus
        # wake. Recognition mimic doesn't seem to be a good model. Something to
        # do with speed ofplayback etc. Grammar not always recognised as a
        # command.
        natlink.playString('{space}', 0x00)
        app = self.windows.appDict[str(appName)]
        gramList = []
        if str(actionKey) in app.mimicCmds:
            # we want to get out of grid mode aftermouse positioning
            # special cases first.
            if str(actionKey) == 'home':
                natlink.recognitionMimic(['mouse', 'window'])
                natlink.recognitionMimic(['go'])
                self.click('rightclick',appName=appName)
            elif str(actionKey) == 'wake':
                natlink.recognitionMimic(['mouse', 'window'])
                natlink.recognitionMimic(['go'])
                self.click('rightclick',appName=appName)
                actionKey = 'drag right'
                natlink.recognitionMimic(['mouse', 'window'])
                gramList = app.mimicCmds[actionKey]
                print(str(self.__module__) + "debug: Grammer list for action '{0}': {1}".format(
                    actionKey, gramList))
                natlink.recognitionMimic(gramList)
                natlink.recognitionMimic(['go'])
                self.drag(dragDirection=actionKey.split()[1], dist=2)
            elif str(actionKey) == 'personal hotspot toggle':
                if app.vert_offset:
                    app.vert_offset = 0
                else:
                    app.vert_offset = app.TOGGLE_VOFFSET
                print(str(self.__module__) + "debug: Toggled vertical offset, before: %d, after: %d"%
                            (old, app.vert_offset))
            elif str(actionKey).startswith("select"):
                pass # function continued in its own handler
            elif str(actionKey).startswith("show"):
                pass
            elif str(actionKey).startswith("drag"):
                natlink.recognitionMimic(['mouse', 'window'])
                gramList = app.mimicCmds[actionKey]
                print(str(self.__module__) + "debug: Grammer list for action '{0}': {1}".format(
                    actionKey, gramList))
                natlink.recognitionMimic(gramList)
                natlink.recognitionMimic(['go'])
                self.drag(dragDirection=actionKey.split()[1])
            else:
                natlink.recognitionMimic(['mouse', 'window'])
                gramList = app.mimicCmds[actionKey]
                print(str(self.__module__) + "debug: Grammer list for action '{0}': {1}".format(
                    actionKey, gramList))
                natlink.recognitionMimic(gramList)
                natlink.recognitionMimic(['go'])
                self.click('leftclick',appName=appName)
            return 0
        else:
            print(str(self.__module__) +  'error:unknown actionKey')
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
