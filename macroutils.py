#
# macro utilities for creating macro-behaviours
#
import logging
import sqlite3
import win32gui as wg

logging.basicConfig(level=logging.ERROR)

class MacroObj():
    def __init__(self,string='',flags=0):
        self.string=string
        self.flags=flags

class FileStore():
    """ manage the storage of configurations to file """

    #logger=logging.getLogger('')

    def __init__(self,defaults_filename='defaults.conf',
                 updates_filename='updates.conf',
                 working_directory='C:\\NatLink\\NatLink\\MacroSystem\\',
                 preDict={},delim="|"):
        self.updates_filename=updates_filename
        self.postDict=preDict
        self.delimchar=delim
        self.wd=working_directory
        logging.info("opening %s" % self.wd+defaults_filename)
        try:
            self.defaults_fd=open(self.wd + defaults_filename,'r')
            logging.info("reading from...")
            self.readfile(self.defaults_fd)
            logging.info("%d macros read from..." % len(self.postDict))
        except:
            logging.error('could not open default configuration: %s'
                        % defaults_filename)

    def readfile(self, fd):
        #paramDict={}
        for line in fd.readlines():
            #param, value=line.split(self.delimchar, 1)
            #paramDict.update({param.strip(): value.strip()})
#            logging.debug("reading %s" % line)
            try:
#                logging.debug(str(line.split('|',3)))
                gram, macro, flags = line.split('|')
                #logging.debug("%s  %s  %s" % (str(gram), str(macro), str(flags).strip(r'r\n')))
                self.postDict[gram] = MacroObj(macro, int(flags.strip(r'r\n ')))
            except:
                logging.info("%s line not a macro entry" % line)

    def writefile(self, output_filename='output.conf'):
        outfile_fd=open(self.wd + output_filename,'w')
        #eogging.debug('keys: %s' % self.postDict.keys())
        outfile_fd.write('keyboard macro (name, string, flag) tuples')
        for gram, macroobj in self.postDict.iteritems():
            try:
                outfile_fd.write('\r\n' + str(self.delimchar).join([gram,
                                                   macroobj.string,
                                                   str(macroobj.flags)]))
            except:
                pass

class AppWindow:

    def __init__(self, names, rect=None, hwin=None):
        self.winNames = names
        self.winRect = rect
        self.winHandle = hwin
        self.vert_offset = 0
        self.TOGGLE_VOFFSET = 9
        buttons = ['home', 'menu', 'back', 'search', 'call', 'end']
        self.mimicCmds = {}.fromkeys(buttons)


class Windows:
    def __init__(self, appDict={}, nullTitles=[]):
        self.appDict=appDict
        self.nullTitles=nullTitles

    def _callBack_popWin(self, hwin, args):
        """ this callback function is called with handle of each top-level
        window. Window handles are used to check the of window in question is
        visible and if so it's title strings checked to see if it is a standard
        application (e.g. not the start button or natlink voice command itself).
        Populate dictionary of window title keys to window handle values. """
        if wg.IsWindowVisible(hwin):
            try:
                winText = wg.GetWindowText(hwin).strip()
                if winText and winText not in self.nullTitles and\
                    winText not in args.values():
                    args.update({hwin: winText})
            except:
                logging.error('cannot retrieve window title')
#               and filter(lambda x: x in args[0], winText.split()):

    def winDiscovery(self, appName=None):
        wins = {}
        hwin = None
        wg.EnumWindows(self._callBack_popWin, wins)
        total_windows = len(wins)
        print total_windows
        print wins
        # the function called without selecting window to find, just return
        # window dictionary.
        if not appName:
            return (None, wins)
        # trying to find window title of selected application within window
        # dictionary( local application context). Checking that the window
        # exists and it has a supportive local application context.
        app = self.appDict[str(appName)]
        namelist=[]
        # checking the window names is a list, handle string occurrence
        if getattr(app.winNames, 'append'):
            namelist=app.winNames
        else:
            namelist.append(app.winNames)
        for name in namelist:
            try:
                index = wins.values().index(name)
                break
            except:
                index = None

        if index is not None:
            #loggingdebug("index of application window: %d" % index)
            hwin = (wins.keys())[index]
            logging.debug(
                "Name: {0}, Handle: {1}".format(wins[hwin], str(hwin)))
            app.winHandle = hwin
            wg.BringWindowToTop(int(hwin))
            wg.SetForegroundWindow(int(hwin))
            # print wg.GetWindowRect(hwin)
            #app.winRect = wg.GetWindowRect(hwin)
            # print str(hwin)
            return (str(hwin), wins)
        else:
            return (False, wins)


