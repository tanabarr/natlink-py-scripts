#
# macro utilities for creating macro-behaviours
#
import logging
from sqlite3 import OperationalError, connect
from win32gui import IsWindowVisible, GetWindowText, EnumWindows, BringWindowToTop, SetForegroundWindow
#import os

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
                 working_directory='c:/Natlink/Natlink/MacroSystem/',
                 preDict={},delim="|",
                 db_filename='natlink.db',
                 schema=None):
        #print os.getcwd()
        self.updates_filename=updates_filename
        self.postDict=preDict
        self.delimchar=delim
        self.wd=working_directory
        if schema:
            logging.info("schema present")
            if not self.readdb(schema):
                return
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
        logging.info("reading from file...")
        for line in fd.readlines():
#            logging.debug("reading %s" % line)
            try:
#                logging.debug(str(line.split('|',3)))
                gram, macro, flags = line.split('|')
                #logging.debug("%s  %s  %s" % (str(gram), str(macro), str(flags).strip(r'r\n')))
                self.postDict[gram] = MacroObj(macro, int(flags.strip(r'r\n ')))
            except:
                logging.info("%s line not a macro entry" % line)

    def writefile(self, output_filename='output.conf'):
        logging.info("writing to file...")
        outfile_fd=open(self.wd + output_filename,'w')
        #eogging.debug('keys: %s' % self.postDict.keys())
        outfile_fd.write('keyboard macro (name, string, flag) tuples')
        for gram, macroobj in self.postDict.iteritems():
            try:
                outfile_fd.write('\n' + str(self.delimchar).join([gram,
                                                   macroobj.string,
                                                   str(macroobj.flags)]))
            except:
                pass

    def readdb(self, schema, db_filename='natlink.db', table_name='kb_macros'):
        logging.info("reading from database...")
        try:
            conn = connect(self.wd + db_filename)
            c = conn.cursor()
            col_names= schema.replace(' text','')
            c.execute("SELECT %s FROM %s" % (col_names, table_name))
            for row in c.fetchall():
                gram, macro_raw, flags = row
                macro=self.customdecodechars(macro_raw)
                self.postDict[gram] = MacroObj(macro, int(flags)) # .strip(r'r\n ')))
#                col_index=0
#                for col_name in col_names.split(','):
#                    #logging.info("col: %s=%s," % (col_name,row_decoded[col_index]))
#                    col_index+=1
            conn.close()
            return 0
        except:
            logging.info("error reading from database...")
            return 1
        # except sqlite3.OperationalError, err:
       #     logging.exception( "OperationalError: %s" % err)
       #     return 1

    def writedb(self, schema, db_filename='natlink.db', table_name='kb_macros'):
        logging.info("writing to database...")
        try:
            conn = connect(self.wd + db_filename)
            c = conn.cursor()
            # Create table
            c.execute("CREATE TABLE IF NOT EXISTS %s (%s)" % (table_name, schema))
            for gram, macroobj in self.postDict.iteritems():
                # Insert a row of data
                macro_string=self.customencodechar(macroobj.string)
                print gram, macroobj.string, macro_string
                c.execute("INSERT INTO %s (%s) VALUES ('%s', '%s', '%s')" %
                        (table_name, schema.replace(' text', ''),
                         gram, macro_string, str(macroobj.flags)))
                conn.commit()
        #except Exception, err:
            conn.close()
        except OperationalError, err:
            logging.exception( "OperationalError: %s" % err)

    def customencodechar(self, string):
        return string.replace("'","SNGL_QUOTE")

    def customdecodechars(self, string):
        if "SNGL_QUOTE" in string:
            new_string= str(string).replace("SNGL_QUOTE", "'")
            #logging.info("old %s, new %s" % (string, new_string))
        return string

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
        self.skipTitle=None

    def _callBack_popWin(self, hwin, args):
        """ this callback function is called with handle of each top-level
        window. Window handles are used to check the of window in question is
        visible and if so it's title strings checked to see if it is a standard
        application (e.g. not the start button or natlink voice command itself).
        Populate dictionary of window title keys to window handle values. """
        #print '.' #self.nullTitles
        #nullTitles = self.nullTitles.append(self.skipTitle)
        #print nullTitles
        if IsWindowVisible(hwin):
            winText = GetWindowText(hwin).strip()
            nt = self.nullTitles + [self.skipTitle,]
            if winText and winText not in nt: # and\
                # enable duplicates #winText not in args.values():
                if winText.count('WinSCP') and winText != 'WinSCP Login':
                    if winText in args.values():
                        return
                args.update({hwin: winText})
            #else:
            #    logging.error('cannot retrieve window title %s' % winText)
#               and filter(lambda x: x in args[0], winText.split()):

    def winDiscovery(self, appName=None, winTitle=None, beginTitle=None,
                     skipTitle=None):
        """ support finding and focusing on application window or simple window
        title. Find the index and focuses on the first match of any of these.
        Applications within the application dictionary could have a number
        window_titles associated. """
        wins = {}
        hwin = None
        index = None
        self.skipTitle = skipTitle
        EnumWindows(self._callBack_popWin, wins)
        self.skipTitle = None
        total_windows = len(wins)
        # creating match lists
        namelist=[]
        partlist=[]
        if winTitle:
            namelist.append(winTitle)
        elif beginTitle:
            for v in wins.values():
                if v.startswith(beginTitle):
                    namelist.append(v)

        if appName:
            # trying to find window title of selected application within window
            # dictionary( local application context). Checking that the window
            # exists and it has a supportive local application context.
            app = self.appDict[str(appName)]
            # checking the window names is a list, handle string occurrence
            try:
                if app.winHandle:
                    BringWindowToTop(int(hwin))
                    SetForegroundWindow(int(hwin))
                    return (str(hwin), wins)
            except:
                pass
            if getattr(app.winNames, 'append'):
                namelist = namelist + app.winNames
            else:
                namelist.append(app.winNames)
        for name in namelist:
            try:
                index = wins.values().index(name)
                break
            except:
                pass

        if index is not None:
            #loggingdebug("index of application window: %d" % index)
            hwin = (wins.keys())[index]
            logging.debug(
                "Name: {0}, Handle: {1}".format(wins[hwin], str(hwin)))
            try:
                app.winHandle = hwin
            except:
                pass
            BringWindowToTop(int(hwin))
            SetForegroundWindow(int(hwin))
            #app.winRect = wg.GetWindowRect(hwin)
            return (str(hwin), wins)
        else:
            return (None, wins)
