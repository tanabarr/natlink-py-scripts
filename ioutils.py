#
# macro utilities for creating macro-behaviours
#
import logging
from sqlite3 import OperationalError, connect
from win32gui import IsWindowVisible, GetWindowText, EnumWindows, BringWindowToTop, SetForegroundWindow
#import os

logging.basicConfig(level=logging.INFO)

### DECORATORS ###

def sanitise_movement(func):
    def checker(*args,**kwargs):
        print args
        print kwargs
        ret = func(*args,**kwargs)
        return ret
    return checker

### CLASSES ###

class MacroObj():
    def __init__(self,string='',flags=0):
        self.string=string
        self.flags=flags

UPDATES_USAGE = "#keyboard macro (name, string, flag) tuples\n" \
                "#flag '0xff' prevents macro preprocessing e.g. vim/screen\n"

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
#        self.updates_filename=updates_filename
        self.postDict=preDict
        self.delimchar=delim
        self.wd=working_directory
        count=0
        if schema:
            logging.info("schema present")
            count = self.readdb(schema)
        if not count:
            count = self.readfile(defaults_filename)
        if count:
            logging.info("%d macros"% count)
            count = self.readfile(updates_filename)
            print count
            if count:
                logging.info("%d updated macros"% count)
                count = self.writefile()
                logging.info("%d macros to file"% (count))
                if schema:
                    count = self.writedb(schema)
                    logging.info("%d macros to db, cleaning updates file"
                                 % (count))
                    with open(self.wd + updates_filename,'w') as myfile:
                        myfile.write(UPDATES_USAGE)
        else:
            logging.error('could not open : %s' %
                                  defaults_filename)

    def readfile(self, filename):
        logging.info("opening %s" % self.wd+filename)
        count=0
        try:
            with open(self.wd + filename,'r') as myfile:
                for line in myfile:
        #            logging.debug("reading %s" % line)
                    if not line.startswith('#'):
                        try:
            #                logging.debug(str(line.split('|',3)))
                            gram, macro, flags = line.split('|')
                            #logging.debug("%s  %s  %s" % (str(gram), str(macro), str(flags).strip(r'r\n')))
                            self.postDict[gram] = MacroObj(macro, int(flags.strip(r'r\n ')))
                            count+=1
                        except:
                            logging.info("%s line not a macro entry" % line)
        except:
            logging.error('could not open : %s' % filename)
        finally:
            return count

    def writefile(self, output_filename='output.conf'):
        logging.info("writing to file...")
        outfile_fd=open(self.wd + output_filename,'w')
        #eogging.debug('keys: %s' % self.postDict.keys())
        outfile_fd.write('keyboard macro (name, string, flag) tuples')
        count=0
        for gram, macroobj in self.postDict.iteritems():
            try:
                outfile_fd.write('\n' + str(self.delimchar).join([gram,
                                            macroobj.string,
                                            str(macroobj.flags)]))
                count+=1
            except:
                pass
        return count

    def readdb(self, schema, db_filename='natlink.db', table_name='kb_macros'):
        logging.info("reading from database...")
        count=0
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
#                print row
                count+=1
            conn.close()
        except:
            logging.info("error reading from database...")
        finally:
            return count
        # except sqlite3.OperationalError, err:
       #     logging.exception( "OperationalError: %s" % err)
       #     return 1

    def writedb(self, schema, db_filename='natlink.db', table_name='kb_macros'):
        logging.info("writing to database...")
        count = 0
        try:
            conn = connect(self.wd + db_filename)
            c = conn.cursor()
            # Create table
            c.execute("DROP TABLE %s" % (table_name))
            c.execute("CREATE TABLE %s (%s)" % (table_name, schema))
            for gram, macroobj in self.postDict.iteritems():
                # Insert a row of data
                macro_string=self.customencodechar(macroobj.string)
                print gram, macroobj.string, macro_string
                c.execute("INSERT INTO %s (%s) VALUES ('%s', '%s', '%s')" %
                        (table_name, schema.replace(' text', ''),
                         gram, macro_string, str(macroobj.flags)))
                conn.commit()
                count+=1
        #except Exception, err:
            conn.close()
        except OperationalError, err:
            logging.exception( "OperationalError: %s" % err)
        finally:
            return count

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
