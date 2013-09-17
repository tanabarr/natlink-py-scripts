#
# macro utilities for creating macro-behaviours
#
import logging
import sqlite3

class MacroObj():
    def __init__(self,string='',flags=0):
        self.string=string
        self.flags=flags

class FileStore():
    def __init__(self,defaults_filename='defaults.conf',
                 updates_filename='updates.conf',
                 working_directory='C:\\NatLink\\NatLink\\MacroSystem\\',
                 preDict={},delim="|"):
        try:
            #self.defaults_fd=open(defaults_filename,'r')
            self.updates_filename=updates_filename
            self.postDict=preDict
            self.wd=working_directory
            self.readfile(self.defaults_fd)
            self.delimchar=delim
            #self.writefile()
        except:
            logging.error('could not open default configuration: %s'
                          % defaults_filename)
            return None

    def readfile(self, fd):
        #paramDict={}
        for line in fd.readlines():
            #param, value=line.split(self.delimchar, 1)
            #paramDict.update({param.strip(): value.strip()})
            if (len(line.split(self.delimchar)) == 3):
                gram, macro, flags = line.split(delimchar)
                self.postDict[gram] = MacroObj(macro, flags)

    def writefile(self, output_filename='output.conf'):
        outfile_fd=open(self.wd + output_filename,'w')
        #logging.debug('keys: %s' % self.postDict.keys())
        outfile_fd.write('keyboard macro (name, string, flag) tuples')
        for gram, macroobj in self.postDict.iteritems():
            try:
                outfile_fd.write('\r\n' + str(self.delimchar).join([gram,
                                                   macroobj.string,
                                                   str(macroobj.flags)]))
            except:
                pass


