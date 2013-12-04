#   Windows Explorer macros with intelligent context aware actions.
#       * open folder by (partial) name
#       * invoke open/edit/print verbs on files by (partial) name
#       * copy/move files and folders into subfolders by (partial) name
#
#    (c) 2004--2008 Daniel J. Rocco
#    Licensed under the Creative Commons Attribution-Noncommercial-Share Alike 3.0 United States License
#    http://creativecommons.org/licenses/by-nc-sa/3.0/us/
#
#   Last modified: $Date: 2009-04-23 10:05:04 -0400 (Thu, 23 Apr 2009) $
#
#   01.14.2005 18:36: split "camel case" file and folder names into words
#
#   09.06.2005 21:26: added command for extracting zip files to a single folder
#
#   01.17.2006 21:26: ignore network paths ('\\' prefix) due to
#                   performance/timeout problems
#
#   01.19.2006 23:00: undid 01.17 change: inelegant and unreliable since Windows
#       substitutes the "My Network Places" alias for some network paths.  Instead,
#       can set a limit to the number of folder items, after which voice commands are
#       disabled
#
#   $Id: explorer.py 454 2009-04-23 14:05:04Z drocco $
#
# to do:
#   * doesn't work in "Desktop" folder
#   * problem with open grammar in folders with no subfolders?
#   * selection commands
#       * invert selection
#       * select by file words
#       * select by extension/file type
#   * refine file/folder words
#       * split at numbers, handle @
#   * zip file extraction uses old filename for renamed files
#   * exclusion


# NatLink imports
import natlink
from natlinkutils import *

# Windows imports
import win32com.client, win32api, win32con

# Python Library imports
import string, re


FOLDER_SIZE_LIMIT = 75

class ThisGrammar(GrammarBase):

    gramSpec = """
        <folder> exported  = {folders}
                	;

        <file> exported = (Open | Edit | Print) {files}
                    ;

        <text>  exported = 'text edit' {files}
                    ;

        <copy> exported = (copy | move) to {folders}
                    ;

        <zip> exported = extract zip files | extract and flatten zip files 
                    ;
    """

    # List definitions
    listDefinition= {
        'folders'   :   [],
        'files'     :   [],
    }
    
    def initialize(self):
        if not self.load(self.gramSpec):
            return None
        self.currentModule = ("","",0)
        
    # gotBegin initialization code modeled after Vocola initialization
    def gotBegin(self, moduleInfo):
        # Return if wrong application
        window = matchWindow(moduleInfo,'explorer','')
        if not window:
            return None
        
        # Return if same window and title as before
        if moduleInfo == self.currentModule:
            return None
        self.currentModule = moduleInfo

        self.deactivateAll()
        shellWindow = getShellWindowByHandle(moduleInfo[2])
        if shellWindow != None:
            self.folderName = shellWindow.Document.Folder.Items().Item().Path
            
            self.folders, self.files =getFolderFileNames(shellWindow,FOLDER_SIZE_LIMIT)

            # 03.11.2005 17:18: remove "CVS" from the folder list
            # FIXME: make this modular
            if ".svn" in self.folders:
                del(self.folders[".svn"])
            
            # print self.folders
            self.folderWordList=createFileWordList(self.folders) 
            # print self.folderWordList
            self.fileWordList=createFileWordList(self.files)
            # print self.fileWordList
            self.folderAction= ""
            
            # FIXED: see getFolderNames below (FIXME: Unicode encoded folder names don't work)
            # done (09.02.2004 djr): break names into "words"
            if len (self.folders) >= 1:
                self.listDefinition["folders"] = self.folderWordList.keys()
                self.setList ('folders', self.folderWordList.keys())
                self.activate('folder', window)
                self.activate ("copy", window)
                
            # (09.02.2004 djr): add logic for invoking verbs on files
            if len (self.files) >= 1:
                self.listDefinition["files"] = self.fileWordList.keys()
                self.setList("files", self.listDefinition["files"])
                self.activate("file", window)
                self.activate ("text", window)
                self.activate ("zip", window)


    def gotResults_folder(self,words,fullResults):
        # print "recognized: "
        # print words

        # (11.17.2004 djr): accept folder names after first word
        folder = None
        for word in words:
            if word in self.folderWordList.keys():
                folder = word
                break
    
        if folder != None:
            self.openFolder(folder)

            
    def openFolder(self,folderWords):
        # retrieve the appropriate FolderItem name from the word list
        if folderWords != None and len(folderWords) >0:
            folderItemName = self.folderWordList[folderWords]
            folderItem= self.folders [folderItemName]
            folderItem.InvokeVerb ()


    def gotResults_zip (self, words, fullResults):
        """
            extract zip files to use the name of the file without the zip extension
        """

        flatten = False
        if words [2] == "flatten":
            flatten = True
            
        # extract all the files
        for filename in self.files.keys():
            # skip files without ZIP extension
            if filename [-3:].lower() != "zip":
                continue

            # create the output directory
            path = os.path.dirname(self.files [filename].Path)
            outputPath = os.path.join(path, string.replace(filename, ".zip", ""))

            # extract the files            
            unzip (os.path.join (path,filename), outputPath, flatten)
            
    def gotResults_copy(self , words,fullResults):
        # print words
        
        keys = {
            "copy"  :   "{ctrl+c}",
            "move"  :   "{ctrl+x}",
        }
        action = keys [words [0]]
        natlink.playString(action)
        self.openFolder(words [2])
        natlink.playString("{ctrl+v}")
        
    def gotResults_file(self,words,fullResults):
        # print "recognized: "
        # print words

        # retrieve the appropriate FolderItem name from the word list
        if words [1] in self.fileWordList.keys ():
            folderItemName = self.fileWordList[words [1]]
            # print folderItemName
            
            folderItem= self.files [folderItemName]
            folderItem.InvokeVerb (words [0])            

    def gotResults_text(self,words,fullResults):
        # print "recognized: "
        # print words

        # retrieve the appropriate FolderItem name from the word list
        if words [1] in self.fileWordList.keys ():
            folderItemName = self.fileWordList[words [1]]
            # print folderItemName
            
            folderItem= self.files [folderItemName]
            win32api.ShellExecute(-1, 'Open', r'C:\Program Files\Win32Pad\win32pad.exe', folderItem.Path, None, win32con.SW_SHOWNORMAL)

############################################################################
#
# This is the top-level code
#
# This code gets executed when this file is (re)loaded.
#

# Every grammar file should contains two lines like this for each grammar
# class defined in the file.  These lines causes an instance of the grammar
# to be created and then initialized.

thisGrammar = ThisGrammar()
thisGrammar.initialize()

#
# Every grammar file must define a function called "unload" which will
# call the method "unload" for every grammar which the file loaded.  The
# Python wrapper code will call unload when this file is reloaded.  Calling
# unload first ensures that the classes are cleaned up.
#

def unload():
    global thisGrammar
    if thisGrammar: thisGrammar.unload()
    thisGrammar = None

        
def getShellWindowByHandle(handle):
    """
        Given a handle to a Windows Explorer window, returns the Windows ShellWindow
        object being displayed by that window.  The ShellWindow object represents
        the folder being shown and contains methods and properties for manipulating
        the window and obtaining information about what is being displayed.
        Examples include:
            * retrieving the list of files and folders being displayed
            * reading/updating the selection state
            * programmatically instructing the window to change directories

        references:
            * http://dbforums.com/t867088.html
            * http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/programmersguide/shell_basics/shell_basics_programming/objectmap.asp
    """
    clsid='{9BA05972-F6A8-11CF-A442-00A0C90A8F39}'
    ShellWindows=win32com.client.Dispatch(clsid)
    for i in range(ShellWindows.Count):
        # 11.22.2004 djr: some objects returned by ShellWindows are not
        #   window objects...
        try:
            if ShellWindows[i].HWND == handle:
                return ShellWindows[i]
        except:
            print "bad window? " +str(ShellWindows [i])
            
    return None

def getFolderFileNames(window, limit =-1):
    """
        given a ShellWindow object (a Windows Explorer shell window), return
        two dictionaries containing the folder and filenames in the window
        mapped to the FolderItem Windows objects for those filenames.
    """
    if window == None:
        return
    
    folders = {}
    files = {}

    # print window.LocationURL
    
    # window.document is a ShellFolderView object
    folderItems=window.document.Folder.Items()

    # 01.19.2006 22:58 don't process folders with more than limit items
    if folderItems.Count >= limit and limit >= 0:
        print "directory over size limit (%s, limit=%s): %s" % (folderItems.Count,limit,window.LocationURL)
        return folders, files
    
    for i in range(folderItems.Count):
        folderItem=folderItems.Item (i)

        # Windows Unicode string encoding        
        itemName=folderItem.Name.encode('mbcs')
        if folderItem.IsFolder:
            # put zip files in the files bin (Windows XP lists them as folders)
            if string.lower(itemName[-4:]) == ".zip":
                files [itemName] =folderItem
            else:
                folders[itemName] = folderItem
        else:
            files [itemName] =folderItem

    return folders, files


def generateAllFileWords(filename):
    """
        take a filename and split it into "words" and generate all contiguous subphrases
        of the filename words.  For example:
            generateAllFileWords("CVS") returns ["CVS"]
            generateAllFileWords("Documents and Settings") returns
                ["Documents and Settings", "Documents and", "Documents",
                 "and Settings", "and", "Settings"]

        01.14.2005 18:03 djr: added camel case processing                             
    """
    fileWords=[filename]
    # words = string.split(filename, " ")
    # split the filename on whitespace characters and "camel casing" (e.g. ThisIsACamelCaseWord)
    # only insert the word in the resultant list if it is non-None and nonempty
    words = [word for word in re.split("([A-Z][^A-Z\s]+)|\s+", filename) if word]
    # print words
    length = len (words)

    if length <= 1:
        return fileWords
    
    for i in range(1,length+1):
        for j in range(length):
            word = ""
            for part in [x+ " " for x in words[j:i+j] if x != 'dot']:
                word += part
            word = word [:-1]

            if word not in fileWords and len (word) >= 1:        
                fileWords.append(word)

    # print fileWords    
    return fileWords    


def friendlyName(filename):
    """
        given a string, return a more "voice friendly" version of the string by
        replacing hard to say characters with spelled-out versions and eliminating
        unnecessary characters.

        characters removed: _ + -
        substitutions:
            '.' becomes " dot "


        01.27.2005 10:39 djr: added hyphen removal
    """
        
    replacements = {
        "-": " ",
        "_": " ",
        ".": " dot ",
        "+": " ",
        
    }

    friendly = filename    
    for unfriendly,replacement in replacements.iteritems():
        friendly = string.replace(friendly, unfriendly, replacement)

    return friendly        


def createFileWordList(filenames):
    fileWordList= {}
    duplicates = []

    for filename in filenames:
        # convert the file name to a more "voice friendly" format
        friendly = friendlyName(filename)
        
        for word in generateAllFileWords(friendly):
            if word not in duplicates:
                if word in fileWordList.keys():
                    duplicates.append (word)
                    del fileWordList[word]
                else:
                    fileWordList[word] = filename

    return fileWordList


# fix spaces in spawn calls, from http://mail.python.org/pipermail/python-bugs-list/2001-July/005937.html
def escape(arg):
    import re
    # If arg contains no space or double-quote then
    # no escaping is needed.
    if not re.search(r'[ "]', arg):
        return arg
    # Otherwise the argument must be quoted and all
    # double-quotes, preceding backslashes, and
    # trailing backslashes, must be escaped.
    def repl(match):
        if match.group(2):
            return match.group(1) * 2 + '\"'
        else:
            return match.group(1) * 2
    return '"' + re.sub(r'(\*)("|$)', repl, arg) + '"'


def unzip (filename,outputPath,flatten=False):
    """
        unzip the file into the given directory
    """

    command = "x"
    if flatten:
        command = "e"
        
    # check if directory exists, if not create it
    if not os.path.exists(outputPath):
        os.mkdir(outputPath)
        
    command = r'"C:\Program Files\7-Zip\7z.exe" ' + command + ' -o"' + outputPath + '" "' + filename + '"'
    os.system(escape(command))
    #print command
    