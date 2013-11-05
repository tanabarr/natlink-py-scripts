#
# Python Macro Language for Dragon NaturallySpeaking
# nextApp
#

import natlink
from natlinkutils import *

class ThisGrammar(GrammarBase):

    gramSpec = """
        <start> exported = show menu;
    """

    def gotResults_start(self,words,fullResults):
        # execute a control-left drag down 30 pixels
        #x,y = natlink.getCursorPos()
        natlink.playString('{F10}',0x01) 
    '''
        natlink.playEvents( [ (wm_syskeydown,0x12,1),
                              (wm_keydown,0x09,1),
                              (wm_keyup,0x09,1),#(wm_lbuttondown,x,y),
                              (wm_keydown,0x09,1),
                              (wm_keyup,0x09,1),#(wm_lbuttondown,x,y),
                              #(wm_mousemove,x,y+30),
                              #(wm_lbuttonup,x,y+30),
                               (wm_syskeyup,0x12,1)
                            ] )
    '''

    def initialize(self):
        self.load(self.gramSpec)
        self.activateAll()

thisGrammar = ThisGrammar()
thisGrammar.initialize()

def unload():
    global thisGrammar
    if thisGrammar: thisGrammar.unload()
    thisGrammar = None
