#
# Python Macro Language for Dragon NaturallySpeaking
#
import natlink
import time
from natlinkutils import *
#import msvcrt # built-in module
import win32api


class ThisGrammar(GrammarBase):

    # Todo: embed this list of strings within grammar to save space
    # mapping of keyboard keys to virtual key code to send as key input
    # VK_SPACE,VK_UP,VK_DOWN,VK_LEFT,VK_RIGHT,VK_RETURN,VK_BACK
    kmap = {'space': 0x20, 'up': 0x26, 'down': 0x28, 'left': 0x25, 'right': 0x27,
            'enter': 0x0d, 'backspace': 0x08, 'delete': 0x2e, 'leftclick': 0x201,
            'rightclick': 0x204, 'doubleclick': 0x202}

    # You can also say "Python stop listening" which simulates sleeping
    # by putting the system into a state where the only thing which will
    # be recognized is "Python start listening"
    testGram = """
        <stop> = start presentation mode;
        <notListening> exported = stop presentation mode;
        <navigate> exported = go (left|right);
        <normalState> exported = <stop>;
    """

    # Load the grammar and activate the rule "normalState".  We use
    # activateSet instead of activate because activateSet is an efficient
    # way of saying "deactivateAll" then "activate(xxx)" for every xxx
    # in the array.
    def initialize(self):
        self.load(self.testGram)
        print("Presentation mode is currently disabled")
        self.activateSet(['normalState'],exclusive=0)

    # For the rule "stop", we activate the "notListening" rule which
    # contains only one subrule.  We also force exclusive state for this
    # grammar which turns off all other non-exclusive grammar in the system.
    def gotResults_stop(self,words,fullResults):
        print("Presentation mode enabled")
        self.activateSet(['notListening','navigate'],exclusive=1)

    # When we get "start listening", restore the default state of this
    # grammar.
    def gotResults_notListening(self,words,fullResults):
        print("Presentation mode disabled")
        self.activateSet(['normalState'],exclusive=0)

    def gotResults_navigate(self,words,fullResults):
        key = self.kmap[str(words[1])]
        # this requires wm_key(down|up)
        natlink.playEvents([(wm_keydown, key, 1), (wm_keyup, key, 1)])

#
# Here is the initialization and termination code.  See wordpad.py for more
# comments.
#

thisGrammar = ThisGrammar()
thisGrammar.initialize()

def unload():
    global thisGrammar
    if thisGrammar: thisGrammar.unload()
    thisGrammar = None
