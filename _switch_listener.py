#
# Python Macro Language for Dragon NaturallySpeaking
#
#TODO: fix or replace kbfunc
#A function can be used to compare monitors, only performed start-up
#polling of sleep state if in the car identified by the below monitor
#wmi can be used to detect device changes, when USB audio detected, run voice
#recognition restart script (nsrst.bat)
#output of car screen
#Found monitor Monitor(-468515699) with geometry Rectangle(0.0, 0.0, 1024.0,
#                                                          768.0).
#output of laptop only
#Found monitor Monitor(-468515699) with geometry Rectangle(0.0, 0.0, 1366.0,
#                                                          768.0).

import natlink
import time
from natlinkutils import *
#import msvcrt # built-in module
import win32api

#def kbfunc():
#    """ Windows specific keyboard polling mechanism (nonblocking) """
#    return ord(msvcrt.getch()) if msvcrt.kbhit() else 0

class ThisGrammar(GrammarBase):

    # You can also say "Python stop listening" which simulates sleeping
    # by putting the system into a state where the only thing which will
    # be recognized is "Python start listening"

    testGram = """
        <stop> = switch active listener | miniature stop;
        <notListening> exported = switch active listener;
        <normalState> exported = <stop>;
    """

    # Load the grammar and activate the rule "normalState".  We use
    # activateSet instead of activate because activateSet is an efficient
    # way of saying "deactivateAll" then "activate(xxx)" for every xxx
    # in the array.

    def initialize(self):
        self.load(self.testGram)
        # problem of Dragon not waiting on initialisation when in sleep,
        # exacerbated by background noise.
        # Seems that cannot change micState until it is waiting in sleeping
        # mode. Waiting in initialisation in natlink delays  Dragon start-up.
        # polling produces a ~68 seconds wait for sleeping mode to activate.
        # TODO: to avoid this delay, need to catch signal from DNS post NatLink
        # init. ( temporary solution to test the dimensions of monitor)
        monitorRects=[mon[2] for mon in win32api.EnumDisplayMonitors(None, None) if
                      mon[2] == (0,0,1024,768)]
        if not monitorRects:
            print("not car monitor")
            self.activateSet(['normalState'],exclusive=0)
        else:
            STEP=4
            count=0
            micstate=natlink.getMicState()
            print "micstate: {0}, {1} sec wait".format(micstate, count)

            # poll until Dragon is in sleep state (select "start dragon in sleep
            # state" option in Dragon options)
            while micstate not in ['sleeping', 'on']: #'off':
                time.sleep(STEP)
                count+=STEP
                micstate=natlink.getMicState()
                print "polling micstate: {0}, {1} sec wait".format(micstate, count)

            natlink.setMicState('on')
            micstate=natlink.getMicState()
            print "micstate: {0}, {1} sec wait".format(micstate, count)

            # now Dragon is on, put in an exclusive state waiting for "switch
            # active listener" grammar
            self.activateSet(['notListening'],exclusive=1)

    # For the rule "stop", we activate the "notListening" rule which
    # contains only one subrule.  We also force exclusive state for this
    # grammar which turns off all other non-exclusive grammar in the system.

    def gotResults_stop(self,words,fullResults):
        self.activateSet(['notListening'],exclusive=1)

    # When we get "start listening", restore the default state of this
    # grammar.

    def gotResults_notListening(self,words,fullResults):
        self.activateSet(['normalState'],exclusive=0)
        natlink.recognitionMimic(["switch", "to","spell", "mode"])

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
