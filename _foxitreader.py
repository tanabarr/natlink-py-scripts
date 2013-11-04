#
# This file is a command-module for Dragonfly.
# (c) Copyright 2008 by Christo Butcher
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>
#

"""
Command-module for **Foxit Reader** PDF viewer
============================================================================

This module offers commands to control `Foxit Reader
<http://www.foxitsoftware.com/pdf/rd_intro.php>`_, a free
and lightweight PDF reader.

Installation
----------------------------------------------------------------------------

If you are using DNS and Natlink, simply place this file in you Natlink
macros directory.  It will then be automatically loaded by Natlink when
you next toggle your microphone or restart Natlink.

"""

try:
    import pkg_resources
    pkg_resources.require("dragonfly >= 0.6.5beta1.dev-r76")
except ImportError:
    pass

from dragonfly import (Grammar, AppContext, MappingRule, Dictation,
                       Key, Text, Config, Section, Item, IntegerRef)


#---------------------------------------------------------------------------
# Set up this module's configuration.

config                             = Config("Foxit reader control")
config.lang                        = Section("Language section")
config.lang.new_win                = Item("new (window | win)")
#config.generate_config_file()
config.load()


#---------------------------------------------------------------------------
# Create the main command rule.

#class CommandRule(MappingRule):
class StaticRule(MappingRule):

    mapping  = {
                "zoom in [<n>]":            Key("c-equals:%(n)d"),
                "zoom out [<n>]":           Key("c-hyphen:%(n)d"),
                "zoom [one] hundred":       Key("c-1"),
                "zoom [whole | full] page": Key("c-2"),
                "zoom [page] width":        Key("c-3"),

                "find <text>":              Key("c-f") + Text("%(text)s")
           	                                 + Key("f3"),
                "find next":                Key("f3"),

                "[go to] page <int>":       Key("c-g") + Text("%(int)d\n"),

                "print file":               Key("c-p"),
                "print setup":              Key("a-f, r"),
                "reading mode":             Key("c-H"),
                "previous page":            Key("c-pgup"),
                # awkward method of switching between bookmark view and reading
                # pane. creates new bookmark, delete it,then can navigate to
                # bookmark.when wanting to return to Reading window, create and/or
                # focuses start tab. Then switch to tab next (assume is only one tab
                # open)
                "navigation window":        Key("c-b, enter, del"),
                "reading window":           Key("ca-s/20:2/20"),
                "close navigation window":  Key("f4"),
                'next':                     Key("c-tab"),
                'previous':                 Key("cs-tab"),
                'close':                    Key('c-w'),
               }
    extras   = [
                IntegerRef("n", 1, 10),
                IntegerRef("int", 1, 10000),
                Dictation("text"),
               ]
    defaults = {
                "n": 1,
               }


#---------------------------------------------------------------------------
# Create and load this module's grammar.

#context = AppContext(executable="Foxit Reader")
#grammar = Grammar("foxit reader", context=context)
grammar = Grammar("foxit reader")
#grammar.add_rule(CommandRule())
grammar.add_rule(StaticRule())
grammar.load()

# Unload function which will be called by natlink at unload time.
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
