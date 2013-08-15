#
# This file is part of Dragonfly.
# (c) Copyright 2007, 2008 by Christo Butcher
# Licensed under the LGPL.
#
#   Dragonfly is free software: you can redistribute it and/or modify it
#   under the terms of the GNU Lesser General Public License as published
#   by the Free Software Foundation, either version 3 of the License, or
#   (at your option) any later version.
#
#   Dragonfly is distributed in the hope that it will be useful, but
#   WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
#   Lesser General Public License for more details.
#
#   You should have received a copy of the GNU Lesser General Public
#   License along with Dragonfly.  If not, see
#   <http://www.gnu.org/licenses/>.
#

"""
    This module is a simple example of Dragonfly use.

    It shows how to use Dragonfly's Grammar, AppContext, and MappingRule
    classes.  This module can be activated in the same way as other
    Natlink macros by placing it in the "My Documents\Natlink folder" or
    "Program Files\NetLink/MacroSystem".

"""

from dragonfly import (Grammar, AppContext, MappingRule, Dictation,
                       Key, Text)


#---------------------------------------------------------------------------
# Create this module's grammar and the context under which it'll be active.

grammar_context = AppContext(executable="chrome")
grammar = Grammar("chrome_example", context=grammar_context)


#---------------------------------------------------------------------------
# Create a mapping rule which maps things you can say to actions.
#
# Note the relationship between the *mapping* and *extras* keyword
#  arguments.  The extras is a list of Dragonfly elements which are
#  available to be used in the specs of the mapping.  In this example
#  the Dictation("text")* extra makes it possible to use "<text>"
#  within a mapping spec and "%(text)s" within the associated action.

example_rule = MappingRule(
    name="example",    # The name of the rule.
    mapping={          # The mapping dict: spec -> action.
                # Google Chrome commands
                "zoom in":            Key("c-plus"),
                "zoom out":           Key("c-minus"),
                'new': Key("c-t"),
                'new window': Key("c-n"),
                'next': Key("c-tab"),
                'previous': Key("cs-tab"),
                'private': Key("a-e, i/20"),
                'close': Key('c-w'),
                'bookmark': Key('c-d/20'),
                'tools': Key('a-e'),
#             "save [file]":            Key("c-s"),
#             "save [file] as":         Key("a-f, a"),
#             "save [file] as <text>":  Key("a-f, a/20") + Text("%(text)s"),
#             "find <text>":            Key("c-f/20") + Text("%(text)s\n"),
                # c-style programming abbreviations
                'vim begin comment': Key('i/* '),
                'vim end comment': Key('i */{enter}'),
                'vim begin long comment': Key('i#{esc}ib{space}'),
                'vim end long comment': Key('i#{esc}ie{enter}'),
                'vim line comment': Key('i#{esc}il{enter}'),
                'vim define': Key('i#{esc}id{space}'),
                'vim include': Key('i#{esc}ii{space}'),
                'vim equals': Key('i{right} = '),
                # vim commands
                'vim format': Key('Q'),
				'vim undo': Key('u'),
                'vim redo': Key('{ctrl+r}'),
				'vim next': Key(':bn'),
                'vim remove buffer': Key(':bd'),
                'vim list buffers': Key(':buffers{enter}:b'),
                'vim previous buffer': Key(':b#'),
                'vim start macro': Key('qz'),
                'vim repeat macro': Key('@z'),
                'vim previous': Key(':bp'),
				'vim save': Key(':w'),
                'vim close': Key(':q'),
				'vim taglist': Key('{ctrl+p}'),
                'vim update': Key(':!ctags -a .'),
                'vim list changes': Key(':changes'),
                'vim previous change': Key('g;'),
                'vim next change': Key('g,'),
                'vim return': ("''",0x00),
                'vim matching': Key('%'),
                'vim undo jump': Key('``'),
                'vim insert space': Key('i{space}{esc}'),
                'vim hash': Key('i#{esc}'),
                'vim insert blank line next': Key('o{up}'),
                'vim insert blank line previous': Key('O{down}'),
                'vim previous command': (':{up}',0xff),
                'vim copy previous line': Key(':-1y'),
                'vim copy next line': Key(':+1y'),
                'vim copy current line': Key('yy'),
                'vim remove previous line': Key(':-1d'),
                'vim remove next line': Key(':+1d'),
                'vim set mark': Key('mz'),
                'vim goto mark': ("'zi",0),
                'vim scroll to top': Key('zt'),
                'vim scroll to bottom': Key('zb'),
                'vim edit another': Key(':e '),
                'vim file browser': Key(':e.'),
                'vim folds': Key('{ctrl+f}'),
                'vim window up': Key('{ctrl+k}'),
                'vim window down': Key('{ctrl+j}'),
                'vim window left': Key('{ctrl+h}'),
                'vim window right': Key('{ctrl+l}'),
                'vim split vertical': Key(':vsp'),
                'vim replace': Key('R'),
                'vim make': Key(':make'),
                'vim next error': Key(':cn'),
                'vim previous error': Key(':cp'),
                'vim list errors': Key(':clist'),
                'vim match bracket': Key('%'),
                'vim change character case': Key('~'),
                'vim beginning previous': Key('-'),
                'vim beginning next': Key('+'),
                #screen commands
                'attach screen ': Key('screen -R{enter}'),
                'attach screen existing': Key('screen -x{enter}'),
                'screen scrollback mode': Key('['),
                'screen scrollback paste': (']',0x00),
                'screen previous': Key('p'),
				'screen next': Key('n'),
                'screen help': Key('?'),
				'screen new': Key('c'),
                'screen detach': Key('d'),
				'screen list': Key('"'),
                'screen kill': Key('k'),
				'screen title': Key('A'),
                # window split related
                'screen switch': Key('{tab}'),
				'screen split': Key('S'),
                'screen vertical': Key('|'),
				'screen crop': Key('Q'),
                'screen remove': Key('X'),
            },
    extras=[           # Special elements in the specs of the mapping.
            Dictation("text"),
           ],
    )
    #convert format:
    #s/(\('.*'\).*/Key(\1),/gc

# Add the action rule to the grammarecurjinstance.
grammar.add_rule(example_rule)


#---------------------------------------------------------------------------
# Load the grammar instance and define how to unload it.

grammar.load()

# Unload function which will be called by natlink at unload time.
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
