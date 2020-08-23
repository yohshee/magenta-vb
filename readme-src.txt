====================================
Readme for Magenta Source Code (VB6)

Version: 0.6 Build 47

Public Release on:
2/22/2003

Readme updated:
8/21/2020
====================================

Preface
=======
As opposed to the C++ version of Magenta being an ancient artifact, this is probably
a prehistoric relic. A lot of code got written in VB6 back in the day, and this is just one 
drop in that gigantic, greatly disliked bucket. Again, much like the C++ version of Magenta, 
this is released just for fun, so that you might be able to see what things were like way back when.

It's almost hilarious how much less complicated it is compared to the C++ version.

As a final note, the rest of this file is mostly the same as it was when I first wrote it in 2003.

- rac, 8/21/2020.

Introduction
============

This code that you've just unwrapped is probably not the best I've written in some 
regards. Sure, it does nifty stuff, but some of that nifty stuff is embedded in 
some very nice ugliness; but ugliness is sometimes needed to do such niftiness...
and I've just repeated myself.

Anyway, the structure of the VB code is hopefully quite apparent, but here's a 
brief outline of the main files (the *.frx files, of course, go with their 
respective *.frm files)

res/              - Contains copies of all graphics used in Magenta
modGlobals.bas    - Holds the Sub Main and crash-handling routines
ParserModule.bas  - Contains some nifty string manipulation routines
OnlinePerson.cls  - Class module abstracting out an online user
frmAbout.frm      - About box
frmChat.frm       - Main Magenta window
frmConsole.frm    - Chat console
frmCrash.frm      - Crash report window
frmIgnored.frm    - List of ignored users
frmMain.frm       - Chat status window (filename is a relic from the 
                    original Magenta)
frmOptions.frm    - Options dialog box
frmOutput.frm     - Debug window that is not currently used.
frmPrivate.frm    - Private chat window
frmProperties.frm - Dialog pulled up when a user in the status window is
                    double-clicked.

Magenta also depends on the following DLLs and ActiveX controls (all of
which come with VB6 Professional):

Microsoft Dialog Automation Objects (DLGOBJS.DLL, the license file should be on your
VB6 CD)
Microsoft Windows Common Controls 6.0 (SP4)
Microsoft Common Dialog Control 6.0
Microsoft Rich Textbox Control 6.0
Microsoft Winsock Control 6.0 (SP4)

(It's imperative that you use these versions, because some Magenta magic depends on
these updated libraries)

Rick "Yohshee" Coogle
mailto:yohshee@hotmail.com
http://www.fadedtwilight.org

Copyright(C) 2002-2003. 