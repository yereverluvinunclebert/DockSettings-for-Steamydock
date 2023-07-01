# DockSettings for Steamydock

DOCK SETTINGS for Steamydock, written in VB6. A WoW64 dock settings
utility for Reactos, XP, Win7, 8 and 10+.

This utility controls the settings of the dock and where the user makes 
configuration changes for the dock itself. The utility is a functional 
reproduction of the original Rocketdock dock settings screen with some 
enhancements. The look and feel of the GUI is limited to emulating and 
enhancing what Rocketdock provided. The idea is that this will make the utility 
quite familiar to Rocketdock users. It operates with Steamydock, my open source 
replacement for Rocketdock.

![themes](https://github.com/yereverluvinunclebert/dockSettings/assets/2788342/f181ab5e-2838-4548-bf1f-55d75c04f4ca)


NOTE: The dock settings tool is Beta-grade software, under development, not yet 
ready to use on a production system - use at your own risk.

NOTE: This tool and the build instructions are being overhauled, do not expect 
it to load without flaws until this message is removed.

![dockS-aboutPane](https://github.com/yereverluvinunclebert/dockSettings/assets/2788342/b0da7f63-3802-47ee-9a70-a44e41444d59)

BUILD: The program runs without any additional Microsoft plugins.

Built using: VB6, MZ-TOOLS 3.0, VBAdvance, CodeHelp Core IDE Extender
Framework 2.2 & Rubberduck 2.4.1

Links:

	MZ-TOOLS https://www.mztools.com/  
	CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1  
	Rubberduck http://rubberduckvba.com/  
	VBAdvance  https://classicvb.net/tools/vbAdvance/
	
	Rocketdock https://punklabs.com/  
	Registry code ALLAPI.COM  
	La Volpe http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1  
	PrivateExtractIcons code http://www.activevb.de/rubriken/  
	Persistent debug code http://www.vbforums.com/member.php?234143-Elroy  
	Elroy on the VBForums for the balloon tooltips
	Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain  


Tested on :

	ReactOS 0.4.14 32bit on virtualBox  
	Windows 7 Professional 32bit on Intel  
	Windows 7 Ultimate 64bit on Intel  
	Windows 7 Professional 64bit on Intel  
	Windows XP SP3 32bit on Intel  
	Windows 10 Home 64bit on Intel  
	Windows 10 Home 64bit on AMD  
	Windows 11 64bit on Intel

Dependencies:

o A windows-alike o/s such as Windows XP, 7-11 or ReactOS.

o Microsoft VB6 IDE installed with its runtime components. The program runs 
without any additional Microsoft OCX components, just the basic controls that 
ship with VB6.

o Requires the SteamyDock program source code to be downloaded and available in 
an adjacent folder as some of the BAS modules are common and shared.

Example folder structure:
	
	E:\VB6\steamydock   ! https://github.com/yereverluvinunclebert/SteamyDock
	E:\VB6\docksettings ! this repo.
	E:\VB6\rocketdock   ! from https://github.com/yereverluvinunclebert/rocketdock

o Krools replacement for the Microsoft Windows Common Controls found in
mscomctl.ocx (treeview, slider) are replicated by the addition of one
dedicated OCX file that is shipped with this package.

o  CCRSlider.ocx

This OCX will reside in the program folder. The program reference to this OCX is 
contained within the supplied resource file Panzer Earth Gauge.RES. It is 
compiled into the binary.

o In the VB6 IDE - project - references - browse - select the OLEEXP.tlb

Project References:
VisualBasic for Applications  
VisualBasic Runtime Objects and Procedures  
VisualBasic Objects and Procedures  

Credits:

I have really tried to maintain the credits as the project has progressed. If I
have made a mistake and left someone out then do forgive me. I will make amends
if anyone points out my mistake in leaving someone out.

MicroSoft in the 90s - MS built good, lean and useful tools in the late 90s and
early 2000s. Thanks for VB6.

Elroys code to add balloon tips to comboBox
https://www.vbforums.com/showthread.php?893844-VB6-QUESTION-How-to-capture-the-MouseOver-Event-on-a-comboBox

Shuja Ali @ codeguru for his settings.ini code.

An unknown, untraceable source, possibly on MSN - for the KillApp code

ALLAPI.COM For the registry reading code.

Elroy on VB forums for his Persistent debug window 
http://www.vbforums.com/member.php?234143-Elroy

Rxbagain on codeguru for his Open File common dialog code without a dependent
OCX http://forums.codeguru.com/member.php?92278-rxbagain

si_the_geek for his special folder code



LICENCE AGREEMENTS:

Copyright 2023 Dean Beedell

In addition to the GNU General Public Licence please be aware that you may use
any of my own imagery in your own creations but commercially only with my
permission. In all other non-commercial cases I require a credit to the
original artist using my name or one of my pseudonyms and a link to my site.
With regard to the commercial use of incorporated images, permission and a
licence would need to be obtained from the original owner and creator, ie. me.

![Uploading dockS-fullscreen.png…]()
