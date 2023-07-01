This utility is a functional reproduction of Rocketdock's original settings screen. The design is limited to enhancing what Rocketdock already provides. This will make the utility familiar to Rocketdock users.

Unfortunately Rocketdock's settings screen has a few annoying bugs or limitations. One of the bugs is the time the extended time Rocketdock takes to respond to a right-click on an icon. This is vital functionality as it is precisely how you change the appearance or the functionality of any Rocketdock icon. Rocketdock has to read its entire stored library of .ICO or.PNG files so it can display a graphical selection of thumbnail icons for the user to choose. This does not affect standard operation but when your custom icon collection consists of hundreds of image files, the larger the default icon folder the longer it takes. As a result a right click can take 20-30 secs on a typical core2duo with a 2.5ghz CPU but even on newer, faster systems there is a many-second delay while each icon is read and stored in the 'cache'. There are library folders of over 3,000 icons, so you can imagine, with a folder such as these this would prove a serious bug that affects Rocketdock's usability. I set out to re-write Rocketdock's settings screen and resolve this single bug but in so doing I enhanced what it provides in general.

This utility improves upon the original in certain areas:

* You can flip to the next Rocketdock item without having to leave the settings screen completely as before.
* It indicates by number which Rocketdock item is currently selected.
* The user can delete unwanted icons directly from the file thumbnail display.
* The icon preview can be resized so the user can see how the icon will look in the dock.
* There are tooltips for all controls (before there were none).
* There is now a readily available help facility.
* The images in the thumbnail view are now more visible at 64x64 rather than 32x32 as they were before.
* The user can flip between file list and thumbnail view as it suits them.
* The new "get more" button is not a dead link and instead takes you to a useful location where there are a lot more icons for the user to download. 
* There is a working icon type filter allowing you to select one type of icon. The old one was non-functional.
* The code is open source so that a user can change the utility themselves.
* The user can refresh the file list at any time if there have been any changes to a folder.
* The utility saves copies of the settings.ini file so that you can always revert your dock back to an earlier state.
* There are many more icon options for automatically creating icon entries.
* It provides a steampunk library of various unique icons.
* The dock is still fully functional while this utility is operating.
* It runs many times faster on the critical icon thumbnail view, taking less than one second rather than 20+secs.


This is the first VB6 project that I have 'undertaken and completed' so forgive the errors in coding
styles and methods. Entirely self-taught and a mere hobbyist in VB6.
'
The reason I created it was to teach myself VB6, to get back into the 'groove'.
Back in the 90s I was programming in QB45 and VB DOS and then VB6 but left VB6 and abandoned my
main project when VB6 was deprecated. My skills were paltry then and were picked up from the days
of Sinclair Zx80s. My aim now is to resurrect such skills that I had and improve upon them.
A secondary aim is to teach myself how to code in technologies that I have encountered. When this
project is complete my next aim is to migrate it to VB.NET through the versions to find out what
problems are typically encountered in a project such as this.
'
Starting with VB6, it was a big surprise to me to find such inadequate native image type handling,
VB6 being unable to handle the various image types without the usage of a great deal of code and
API calls. I learnt that VB6 can 'do' anything but it can also be hard work to make it do so. I
could not have made this utility without the help of code from the various projects I have listed below.
'
I hope you enjoy the functionality this utility provides. If you think you can improve anything then please
feel free to do so. If you dislike my programming style then do keep those thoughts to yourself. :)
'
Built on a 2.5ghz core2duo Dell Latitude E5400 running Windows 7 Ultimate 64bit using VB6 SP6.
'
'   Tested on :
Windows 7 Pro 32bit on Intel
Windows 7 Ultimate 64bit on Intel
Windows XP SP3 on Intel
Windows 10 Home 64bit on AMD and Intel
'
' Dependencies:
Microsoft Windows Common Controls 6.0 SP3   mscomctl.ocx (treeview, slider)
RocketDock 1.3.5
'
' Notes:
Integers are retained (rather than longs) as some of these are passed to
library API functions in code that is not my own so I am loathe to change.
A lot of the code provided (by better devs than me) seems to have code quality
issues - I haven't gone through all their code to fix every problem but I have fixed lots...
'
The icons are displayed using Lavolpe's transparent DIB image code,
except for the .ico files which use his earlier StdPictureEx class.
The original ico code caused many strange visual artifacts and complete failures to show .ico files.
especially when other image types were displayed on screen simultaneously.
'
' Summary:
The program reads a default icon folder from Rocketdock's settings.ini or registry.
It reads the contents of the folder and sub-folders into a treeview and displays the first 12 of the
icons using 12 dynamically created picboxes. The icons are displayed using Lavolpe's
transparent DIB image code, except for the .ico files which use the earlier StdPictureEx class.
DLLs and EXEs with embedded icons are handled using an undocumented API named PrivateExtractIcons.
One selected image is extracted and displayed in larger size using the above code in the preview window.
'
A copy of Rocketdock's settings are transferred from the registry or settings.ini into an interim
settings file which provides a common method of handling the data.
The icon details are read from this file and the details
of the selected icon are displayed in the text boxes in the 'properties' frame. This data is also
read when the user chooses to the display the Rocketdock map.
'
In that 'map' each dock image is displayed in smaller form in dynamically created picboxes.
The RD map acts a cache of images that takes a few seconds to create but
doing it this way means there is no subsequent delay when viewing any other part of the map.
The images on the map can then be scrolled into visibility viewing fifteen icons at a time. It has
been tested with a map containing up to 67 icons.
'
The icon details are written to the registry or the settings file but only after Rocketdock
has been closed and just before it is restarted otherwise it will overwrite any settings
changes when it exits.
'
The utility itself has some configuration details that it stores in its own local settings.ini file.
'
Rather than create a manifest and bundle the OCX within a .res file for extraction on the utilities'
first run, I have created instead my own installer program that attempts to place the required OCX
file(s) in the correct location.
'
The reason that an SxS configuration is absent is that errors were generated as soon as comctl32 was
added to the manifest. I became so bored always trying to fix the many manifest errors that I gave up...
A finicky and therefore useless method to packaging up a VB6 app.
'
The font selection and file/folder dialogs are generated using Win32 APIs rather than the
common dialog OCX which dispensed with another OCX.
'
I made an attempt to replace the mscomctl.ocx with an in-built treeview replacement using
Win32 APIs but that was a fair bit of work so that task remains unfinished. I have that version
put aside and may complete it later. This will free the program of all external dependencies.
'
' Missing:
The only component not yet functional is the 'generate dock' button. At the moment
it only tests the registry for certain entries in the uninstall section of the registry.
Eventually, it will generate a settings.ini file containing all the useful software you have in your
system.
'
' Licence:
Copyright � 2019 Dean Beedell
'
This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public Licence as published by the Free Software Foundation; either version 2 of the
License, or (at your option) any later version.
'
This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
General Public Licence for more details.
'
You should have received a copy of the GNU General Public Licence along with this program; if not,
write to the Free Software Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301
USA
'
If you use this software in any way whatsoever then that implies acceptance of the licence. If you
do not wish to comply with the licence terms then please remove the download, binary and source code
from your systems immediately.

Credits:

 Credits : LA Volpe (VB Forums) for his transparent picture handling.

*Shuja Ali (codeguru.com) for his settings.ini code.
*KillApp code from an unknown, untraceable source, possibly on MSN.
*Registry reading code from ALLAPI.COM.
*Punklabs for the original inspiration and for Rocketdock, Skunkie in particular.
*Active VB Germany for information on the undocumented PrivateExtractIcons API.
*Elroy on VB forums for his Persistent debug window
*Rxbagain on codeguru for his Open File common dialog code without dependent OCX

Built using: VB6, MZ-TOOLS 3.0, CodeHelp Core IDE Extender Framework 2.2 & Rubberduck 2.4.1

*MZ-TOOLS https://www.mztools.com/
*CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1
*Rubberduck http://rubberduckvba.com/
*Rocketdock https://punklabs.com/
*Registry code ALLAPI.COM
*http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1
*PrivateExtractIcons code http://www.activevb.de/rubriken/
*Persistent debug code http://www.vbforums.com/member.php?234143-Elroy
*Open File common dialog code without dependent OCX-    
                http://forums.codeguru.com/member.php?92278-rxbagain
*Open font dialog code without dependent OCX



USAGE & HELP:

Note: If you hover your mouse cursor on the various components that comprise the utility a tooltip will appear that will give more information on each item. There is a help button on the bottom right that will provide further detail at any time.

Folders

At the top left you will see a list of all the folders you currently have available to you. This display is called a treeview. The top folder is named 'icons'.
That folder is typically located at: C:\Program Files\RocketDock\Icons (default location).

The folder beneath that is named "My Collection" and it contains the Steampunk icons that are packaged with this tool. You can select these folders by clicking on any of  these folders and the icons within each will be displayed.

Beneath this location lies any custom folders of your own that you wish to add. Initially there will be none but you can add any icons to this folder and have them available to select as you choose. The - and + buttons allow you to add and remove your own icon folders. They will be remembered when the tool is next re-opened.

Two small tick boxes indicate how Rocketdock is currently saving its settings, either to a file or the registry, this is for information only.

Icons

This pane (on the top right) will show you a preview of any icons available in the folder you have selected in the treeview. The drop down menu below the Icons list allows you to select icon types (gif, png, ico, bmp &c). A small red 'x' allows you to select a specific icon file for deletion. You can select one of three view types, a file list, a large thumbnail view or a small thumbnail view. A right click on the thumbnails gives you the alternative choice. A click on the small filelist button (top right) selects the file list view. There is a refresh button to the right that will cause the icon list to be re-read from the folder you have selected.

A single-click on any icon in the icon pane will show the icon in larger size in the preview pane below.

A double-click on any icon in the icon pane will select the chosen icon for insertion into the icon map. The + button at the top right does the same. The preview will update and only a 'save' is required to update the map. You can use left and right keys to navigate the dock or the icon slider as you wish.

Icon Map
The icon map is hidden when the tool first starts. A small down button on the right hand side will cause the icon map to appear. It is kept hidden so that the overall look of the tool matches the appearance of the Rocketdock icon settings screen. A single-click on any icon in the icon map will show the icon in larger size in the preview pane below. The icons in the map relate to the icons as shown in the Rocketdock. They appear in the same order and will have the same appearance. The icons are numbered from one upward. The dock can contain as many as seventy icons or more depending upon how much you intend to use Rocketdock. A right click on the map gives you more choices, the option to add or delete an icon as well as the ability to re-order the icons as you see fit. There is a refresh button to the right that will cause the map to be re-read from Rocketdock's own settings.

You can use left and right keys to navigate the dock. Other controls consist of a slider, two large navigation buttons and an 'up' button to hide the dock. Note the utility starts much more quickly when the dock is hidden.  

Preview

This pane allows you to see which icon you have currently selected to view. These are selected from the icon map in the middle or the icon pane at the top right. The size of the displayed icon can be modified using the slider at the bottom. There are also two slim buttons on the left and right which allow you to select the next or previous icons, those subsequently displayed are the icons on the map. The images size is displayed where it is appropriate to do so.

Properties

Here is where you change the item title, the target and other special actions that are available. There is a large number on the right hand side, that corresponds to the location of the icon in the icon map. As you click the right or left button on the preview pane that number will change accordingly. The icon indicated in the map will also change. 

* Name: Set the label that will appear above your icon when your mouse cursor is hovering over the dock.
* Current Icon: When you have selected an icon from the icon pane the full path of the icon will appear here.
* Target: Set the target location of the item on your computer, this can be a file, folder, URL or program.

Next to the target property field is a button that when pressed, will disclose a file selection dialog allowing you to choose a target file, program or image. A right click on this button will provide a number of alternative target options such as folder, network &c.

* Start In: This sets the working directory for the target application if the target program requires a default folder to operate within.

Next to the 'start in' property field is a button, that when pressed, will disclose a folder selection dialog allowing you to choose a target folder.

* Arguments: Sets optional parameters for the target application.
* Run: This sets the minimised/maximised state of the window when the item is launched from RocketDock.
* Open Running: This drop down menu allows you to override the "Open Running Application Instance" on a per icon basis. You get the choice of: "Use Global Setting," "Always," and "Never."
* Popup Menu: This enables additional actions to be displayed in the RocketDock context menu for the specific icon.

As you make changes to the above property fields ensure that you click "save" or your changes will be lost as you swwitch to the next icon. Any icon changes will then appear in the icon map. An icon will not appear in the map until save is pressed.

None of your changes will appear in Rocketdock itself until you press "save & restart". The reason for this is that Rocketdock does not read its settings except on startup. A quick restart causes your new icons to appear in the dock straight away.

The backup button causes a backup of the settings.ini file to take place, it also takes you to the backup folder where the settings backups are stored. To restore your icon set up simply configure Rocketdock into a portable settings.ini mode and copy the bkpSettings.ini file that you find there to replace the settings.ini in Rocketdock's own folder. and re-run the utility. This tool will find the icons and you may then "save & restart" Rocketdock and the icons will re-appear in Rocketdock too.

A check box toggles the information dialog on/off. When it is selected confirmation messages will be given before any radical change takes place. Turn it on or off as you require.

Menus
A right click here and there will bring up other menu options.

A right click on the map shows the map menu.

The target menu is disclosed when right-clicking on the target button.

The main menu is disclosed when right-clicking on everywhere else.

Font Selection from the main right-click menu.

Other options include a theme change (only partially implemented), changing Rocketdock's installation folder and a debug option in case the program throws an error. The other menu options provide information and social-media URLs.

The utility is created using VB6, Microsoft's once vaunted flagship language. It uses VB6 to prove it can be done and to reacquaint myself with the technology. The utility will be migrated to .NET so please watch this space.


