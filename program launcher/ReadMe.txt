*************************************************************************************************Program Launcher Version 1.0
Developed by: P. Meyyappan
EMail: meyyappan_rmp@yahoo.com
FREEWARE: You can very well plugin this technology into your application. I am giving this source for the sake of sharing the knowledge. Please feel free to send your feedback/enhancements/new requirements. 
*************************************************************************************************
About Program Launcher:

	Program Launcher is a simple utility to launch files/folders by pressing a combination of keys, irrespective of whichever application is active. This utility is similar to the one used by multimedia keyboards. All multimedia keyboards makes use of a program which is running in the background. The only difference being that the multimedia keyboards have special keys like volume control which does not need that program to be active.


Technical Details:
	
	I have developed this program using Microsoft Visual Basic 6. I have integrated many concept/technology/logic into an useful utility. Some raw materials have been processed into an useful product.

	This program makes use of several API functions and other techniques posted in VB source code websites like allapi.net, vbaccelerator.com, vbcode.com, freevbcode.com, etc. The backbone of this program is the RegisterHotKey API function.

	Once the program is started, a keyboard icon will appear next to the system clock. Right click on the the icon to display a menu. The details of the application/file, path, keystrokes are written to a binary file called data.dat. Press Win+A to show this program.

Functionality:
	
	Assume you want to open Notepad when you press Win+Shift+N. Follow these steps:
1. Register the keys with RegisterHotKey funtion be passing an unique hotkey id and the keystrokes.
2. In an infinite loop listen to Windows messages. When an hotkey message is found, check whether that hotkey was registered by your program using handle to the window. If so, open Notepad.
3. To remove the hotkey any time, call UnregisterHotKey by passing the hotkeyid.

What you will learn:
	
	From this project, you will come to know how to implement the following through VB6:
1. How to display form on top of all other windows.
2. How to add an icon to System Tray.
3. How to open files of any type with the associated application.
4. How to remove the title bar during run-time.
5. How to create binary files using Open method; read record by record from the file
6. How to use structures in VB
7. How to create shortcut files(.lnk)
8. How to create gradient color effect in a form
8. How to restrict the number of instances of the application running. (Another instance of Registry Editor, Yahoo Messenger or Winamp cannot be opened). Though this can be implemented effortlessly using "App.PreviousInstance" property, still another instance can be opened using a seperate copy of the exe file. In this project, you wil find how to check whether another instance is already running by using FindWindow API funcion.

Last but not the least, about the cool interface of Program Launcher.

You can include this utility in your projects and make it more useable for your clients. Please do send in your feedback/enhancement/new requirements.

-Meyyappan