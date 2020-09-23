VB Code Auto Indenter Add-In - ReadMe File

Author:	Andrew Davey
Email:	andrewdavey@hotmail.com
Web:	www.starsoftsoftware.co.uk

IF YOU DON'T READ THIS AND SOMETHING DOESN'T WORK
IT'S YOUR OWN DAMN FAULT!

Hi there fellow VB user! 
How often do you download VB code from the net? 
Isn't it annoying when the block indentation is crap, 
you just can't read the code!

For example:
	Function Factorial(n As Long) As Long
	Dim i As Long, t As Long
	t = 1
	For i = 1 To n
	t = t * i
	Next i
	Factorial = t
	End Function

Compared to:
	Function Factorial(n As Long) As Long
	    Dim i As Long, t As Long
	    t = 1
	    For i = 1 To n
	        t = t * i
	    Next i
	    Factorial = t
	End Function

No contest, correct indentation increases readibilty no end.
Some developers just don't bother to do it (kill them all!).
So this add-in has been written to fix the problem. 
Just go to the code that has the problem and run Auto Indenter. 
This will go through the 
code and correctly 'tab' out blocks.

Installing and add-in can be a pain! Trying to register the
dll and then add it to the vbaddin.ini can be tricky so...
To Install:
	1. If files are zipped then unzip to a new folder.
	2. Open the project file (AutoIndent.vbp) in VB.
	3. Click 'File' then 'Make AutoIndent.dll'
	4. Save the dll in the same folder as the other files.
	5. Exit VB. Install complete!

To Use:
	1. Open the code in VB.
	2. Click 'Add-Ins' then 'Add-In Manager...'
	3. Find 'VB Code Auto Indenter'
	4. Check the 'Loaded' box (and 'Start-up' if you want 
it to to always load with VB)
	5. Click 'Add-Ins' then 'Auto Indent Code'
	6. Your code will be transformed! 
	   (Any error in code will result in wierd results.)

I think have taken into account all the various VB constructs.
If have missed anything then please email me (see above).
At the minute the indent is 4 spaces, I will try to get round 
to accessing the VB registry settings sometime! But for now 
you can either put up with 4 spaces (Which I think is best) 
or change the code. If you do modify the code then DO NOT 
restribute to others without consulting me first. 

I appreciate any queries and suggestions - email me.

Legal stuff:
I can accept no responsibilty for any loss or damage incurred by 
using this programme. Use it at your own risk. It is supplied to you for free so don't
sell it, although you may pass it on to others for free.

Copyright © 2000 Andrew Davey