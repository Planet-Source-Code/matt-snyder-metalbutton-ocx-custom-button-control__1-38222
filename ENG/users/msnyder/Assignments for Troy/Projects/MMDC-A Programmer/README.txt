---------------------------------------
MetalButton OCX - Custom Button Control
---------------------------------------


Updated 08/23/02 at 3:30 PM
---------------------------
> Fixed the flickering problem. (Let me know if anyone still notices it please.)
> Added two more color options, red and blue, which can be set with the "ButtonColor" property
	- Green: ButtonColor = 0
	- Blue: ButtonColor = 1
	- Red: ButtonColor = 2
> Have not masked the corners yet, but I will get to that in a later update.

Updated 08/26/02 at 9:30 AM
---------------------------
> Fixed the flickering problem, once and for all.
> Added realtime resizing
> Added the ability to change the button's font properties
> Have not masked the corners yet, but I will get to that in a later update. Really, I promise :)

Updated 08/26/02 at 5:00 PM
---------------------------
> Masked the corners, finally!


Source Code
-----------
You may still have to set up these objects in the source code to get it to work:

imgButtonNormal(0).Picture = GreenButton.bmp
imgButtonOver(0).Picture = GreenButtonOver.bmp
imgButtonClick(0).Picture = GreenButtonDown.bmp
imgButtonNormal(1).Picture = BlueButton.bmp
imgButtonOver(1).Picture = BlueButtonOver.bmp
imgButtonClick(1).Picture = BlueButtonDown.bmp
imgButtonNormal(2).Picture = RedButton.bmp
imgButtonOver(2).Picture = RedButtonOver.bmp
imgButtonClick(2).Picture = RedButtonDown.bmp


Included OCX Instructions
-------------------------
Extract the MetalButton.renameMe file to your C:\Windows\System or C:\WINNT\System32
folder and rename it to MetalButton.ocx for it to work. If there are anymore questions
or comments let me know so I may address them ASAP. Thanks.

Enjoy,

Matt Snyder
msnyder@universitymail.com