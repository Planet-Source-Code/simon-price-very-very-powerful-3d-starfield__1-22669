FAST 3D STARFIELD DEMO BY SIMON PRICE

About

This program draws and animates a starfield. It is written in Visual Basic 6 and uses Win32 API function BUT NO DIRECTX or other graphics DLL's. If is designed with pure speed and processing power in mind, so there are few features. 

On my 400 Mhz PC, this program manages rendering 10000 stars easily at a speed of 30 frames per second! That's fast! PLEASE RUN THE COMPILED PROGRAM FOR SPEED. RUNNING FROM THE VB IDE WILL BE VERY SLOW.

You can choose the number of stars and the speed they move at upon program startup, but please do not change from the default values for the first time you run the program, since they are good values. Afterwards you can experiment as much as you like, I'd like to see how the program performs on other PC's!

Please vote for this code at www.planet-source-code.com/vb

Requirements

The VB6 and VB5 run time files, Windows 9x maybe 2k works too (I tested on 98), and a video card capable of 800 x 600 resolution in 24 bit color.

Logic

Here I'll try to explain how I designed this demo to run really fast. 

1/ There is no form you'll notice, the graphics are blitted straight onto the desktop! 
2/ All pixel manipulation is done using direct memory access, using pointers from VB! 
3/ Only 4 bytes are used per star, with the z position being calculated implicitly while rendering. 
4/ The star animation requires very little processing, since the stars are being moved along the z axis, and the z data is implicit and therefore not stored and does not need processing per star!
5/ The graphics are cleared in one API call to ZeroMemory, which is very fast.
6/ There are 4 if statements in the inner loop to check bounds - this is slow! I hear that it's possible to take these out by using And, but I can't get it to work. If anyone can fix that bit of source code, the program will run alot faster still! See the comments in the source for more details.

Disclaimer

Run this program at your own risk! I provide no guarantee that this works at all and any problems it causes you are your problems, not mine!

Credits

Mainly invented by myself, but some of the code to hide and show windows was taken and edited to my needs from PSC.

Contact Info

Email: Simon@VBgames.co.uk
Website: www.VBgames.co.uk