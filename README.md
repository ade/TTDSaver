# TTDSaver

This is the Visual Basic 6 source for my TTD screen saver I made way back. All the relevant files I could find are included in this repository. The original readme follows below.

aDe's TTD Saver - version 1.1 / 2003-01-08
------------------------------------------------------------------------------------

Hello and thanks for downloading my TTD screensaver. It features a random flat piece of land, with random 
buildings and a train track, and on the track is a train. It's not too complicated, a simple install is enough to 
operate it, but you can also:
  - change the number of buildings drawn
  - change the length of the train
  - change/add graphics
  - more options

But there's also a few things you can't:
  - change resolution. Only 1024x768 is supported!
  - have elevated terrain. Sorry, too much work =)

If you want to know more about how to customize the graphics, read on. It requires some skills in general file 
management and image editing.


Changing / Adding graphics
--------------------------------------
First, you need to export the graphics so you can edit them. Open up the settings page of the saver by going to 
the display properties in windows. Then select a path where you want to keep the tilesets, and click the export 
button. You'll now (hopefully) have some BMP files created in that directory, along with a TXT file. If you want to 
change a graphic, just open up and edit one of the BMP's.

But if you need to ADD more..:
- for train engine/wagon graphic, open up the right BMP and add 18 pixels to the bottom, then just add in the 
new car in a similar pattern as the others.

- for building graphic, open up a building BMP and add 128 pixels to the bottom for large buildings 
(buildings_big.bmp), 64 pixels for medium/small buildings (buildings_med.bmp)

It is important you make these resizes exact, or things will be messed up.
If you're using photoshop, make sure you do a "Canvas resize" and not a "Image resize" so that you don't 
stretch the existing images. In canvas resize, press the arrow facing straight up. For other programs you'll need 
to figure it out yourself :)

The transparent color is determined from the top left picture in the BMP.
Read on if you are adding buildings...


The "tileset.txt" file
---------------------------
Notice that it's changed from tilesets.txt to tileset.txt!
Don't edit the txt file with Word or Wordpad or something like that, because you need to save the file in plain 
text format.. so use notepad.

Here you need to set the amount of buildings in the BMP file. You start out with 16 small buildings and 8 big 
ones. But that's already written in the file, just add to those values, you'll see corresponding filenames in the txt 
file.

You also set wheter engines are doubleheaded or not (on engine in front, one reversed in back).
The first engine in the bitmap is engine0, second engine1 and so on. Simply set the value to 1 to enable.

The best would be if you made a copy of the "default" directory in the "tilesets" directory where you installed the saver, rename the directory to your custom name and then experiment from there. Then simply zip up the directory and send it to others =)


Changelog
v. 1.1
------------------------------
  - added train styles (random, uniform, semiuniform)
  - fixed the directx controls popup error (probably)
  - changed the way tilesets are handled. multiple ones can be installed, and the one to be used selected in settings
  - added double train engines
  - added monorail tileset by Prof. Frink
  - added random tileset mode
  - added direct/progressive track draw option
  - fixed screensaver preview bug
  - lots of bugfixes

v. 1.0
------------------------------
  - initial release
