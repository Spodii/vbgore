===============================
== What are 3rd party tools? ==
===============================

Third party tools are tools submitted by other users for vbGORE, but are not
maintained or updated by vbGORE. Any problems with these tools will have to be
directed to the owners of the tools, and not the vbGORE staff. The vbGORE staff
will try to help you with the tools, but do not be surprised if they do not
know how to use them.

Any questions on the tools can be directed to the vbGORE Questions section of the forum:
http://www.vbgore.com/forums/viewforum.php?f=6

=============================================
== Codehead's Bitmap Font Generator (CBFG) ==
=============================================
Created by:
Codehead

Home page:
http://www.codehead.co.uk/cbfg/

Description:
CBFG is a program created for making just as the same suggests - bitmap fonts.
Bitmap fonts are used in vbGORE to display the custom fonts, which is much faster
and much more flexable then using system fonts. vbGORE supports, by default, the
Binary Data Format (DAT) file export.

Note that CBFG was not created for vbGORE, but was a tool created by a friend of Spodi's.
It is included into vbGORE with the owner's permission.

The guide for using CBFG can be found on the vbGORE site at:
http://www.vbgore.com/index.php?title=Using_CBFG

==================
== Grh3RawMaker ==
==================
Created by:
Van

Description:
Grh3RawMaker is a quite outdated program created to help ease the creation of Grh values.
The program still works, but not as well as it did back when it was released back for
version 0.0.2 of vbGORE. It should still get the job done, though.

The guide for Grh3RawMaker can be found on the vbGORE site at:
http://www.vbgore.com/index.php?title=Using_GrhRawMaker

=================
== Grh Crafter ==
=================
Created by:
The Drasil team (http://www.vbgore.com/index.php?title=Drasil)

Description:
Grh Crafter is a replacement for Grh3RawMaker, created by the staff of Drasil. It is supposed
to help you create your Grh1.raw file even easier then the Grh3RawMaker, along with is much
newer so is more likely to work with the later versions. Use whichever you like more, though,
as Grh Crater and Grh3RawMaker are both good tools.

=============
== OptiPNG ==
=============
Created by:
Spodi

OptiPNG website:
http://optipng.sourceforge.net/

PNGRewrite website:
http://entropymine.com/jason/pngrewrite/

PNGOut website:
http://advsys.net/ken/utils.htm#pngout

ADVPng website:
http://advancemame.sourceforge.net/comp-readme.html

Description:
The OptiPNG project is a wrapper for the OptiPNG compression library that is designed to optimize the
file format of PNG files, the main format used by vbGORE, to a smaller size. The library works by restructuring 
the data in the file to be more effecient. Although it is a lossy compression in a technical sense, the lost 
data is not a part of the visual display. Only unused information that does not affect the visual aspect 
of the PNG files is changed.
OptiPNG now also offers compression with pngout, advpng and pngrewrite. The compression takes at least twice 
as long now, but it is definitely worth it. It goes in the order of pngrewrite -> OptiPNG -> pngout -> advpng.

=========
== UPX ==
=========
Created by:
Spodi

UPX website:
http://upx.sourceforge.net/

Description:
The UPX project is a quick wrapper made for UPX, "the Ultimate Packer for eXecuteables". UPX compresses
EXE files into a smaller size without corrupting them or losing any data. The cost is a bit slower of a runtime,
but nothing you should notice or worry about. The UPX wrapper runs with the parameter --brute which forces the
best compression ratio possible. This project is mostly just a copy-and-paste of the OptiPNG project, but with
minor changes to support using UPX on EXE files instead of OptiPNG on PNG files.