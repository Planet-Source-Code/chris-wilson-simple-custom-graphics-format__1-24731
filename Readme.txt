VB Graphics Conversion 

By Chris Wilson, 2001

This is a simple program you can use to protect the graphics in your program.
It really just fakes a custom file format for bitmap graphics, by taking out a few bytes.

You can use this utility to convert your bitmaps to "GFX" files or whatever you want.

Usage:

copy the FixBitmap and revertBitmap subs into your program, before you load a graphic
fixbitmap(Filename) to put the bytes back in, then once you load it RevertBitmap(Filename)
to put it back.

e.g.
    FixBitmap App.Path & "\gra\" & uDDSurf.Filename
    Set uDDSurf.Surface = DD.CreateSurfaceFromFile(App.Path & "\gra\" & uDDSurf.Filename, uDDSurf.SurfaceDesc)
    RevertBitmap App.Path & "\gra\" & uDDSurf.Filename

Have fun!
