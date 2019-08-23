# Overview

We recognize that nobody "chooses" to work in VB6 in 2017+, but for those who 
are forced to do so by legacy or contingency, we'll do our best to make your 
integration as painless as possible.

# Known Issues

VB6 doesn't support unsigned types.  The standard Spectrometer class and 
ISpectrometer interface in Wasatch.NET make heavy use of unsigned types
(e.g. spectrometer.pixels), as does EEPROM, etc.  My only thought on how to
support such types would be to create complete standalone interfaces like
SpectrometerVBA, comparable to the existing DriverVBA. I could do that, but
it would only help VB6 customers, and currently we don't have anyone lined
up requesting that, so I'm not doing it at this time.

Nor am I "breaking" the current Wasatch.NET API and exposing a lot of 
legitimately unsigned values as though they were signed.  That wouldn't
improve anyone's life :-(

# Design Thoughts

Our assumption going in was that Wasatch.NET works from VBA (see Wasatch.Excel),
so a VB6 COM version shouldn't be that difficult.

All testing was done under Win10-64 using Visual Basic 6.0, installed per this 
guidance:

https://www.fortypoundhead.com/showcontent.asp?artid=23993

The process would probably be a little different if tested under WinXP-32, which 
we can do as needed.

Note that one option is to rebuild Wasatch.NET under Visual Studio 2005 using an 
older .NET target; I haven't tried that.
