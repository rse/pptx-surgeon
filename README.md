
pptx-surgeon
============

**PowerPoint OpenXML File Surgeon**

<p/>
<img src="https://nodei.co/npm/pptx-surgeon.png?downloads=true&stars=true" alt=""/>

<p/>
<img src="https://david-dm.org/rse/pptx-surgeon.png" alt=""/>

Abstract
--------

This is a small utility for performing some font-related
surgical operations on PowerPoint OpenXML files (PPTX).
PowerPoint as of at least January 2020 sometimes produces
"broken" PPTX files:

1. Sometimes it contains some font embedding information,
   but the actual font data is partially missing and PowerPoint is not
   willing to correct this through any means of its user interface. For
   this `pptx-surgeon` provides the option to completely remove all font
   embedding information which resets the PPTX file back to a consistent
   state.

2. Often it contains references to previously used fonts although no
   user interface reachable shape any longer uses these fonts. For this
   `pptx-surgeon` provides the option to map font names.

Installation
------------

```
$ npm install -g pptx-surgeon
```

License
-------

Copyright (c) 2020 Dr. Ralf S. Engelschall (http://engelschall.com/)

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be included
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

