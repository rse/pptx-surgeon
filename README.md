
pptx-surgeon
============

**Microsoft PowerPoint OpenXML File Surgeon**

<p/>
<img src="https://nodei.co/npm/pptx-surgeon.png?downloads=true&stars=true" alt=""/>

<p/>
<img src="https://david-dm.org/rse/pptx-surgeon.png" alt=""/>

Abstract
--------

This is a small utility for performing some font-related surgical
operations on Microsoft PowerPoint OpenXML files (PPTX). Microsoft
PowerPoint, as of at least January 2020, can produce PPTX files which
are "broken" when it comes to fonts:

1. **Broken Font Embedding**:
   Sometimes it contains some font embedding information,
   but the actual font data is partially missing and PowerPoint is not
   willing to correct this through any means of its user interface
   (including disabling and re-enabling the font embedding option or
   exporting and re-importing via XML format). For this, `pptx-surgeon`
   provides the possibility to completely remove all font embedding
   information which resets the PPTX file back to a consistent state
   (where PowerPoint again is willing to correctly embed the fonts from
   scratch).

2. **Hidden Font References**:
   Often it contains references to previously used fonts, although no
   user interface reachable shape any longer uses these fonts. For this,
   `pptx-surgeon` provides the possibility to map font names in both
   the theme, slide master and slides. This gets rid of the previous
   references to old fonts.

Installation
------------

- download pre-built binary for Windows (x64):<br/>
  https://github.com/rse/pptx-surgeon/releases/download/0.9.4/pptx-surgeon-win-x64.exe

- download pre-built binary for macOS (x64):<br/>
  https://github.com/rse/pptx-surgeon/releases/download/0.9.4/pptx-surgeon-mac-x64

- download pre-built binary for GNU/Linux (x64):<br/>
  https://github.com/rse/pptx-surgeon/releases/download/0.9.4/pptx-surgeon-lnx-x64

- via Node.js/NPM for any platform:<br/>
  `$ npm install -g pptx-surgeon`

Usage
-----

```
$ pptx-surgeon \
  [-v|--verbose <level>] \
  [-k|--keep-temporary] \
  [-o|--output <pptx-file>] \
  [-d|--font-dump-info] \
  [-r|--font-remove-embed] \
  [-m|--font-map-name <name-old>=<name-new>] \
  [-c|--font-cleanup <name-primary>,<name-secondary>,...] \
  <pptx-file>
```

Examples
--------

```
# show all font information
$ pptx-surgeon -d sample.pptx

# patch PPTX by removing font embeddings
$ pptx-surgeon -r -o sample-patched.pptx sample.pptx

# patch PPTX by mapping font names
$ pptx-surgeon -m "Arial=msg CI Text" -o sample-patched.pptx sample.pptx

# patch PPTX by performing an all-in-one cleanup
# (the listed fonts are all kept and everything else is mapped to "msg CI Text")
$ pptx-surgeon -c "msg CI Text,msg CI Signal,msg CS Code,msg CS Note,Wingdings,Symbol" \
  -o sample-patched.pptx sample.pptx
```

License
-------

Copyright &copy; 2020 Dr. Ralf S. Engelschall (http://engelschall.com/)

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

