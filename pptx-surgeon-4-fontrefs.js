/*
**  pptx-surgeon -- PowerPoint OpenXML File Surgeon
**  Copyright (c) 2020 Dr. Ralf S. Engelschall <rse@engelschall.com>
**
**  Permission is hereby granted, free of charge, to any person obtaining
**  a copy of this software and associated documentation files (the
**  "Software"), to deal in the Software without restriction, including
**  without limitation the rights to use, copy, modify, merge, publish,
**  distribute, sublicense, and/or sell copies of the Software, and to
**  permit persons to whom the Software is furnished to do so, subject to
**  the following conditions:
**
**  The above copyright notice and this permission notice shall be included
**  in all copies or substantial portions of the Software.
**
**  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
**  EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
**  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
**  IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
**  CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
**  TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
**  SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/

const chalk = require("chalk")

module.exports = class FontEmbed {
    /*  configuration options  */
    constructor (options = {}) {
        this.options = Object.assign({}, {
            log:     (level, msg) => undefined,
            xml:     null,
            pptx:    null
        }, options)
        if (this.options.xml === null)
            throw new Error("require XML facility")
        if (this.options.pptx === null)
            throw new Error("require PPTX facility")
    }

    /*  read font reference information  */
    async read () {
        const info = {}

        /*  load theme XML  */
        const themes = await this.options.pptx.parts("theme")
        info.fontTheme = {}
        for (const theme of themes) {
            const xml = await this.options.xml.load(`${this.options.pptx.basedir}/${theme}`)
            const combis = [
                { id: "+mj-lt", section: "a:majorFont", type: "a:latin" },
                { id: "+mj-ea", section: "a:majorFont", type: "a:ea" },
                { id: "+mj-cs", section: "a:majorFont", type: "a:cs" },
                { id: "+mn-lt", section: "a:minorFont", type: "a:latin" },
                { id: "+mn-ea", section: "a:minorFont", type: "a:ea" },
                { id: "+mn-cs", section: "a:minorFont", type: "a:cs" }
            ]
            for (const combi of combis) {
                const tf = this.options.xml.query(xml,
                    `// a:fontScheme / ${combi.section} / ${combi.type} / @typeface`,
                    { single: true, type: "string" })
                if (tf !== undefined)
                    info.fontTheme[combi.id] = tf
            }
        }
        Object.keys(info.fontTheme)
            .sort((a, b) => a.localeCompare(b))
            .map((font) => `${font}=${info.fontTheme[font]}`)
            .forEach((entry) => {
                this.options.log(1, `PPTX: theme font mapping: ${chalk.blue(entry)}`)
            })

        /*  load slide layout, slide master and slide XMLs  */
        info.fontRefs = {}
        const types = [ "slideLayout", "slideMaster", "slide", "notesMaster", "notesSlide" ]
        for (const type of types) {
            info.fontRefs[type] = {}
            const slides = await this.options.pptx.parts(`presentationml.${type}`)
            for (const slide of slides) {
                const xml = await this.options.xml.load(`${this.options.pptx.basedir}/${slide}`)
                const tfs = this.options.xml.query(xml,
                    "// * [ @typeface ] / @typeface", { type: "string" })
                for (const tf of tfs) {
                    if (info.fontRefs[type][tf] === undefined)
                        info.fontRefs[type][tf] = 0
                    info.fontRefs[type][tf]++
                }
            }
            Object.keys(info.fontRefs[type])
                .sort((a, b) => a.localeCompare(b))
                .map((font) => `${font}=${info.fontRefs[type][font]}`)
                .forEach((entry) => {
                    this.options.log(1, `PPTX: ${type} font usage: ${chalk.blue(entry)}`)
                })
        }

        return info
    }

    /*  delete font reference information  */
    async map (mappings) {
        this.options.log(1, "map font references")

        /*  patch theme XML files  */
        const themes = await this.options.pptx.parts("theme")
        for (const theme of themes) {
            const xml = await this.options.xml.load(`${this.options.pptx.basedir}/${theme}`)
            const sections = [ "a:majorFont", "a:minorFont "]
            const types = [ "a:latin", "a:ea", "a:cs", "a:font" ]
            for (const section of sections) {
                for (const type of types) {
                    for (const mapping of mappings)
                        await this.options.xml.edit(xml, `
                            for $n in // a:fontScheme / ${section} / ${type} [ @typeface = "${mapping.from}" ]
                            return replace value of node $n / @typeface with "${mapping.to}"`)
                }
            }
            await this.options.xml.save(xml, `${this.options.pptx.basedir}/${theme}`)
        }

        /*  patch slide layout, slide master and slide XML files  */
        const types = [ "slideLayout", "slideMaster", "slide", "notesMaster", "notesSlide" ]
        for (const type of types) {
            const slides = await this.options.pptx.parts(`presentationml.${type}`)
            for (const slide of slides) {
                const xml = await this.options.xml.load(`${this.options.pptx.basedir}/${slide}`)
                for (const mapping of mappings)
                    await this.options.xml.edit(xml, `
                        for $n in // * [ @typeface = "${mapping.from}" ]
                        return replace value of node $n / @typeface with "${mapping.to}"`)
                await this.options.xml.save(xml, `${this.options.pptx.basedir}/${slide}`)
            }
        }
    }
}

