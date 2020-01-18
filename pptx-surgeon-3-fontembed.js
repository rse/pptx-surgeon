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

    /*  read font embedding information  */
    async read () {
        let info = {}

        /*  load presentation main XML  */
        const mainxml = await this.options.pptx.parts("presentationml.presentation.main", true)
        let xml = await this.options.xml.load(`${this.options.pptx.basedir}/${mainxml}`)
        const ettf = this.options.xml.query(xml,
            "/ p:presentation / @embedTrueTypeFonts", { single: true, type: "string" })
        info.fontEmbedFlag = !!ettf
        if (ettf)
            this.options.log(1, `PPTX: global flag: embedTrueTypeFonts="${ettf}"`)
        info.fontEmbedList = []
        const efs = this.options.xml.query(xml, "// p:embeddedFontLst / p:embeddedFont")
        for (const ef of efs) {
            const tf = this.options.xml.query(ef,
                ". / p:font / @typeface", { single: true, type: "string" })
            if (tf) {
                const styleNames = [ "regular", "bold", "italic", "boldItalic" ]
                for (const styleName of styleNames) {
                    const id = this.options.xml.query(ef, `. / p:${styleName} / @r:id`, { single: true, type: "string" })
                    if (id) {
                        this.options.log(1, "PPTX: embedded font: " +
                            `typeface=${chalk.blue(tf)}, style=${chalk.blue(styleName)}, id=${chalk.blue(id)}`)
                        info.fontEmbedList.push({ typeface: tf, style: styleName, id, asset: null })
                    }
                }
            }
        }

        /*  load relationships XML  */
        const rel = mainxml.replace(/^(.*\/)([^/]+)$/, "$1_rels/$2.rels")
        xml = await this.options.xml.load(`${this.options.pptx.basedir}/${rel}`)
        const rels = this.options.xml.query(xml, `
            / Relationships
            / Relationship [
                @Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
            ]
        `)
        for (const rel of rels) {
            const id     = this.options.xml.query(rel, ". / @Id",     { single: true, type: "string" })
            const target = this.options.xml.query(rel, ". / @Target", { single: true, type: "string" })
            this.options.log(1, `PPTX: font relationship: id=${chalk.blue(id)}, target=${chalk.blue(target)}`)
            let size = await this.options.pptx.partSize(mainxml, target)
            let entry = info.fontEmbedList.find((entry) => entry.id === id)
            if (entry) {
                entry.asset = target
                entry.size  = size
            }
            else
                info.fontEmbedList.push({ typeface: null, style: null, id, asset: target, size })
        }

        /*  remove internal ids  */
        info.fontEmbedList.forEach((entry) => delete entry.id)

        return info
    }

    /*  delete font embedding information  */
    async delete () {
        this.options.log(1, "remove font embedding information")

        /*  edit presentation main XML  */
        const mainxml = await this.options.pptx.parts("presentationml.presentation.main", true)
        let xml = await this.options.xml.load(`${this.options.pptx.basedir}/${mainxml}`)
        await this.options.xml.edit(xml, "delete node / p:presentation / @embedTrueTypeFonts")
        await this.options.xml.edit(xml, "delete node // p:embeddedFontLst")
        await this.options.xml.save(xml, `${this.options.pptx.basedir}/${mainxml}`)

        /*  edit relationships XML  */
        const relfile = mainxml.replace(/^(.*\/)([^/]+)$/, "$1_rels/$2.rels")
        xml = await this.options.xml.load(`${this.options.pptx.basedir}/${relfile}`)
        const rels = this.options.xml.query(xml, `
            / Relationships
            / Relationship [
                @Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
            ]
        `)
        for (const rel of rels) {
            await this.options.xml.edit(rel, "delete node .")
            let target = this.options.xml.query(rel, ". / @Target", { type: "string", single: true })
            target = target.replace(/^\//, "")
            await this.options.pptx.partDelete(mainxml, target)
        }
        await this.options.xml.save(xml, `${this.options.pptx.basedir}/${relfile}`)
    }
}

