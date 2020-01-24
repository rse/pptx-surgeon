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

const path   = require("path")
const fs     = require("fs").promises
const chalk  = require("chalk")
const JSZip  = require("jszip")
const Tmp    = require("tmp")
const mkdirp = require("mkdirp")

module.exports = class Archive {
    /*  configuration options  */
    constructor (options = {}) {
        this.options = Object.assign({}, {
            log:     (level, msg) => undefined,
            xml:     null,
            basedir: null,
            keep:    false,
            tool:    "pptx-surgeon"
        }, options)
    }

    /*  load PPTX file  */
    async load (pptxfile) {
        /*  determine base directory  */
        this.tmp = null
        this.basedir = this.options.basedir
        if (this.basedir === null) {
            this.tmp = Tmp.dirSync()
            this.basedir = this.tmp.name
            this.options.log(1, `created temporary directory ${chalk.blue(this.basedir)}`)
        }

        /*  read PPTX file  */
        const data = await fs.readFile(pptxfile, { encoding: null })
        this.options.log(1, `read PPTX file ${chalk.blue(pptxfile)}: ${data.length} bytes`)

        /*  parse and unpack PPTX file  */
        this.options.log(1, `parsing PPTX file ${chalk.blue(pptxfile)}`)
        this.zip = new JSZip()
        const contents = await this.zip.loadAsync(data)
        this.manifest = []
        for (const filename of Object.keys(contents.files)) {
            /*  unpack OpenXML content part  */
            const file = this.zip.file(filename)
            const content = await file.async("nodebuffer")
            const filepath = `${this.basedir}/${filename}`
            const filedir = path.dirname(filepath)
            await mkdirp.sync(filedir, { mode: 0o755 })
            await fs.writeFile(`${this.basedir}/${filename}`, content, { encoding: null })
            this.manifest.push(filename)
            this.options.log(2, `retrieved OpenXML part ${chalk.blue(filename)}: ${content.length} bytes`)
        }
    }

    /*  determine OpenXML part(s)  */
    async parts (type, single = false) {
        if (this.options.xml === null)
            throw new Error("require XML manipulation instance")
        const xml = await this.options.xml.load(`${this.basedir}/[Content_Types].xml`)
        let result = this.options.xml.query(xml, `
            // Override [
                @ContentType = 'application/vnd.openxmlformats-officedocument.${type}+xml'
            ] / @PartName
        `, { single, type: "string" })
        if (typeof result === "string")
            result = result.replace(/^\//, "")
        else if (typeof result === "object" && result instanceof Array)
            result = result.map((f) => f.replace(/^\//, ""))
        return result
    }

    /*  determine size of OpenXML part  */
    async partSize (basePart, target) {
        const file = path.resolve(path.resolve(this.basedir, path.dirname(basePart)), target)
        const base = path.resolve(this.basedir)
        if (!(file.length >= base.length && file.substr(0, base.length) === base))
            throw new Error("fatal internal error: part not under base directory")
        const stats = await (fs.stat(file).catch(() => undefined))
        return (stats ? stats.size : 0)
    }

    /*  delete OpenXML part  */
    async partDelete (basePart, target) {
        const file = path.resolve(path.resolve(this.basedir, path.dirname(basePart)), target)
        const base = path.resolve(this.basedir)
        if (!(file.length >= base.length && file.substr(0, base.length) === base))
            throw new Error("fatal internal error: part not under base directory")
        const part = file.substr(base.length).replace(/\\/g, "/").replace(/^\//, "")
        this.options.log(2, `removing PPTX part ${chalk.blue(part)}`)
        await (fs.unlink(file).catch(() => undefined))
        this.manifest = this.manifest.filter((filename) => filename !== part)
    }

    /*  backup PPTX file  */
    async backup (pptxfile, backupfile) {
        await fs.copyFile(pptxfile, backupfile)
    }

    /*  save PPTX file  */
    async save (pptxfile) {
        /*  pack PPTX file  */
        this.options.log(1, "packing PPTX content")
        const nzip = new JSZip()
        for (const filename of this.manifest) {
            const content = await fs.readFile(`${this.basedir}/${filename}`, { encoding: null })
            this.options.log(2, `storing OpenXML part ${chalk.blue(filename)}: ${content.length} bytes`)
            nzip.file(filename, content, {
                date:            this.zip.file(filename).date,
                unixPermissions: this.zip.file(filename).unixPermissions,
                comment:         this.zip.file(filename).comment,
                createFolders:   false
            })
        }
        const data = await nzip.generateAsync({
            type:             "nodebuffer",
            compression:      "DEFLATE",
            compressionLevel: { level: 9 },
            comment:          `manipulated with ${this.options.tool}`,
            platform:         process.platform,
            streamFiles:      false
        })

        /*  write PPTX file  */
        this.options.log(1, `write PPTX file ${chalk.blue(pptxfile)}: ${data.length} bytes`)
        await fs.writeFile(pptxfile, data, { encoding: null })

        /*  delete temporary filesystem area  */
        if (this.tmp !== null && !this.options.keep) {
            this.tmp.removeCallback()
            this.tmp = null
        }
    }
}

