#!/usr/bin/env node
/*!
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

/*  own package information  */
const my          = require("./package.json")

/*  internal requirements  */
const path        = require("path")
const fs          = require("fs").promises

/*  external requirements  */
const yargs       = require("yargs")
const chalk       = require("chalk")
const execa       = require("execa")
const stripAnsi   = require("strip-ansi")
const JSZip       = require("jszip")
const Tmp         = require("tmp")
const mkdirp      = require("mkdirp")
const xmlformat   = require("xml-formatter")
const slimdom     = require("slimdom")
const slimdomSAX  = require("slimdom-sax-parser")
const fontoxpath  = require("fontoxpath")

;(async () => {
    /* eslint indent: off */

    /*  parsing command-line options  */
    const opts = yargs()
        .parserConfiguration({
            "set-placeholder-key": true,
            "halt-at-non-option":  true
        })
        .usage("Usage: pptx-surgeon [-v|--verbose <level>] [-k] [-r|font-remove-embedding] [-m|font-map-name <from>=<to>] <pptx-file>")
        .option("v", { alias: "verbose", type: "number",  describe: "level of verbose output", nargs: 1, default: 0 })
        .option("k", { alias: "keep", type: "boolean", describe: "keep expanded PPTX content", default: false })
        .option("r", { alias: "font-remove-embedding", type: "boolean",  describe: "remove font usages", default: false })
        .option("m", { alias: "font-map-name", type: "string", describe: "map font", nargs: 1, default: null })
        .version(false)
        .help(true)
        .showHelpOnFail(true)
        .strict(true)
        .parse(process.argv.slice(2))
    if (opts.fontMapName === null)
        opts.fontMapName = []
    else if (typeof opts.fontMapName === "string")
        opts.fontMapName = [ opts.fontMapName ]
    if (opts._.length !== 1)
        throw new Error("PPTX file argument missing")
    const pptxfile = opts._[0]

    /*  helper function for verbose log output  */
    const logLevels = [ "NONE", chalk.blue("INFO"), chalk.yellow("DEBUG") ]
    const log = (level, msg) => {
        if (level > 0 && level < logLevels.length && level <= opts.verbose) {
            msg = `pptx-surgeon: ${chalk.blue(logLevels[level])}: ${msg}\n`
            if (opts.outputNocolor || !process.stderr.isTTY)
                msg = stripAnsi(msg)
            process.stderr.write(msg)
        }
    }

    /*  create temporary filesystem area  */
    const tmp = Tmp.dirSync()
    let tmpdir = opts.keep ? pptxfile + ".d" : tmp.name
    log(1, `creating temporary directory ${chalk.blue(tmpdir)}`)

    /*
     *  ==== Save PPTX ====
     */

    /*  read PPTX file (which actually is just the unpacking of a ZIP format file)  */
    log(1, "unpacking PPTX content")
    let data = await fs.readFile(pptxfile, { encoding: null })
    log(1, `read PPTX file ${chalk.blue(pptxfile)}: ${data.length} bytes`)
    let zip = new JSZip()
    log(1, `parsing PPTX file ${chalk.blue(pptxfile)}`)
    let contents = await zip.loadAsync(data)
    let manifest = []
    for (filename of Object.keys(contents.files)) {
        let file = zip.file(filename)
        let content = await file.async("nodebuffer")
        let filepath = `${tmpdir}/${filename}`
        let filedir = path.dirname(filepath)
        await mkdirp.sync(filedir, { mode: 0o755 })
        await fs.writeFile(`${tmpdir}/${filename}`, content, { encoding: null })
        manifest.push(filename)
        log(2, `extracted PPTX part ${chalk.blue(filename)}: ${content.length} bytes`)
    }

    /*
     *  ==== Helper Functions for PPTX Management ====
     */

    /*  helper functions for XML manipulation  */
    const xmlLoad = async (filename) => {
        let xml = await fs.readFile(filename, { encoding: "utf8" })
        log(2, `loading PPTX part ${chalk.blue(filename)}: ${xml.length} bytes`)
        let dom = slimdomSAX.sync(xml, { position: false })
        let m
        if ((m = xml.match(/^(<\?xml.+?\?>\r?\n)/)) !== null)
            dom.PI = m[1]
        return dom
    }
    const xmlQuery = (dom, query, options = {}) => {
        options = Object.assign({}, { single: false, type: "nodes" }, options)
        let result
        if (options.type === "nodes")
            result = fontoxpath.evaluateXPathToNodes(query, dom)
        else if (options.type === "string")
            result = fontoxpath.evaluateXPathToStrings(query, dom)
        else if (options.type === "number")
            result = fontoxpath.evaluateXPathToNumbers(query, dom)
        if (options.single) {
            if (result.length === 0)
                result = undefined
            else if (result.length === 1)
                result = result[0]
            else
                throw new Error("requested single result, but multiple nodes found")
        }
        return result
    }
    const xmlEdit = async (dom, expr, options = {}) => {
        let result = await fontoxpath.evaluateUpdatingExpression(expr, dom)
        fontoxpath.executePendingUpdateList(result.pendingUpdateList)
    }
    const xmlSave = async (dom, filename) => {
        let xml = slimdom.serializeToWellFormedString(dom)
        if (!xml.match(/^<\?xml.+?\?>/) && dom.PI)
            xml = dom.PI + xml
        log(2, `saving PPTX part ${chalk.blue(filename)}: ${xml.length} bytes`)
        await fs.writeFile(filename, xml, { encoding: "utf8" })
    }

    /*  helper function for determining files of OpenXML  */
    const partsOfType = async (type, single = false) => {
        let xml = await xmlLoad(`${tmpdir}/[Content_Types].xml`)
        let result = xmlQuery(xml, `
            // Override [
                @ContentType = 'application/vnd.openxmlformats-officedocument.${type}+xml'
            ] / @PartName
        `, { single, type: "string" })
        if ((single && result === undefined) || (!single && result.length === 0))
            throw new Error(`no part found for type "${type}"`)
        if (typeof result === "string")
            result = result.replace(/^\//, "")
        else
            result = result.map((f) => f.replace(/^\//, ""))
        return result
    }

    /*
     *  ==== Display Font Information ====
     */

    /*  load presentation main XML  */
    let mainxml = await partsOfType("presentationml.presentation.main", true)
    let xml = await xmlLoad(`${tmpdir}/${mainxml}`)
    let ettf = xmlQuery(xml, "/ p:presentation / @embedTrueTypeFonts", { single: true, type: "string" })
    if (ettf)
        log(1, `PPTX: global flag: embedTrueTypeFonts="${ettf}"`)
    let efs = xmlQuery(xml, "// p:embeddedFontLst / p:embeddedFont")
    for (ef of efs) {
        let tf = xmlQuery(ef, ". / p:font / @typeface", { single: true, type: "string" })
        if (tf) {
            let styleNames = [ "regular", "bold", "italic", "boldItalic" ]
            for (styleName of styleNames) {
                let id = xmlQuery(ef, `. / p:${styleName} / @r:id`, { single: true, type: "string" })
                if (id)
                    log(1, `PPTX: embedded font: typeface=${tf}, style=${styleName}, id=${id}`)
            }
        }
    }

    /*  load relationships XML  */
    let rel = mainxml.replace(/^(.*\/)([^\/]+)$/, "$1_rels/$2.rels")
    xml = await xmlLoad(`${tmpdir}/${rel}`)
    let rels = xmlQuery(xml, `
        / Relationships
        / Relationship [
            @Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
        ]
    `)
    for (let rel of rels) {
        let id     = xmlQuery(rel, ". / @Id",     { single: true, type: "string" })
        let target = xmlQuery(rel, ". / @Target", { single: true, type: "string" })
        log(1, `PPTX: font relationship: id=${id}, target=${target}`)
    }

    /*  load document property XML  */
    let prop = await partsOfType("extended-properties", true)
    xml = await xmlLoad(`${tmpdir}/${prop}`)
    let titles = xmlQuery(xml, `
        / Properties
        / TitlesOfParts
        / vt:vector
        / vt:lpstr
    `, { type: "string" })
    for (let title of titles)
        log(1, `PPTX: document property part: title=${title}`)

    /*  load theme XML  */
    let themes = await partsOfType("theme")
    let T = {}
    for (theme of themes) {
        let xml = await xmlLoad(`${tmpdir}/${theme}`)
        let combis = [
            { id: "+mj-lt", section: "a:majorFont", type: "a:latin" },
            { id: "+mj-ea", section: "a:majorFont", type: "a:ea" },
            { id: "+mj-cs", section: "a:majorFont", type: "a:cs" },
            { id: "+mn-lt", section: "a:minorFont", type: "a:latin" },
            { id: "+mn-ea", section: "a:minorFont", type: "a:ea" },
            { id: "+mn-cs", section: "a:minorFont", type: "a:cs" },
        ]
        for (const combi of combis) {
            let tf = xmlQuery(xml, `// a:fontScheme / ${combi.section} / ${combi.type} / @typeface`,
                { single: true, type: "string" })
            if (tf !== undefined)
                T[combi.id] = tf
        }
    }
    Object.keys(T)
        .sort((a, b) => a.localeCompare(b))
        .map((font) => `${font}=${T[font]}`)
        .forEach((entry) => {
        log(1, `PPTX: theme font mapping: ${entry}`)
    })

    /*  load slide master and slide XMLs  */
    let types = [ "slideMaster", "slide" ]
    for (type of types) {
        let idx = {}
        let slides = await partsOfType(`presentationml.${type}`)
        for (slide of slides) {
            xml = await xmlLoad(`${tmpdir}/${slide}`)
            let tfs = xmlQuery(xml, `// * [ @typeface ] / @typeface`, { type: "string" })
            for (tf of tfs) {
                if (idx[tf] === undefined)
                    idx[tf] = 0
                idx[tf]++
            }
        }
        Object.keys(idx)
            .sort((a, b) => a.localeCompare(b))
            .map((font) => `${font}=${idx[font]}`)
            .forEach((entry) => {
            log(1, `PPTX: ${type} font usage: ${entry}`)
        })
    }

    /*
     *  ==== Optionally Remove Font Embedding Information ====
     */

    if (opts.fontRemoveEmbedding) {
        log(1, "remove font embedding information")

        /*  edit presentation main XML  */
        let mainxml = await partsOfType("presentationml.presentation.main", true)
        let xml = await xmlLoad(`${tmpdir}/${mainxml}`)
        await xmlEdit(xml, "delete node / p:presentation / @embedTrueTypeFonts")
        await xmlEdit(xml, "delete node // p:embeddedFontLst")
        await xmlSave(xml, `${tmpdir}/${mainxml}`)

        /*  edit relationships XML  */
        let relfile = mainxml.replace(/^(.*\/)([^\/]+)$/, "$1_rels/$2.rels")
        xml = await xmlLoad(`${tmpdir}/${relfile}`)
        let rels = xmlQuery(xml, `
            / Relationships
            / Relationship [
                @Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
            ]
        `)
        for (let rel of rels) {
            await xmlEdit(rel, "delete node .")
            let target = xmlQuery(rel, ". / @Target", { type: "string", single: true })
            target = target.replace(/^\//, "")
            let file = path.resolve(path.dirname(`${tmpdir}/${mainxml}`), target)
            let base = path.resolve(tmpdir)
            if (!(file.length >= base.length && file.substr(0, base.length) === base))
                throw new Error("fatal error")
            let part = file.substr(base.length).replace(/^\//, "")
            log(2, `removing PPTX part ${chalk.blue(part)}`)
            await fs.unlink(file)
            manifest = manifest.filter((filename) => filename !== part)
        }
        await xmlSave(xml, `${tmpdir}/${relfile}`)
    }

    /*
     *  ==== Optionally Map Font Name Information ====
     */

    /*  optionally map font  */
    if (opts.fontMapName.length > 0) {
        log(1, "map font name information")

        /*  prune theme XML files  */
        let themes = await partsOfType("theme")
        for (theme of themes) {
            let xml = await xmlLoad(`${tmpdir}/${theme}`)
            let sections = [ "a:majorFont", "a:minorFont "]
            let types = [ "a:latin", "a:ea", "a:cs", "a:font" ]
            for (section of sections) {
                for (type of types) {
                    for (fontmapname of opts.fontMapName) {
                        let m = fontmapname.match(/^(.+)=(.+)$/)
                        if (m == null)
                            throw new Error("invalid font mapping syntax")
                        let [ , valold, valnew ] = m
                        let xquf = `
                            for $n in // a:fontScheme / ${section} / ${type} [ @typeface = "${valold}" ]
                            return replace value of node $n / @typeface with "${valnew}"
                        `
                        await xmlEdit(xml, xquf)
                    }
                }
            }
            await xmlSave(xml, `${tmpdir}/${theme}`)
        }

        /*  prune slide master and slide XML files  */
        let types = [ "slideMaster", "slide" ]
        for (type of types) {
            let slides = await partsOfType(`presentationml.${type}`)
            for (slide of slides) {
                xml = await xmlLoad(`${tmpdir}/${slide}`)
                for (fontmapname of opts.fontMapName) {
                    let m = fontmapname.match(/^(.+)=(.+)$/)
                    if (m == null)
                        throw new Error("invalid font mapping syntax")
                    let [ , valold, valnew ] = m
                    let xquf = `
                        for $n in // * [ @typeface = "${valold}" ]
                        return replace value of node $n / @typeface with "${valnew}"
                    `
                    await xmlEdit(xml, xquf)
                }
                await xmlSave(xml, `${tmpdir}/${slide}`)
            }
        }
    }

    /*
     *  ==== Save PPTX ====
     */

    /*  write PPTX file (which actually is just the packing of a ZIP format file)  */
    log(1, "packing PPTX content")
    nzip = new JSZip()
    for (filename of manifest) {
        let content = await fs.readFile(`${tmpdir}/${filename}`, { encoding: null })
        log(2, `storing PPTX part ${chalk.blue(filename)}: ${content.length} bytes`)
        nzip.file(filename, content, {
            date:            zip.file(filename).date,
            unixPermissions: zip.file(filename).unixPermissions,
            comment:         zip.file(filename).comment,
            createFolders:   false
        })
    }
    data = await nzip.generateAsync({
        type:             "nodebuffer",
        compression:      "DEFLATE",
        compressionLevel: { level: 9 },
        comment:          `pptx-surgeon/${my.version}`,
        platform:         process.platform,
        streamFiles:      false
    })
    log(1, `write PPTX file ${chalk.blue(pptxfile)}: ${data.length} bytes`)
    await fs.writeFile(`${pptxfile}.new`, data, { encoding: null })

    /*  delete temporary filesystem area  */
    tmp.removeCallback()

    process.exit(0)

})().catch((err) => {
    /*  fatal error  */
    process.stderr.write(`pptx-surgeon: ${chalk.red("ERROR:")} ${err.stack}\n`)
    process.exit(1)
})

