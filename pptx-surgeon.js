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

/*  external requirements  */
const yargs       = require("yargs")
const chalk       = require("chalk")
const stripAnsi   = require("strip-ansi")
const jsYAML      = require("js-yaml")

/*  own package information  */
const my          = require("./package.json")
const XML         = require("./pptx-surgeon-1-xml")
const PPTX        = require("./pptx-surgeon-2-pptx")
const FontEmbed   = require("./pptx-surgeon-3-fontembed")
const FontRefs    = require("./pptx-surgeon-4-fontrefs")

;(async () => {
    /*  parse command-line options  */
    const opts = yargs()
        .parserConfiguration({
            "set-placeholder-key": true,
            "halt-at-non-option":  true
        })
        .usage(
            "Usage: pptx-surgeon" +
            " [-v|--verbose <level>]" +
            " [-k|--keep-temporary]" +
            " [-o|--output <pptx-file>]" +
            " [-d|--font-dump-info]" +
            " [-r|--font-remove-embed]" +
            " [-m|--font-map-name <from>=<to>]" +
            " <pptx-file>"
        )
        .option("v", {
            alias:    "verbose",
            type:     "number",
            describe: "level of verbose output",
            nargs:    1,
            default:  0
        })
        .option("k", {
            alias:    "keep-temporary",
            type:     "boolean",
            describe: "keep expanded PPTX content",
            default:  false
        })
        .option("o", {
            alias:    "output",
            type:     "string",
            describe: "output file",
            default:  ""
        })
        .option("d", {
            alias:    "font-dump",
            type:     "boolean",
            describe: "dump font information",
            default:  false
        })
        .option("r", {
            alias:    "font-remove-embed",
            type:     "boolean",
            describe: "remove font embeddings",
            default:  false
        })
        .option("m", {
            alias:    "font-map-name",
            type:     "string",
            describe: "map font names",
            nargs:    1,
            default:  null
        })
        .version(false)
        .help(true)
        .showHelpOnFail(true)
        .strict(true)
        .demand(1)
        .parse(process.argv.slice(2))
    const pptxfile = opts._[0]
    const pptxfileOut = opts.output !== "" ? opts.output : pptxfile

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

    /*  create XML manipulation facility  */
    const xml = new XML({ log })

    /*  load PPTX file  */
    const pptx = new PPTX({ log, xml, keep: opts.keepTemporary, tool: `pptx-surgeon/${my.version}` })
    await pptx.load(pptxfile)

    /*  display font embedding information  */
    const fontembed = new FontEmbed({ log, xml, pptx })
    const info1 = await fontembed.read()
    if (opts.fontDumpInfo)
        process.stdout.write(jsYAML.safeDump(info1, {}))

    /*  display font references  */
    const fontrefs = new FontRefs({ log, xml, pptx })
    const info2 = await fontrefs.read()
    if (opts.fontDumpInfo)
        process.stdout.write(jsYAML.safeDump(info2, {}))

    /*  optionally remove font embeddings  */
    if (opts.fontRemoveEmbed)
        await fontembed.delete()

    /*  optionally map font  */
    if (opts.fontMapName !== null) {
        let fontMapNames = opts.fontMapName
        if (typeof fontMapNames === "string")
            fontMapNames = [ fontMapNames ]
        const mappings = []
        for (const fontMapName of fontMapNames) {
            const m = fontMapName.match(/^(.+)=(.+)$/)
            if (m == null)
                throw new Error("invalid font mapping syntax")
            const [ , from, to ] = m
            mappings.push({ from, to })
        }
        await fontrefs.map(mappings)
    }

    /*  save PPTX file  */
    if (pptxfile !== pptxfileOut)
        await pptx.save(pptxfileOut)
    else {
        await pptx.backup(pptxfile, `${pptxfile}.bak`)
        await pptx.save(pptxfile)
    }

    /*  gracefully terminate  */
    process.exit(0)
})().catch((err) => {
    /*  fatal error  */
    process.stderr.write(`pptx-surgeon: ${chalk.red("ERROR:")} ${err.stack}\n`)
    process.exit(1)
})

