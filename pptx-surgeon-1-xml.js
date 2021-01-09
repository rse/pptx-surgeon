/*
**  pptx-surgeon -- PowerPoint OpenXML File Surgeon
**  Copyright (c) 2020-2021 Dr. Ralf S. Engelschall <rse@engelschall.com>
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

const fs          = require("fs").promises
const chalk       = require("chalk")
const slimdom     = require("slimdom")
const slimdomSAX  = require("slimdom-sax-parser")
const fontoxpath  = require("fontoxpath")
const xmlformat   = require("xml-formatter")

module.exports = class XML {
    /*  configuration options  */
    constructor (options = {}) {
        this.options = Object.assign({}, {
            log:     (level, msg) => undefined
        }, options)
    }

    /*  load XML into DOM  */
    async load (filename) {
        const xml = await fs.readFile(filename, { encoding: "utf8" })
        this.options.log(2, `loading XML file ${chalk.blue(filename)}: ${xml.length} bytes`)
        const dom = slimdomSAX.sync(xml, { position: false })
        let m
        if ((m = xml.match(/^(<\?xml.+?\?>\r?\n)/)) !== null)
            dom.PI = m[1]
        return dom
    }

    /*  dump DOM node(s)  */
    dump (dom) {
        let out = ""
        if (typeof dom === "object" && dom instanceof Array) {
            for (const el of dom) {
                let xml = slimdom.serializeToWellFormedString(el)
                xml = xmlformat(xml)
                if (out !== "")
                    out += "====\n"
                out += xml
            }
        }
        else {
            let xml = slimdom.serializeToWellFormedString(dom)
            xml = xmlformat(xml)
            out = xml
        }
        return out
    }

    /*  query DOM nodes via XPath (https://www.w3.org/TR/xpath-31/)  */
    query (dom, query, options = {}) {
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

    /*  edit DOM via XQuery Update Facility (https://www.w3.org/TR/xquery-update-30/)  */
    async edit (dom, expr) {
        const result = await fontoxpath.evaluateUpdatingExpression(expr, dom)
        fontoxpath.executePendingUpdateList(result.pendingUpdateList)
    }

    /*  save DOM to XML  */
    async save (dom, filename) {
        let xml = slimdom.serializeToWellFormedString(dom)
        if (!xml.match(/^<\?xml.+?\?>/) && dom.PI)
            xml = dom.PI + xml
        this.options.log(2, `saving XML file ${chalk.blue(filename)}: ${xml.length} bytes`)
        await fs.writeFile(filename, xml, { encoding: "utf8" })
    }
}

