#!/usr/bin/env node
/*!
**  psd2pptx -- Convert Photoshop (PSD) layers to PowerPoint (PPTX) slides
**  Copyright (c) 2019 Dr. Ralf S. Engelschall <rse@engelschall.com>
**  Licensed under MIT Open Source license.
*/

/*  internal requirements  */
const fs         = require("fs")

/*  external requirements  */
const tmp        = require("tmp")
const mkdirp     = require("mkdirp")
const bluebird   = require("bluebird")
const execa      = require("execa")
const Jimp       = require("jimp")
const PPTXGenJS  = require("pptxgenjs")
const yargs      = require("yargs")
const chalk      = require("chalk")
const PSD        = require("psd.js")
const zipProcess = require("zip-process")

/*  act in an asynchronous context  */
;(async () => {
    /*  command-line option parsing  */
    const argv = yargs
        /* eslint indent: off */
        .usage("Usage: $0 [-h] [-v] [-o <pptx-file>] <psd-file>")
        .help("h").alias("h", "help").default("h", false)
            .describe("h", "show usage help")
        .boolean("v").alias("v", "verbose").default("v", false)
            .describe("v", "print verbose messages")
        .string("o").nargs("o", 1).alias("o", "output").default("o", "")
            .describe("o", "output PPTX file")
        .string("c").nargs("c", 1).alias("c", "canvas").default("c", "Canvas")
            .describe("c", "name of canvas layer group (default: \"Canvas\")")
        .string("s").nargs("s", 1).alias("s", "skip").default("s", "^Background$")
            .describe("s", "regular expression matching layers to skip (default: \"^Background$\")")
        .version(false)
        .strict()
        .showHelpOnFail(true)
        .demand(1)
        .parse(process.argv.slice(2))

    /*  verbose message printing  */
    const verbose = (msg) => {
        if (argv.verbose)
            process.stdout.write(`++ ${msg}\n`)
    }

    /*  create temporary filesystem area  */
    const tmpdir = tmp.dirSync()

    /*  read and parse PSD file  */
    const psdfile = argv._[0]
    verbose(`reading PSD file: ${chalk.blue(psdfile)}`)
    const psd = PSD.fromFile(psdfile)
    psd.parse()

    /*  extract PSD layers (as PNG files)  */
    const basename = psdfile.replace(/\.psd$/, "")
    let layers = []
    const walkNode = async (node) => {
        if (!node)
            return
        if (node.hasChildren()) {
            /*  recursively enter child layers  */
            await bluebird.each(node.children(), async (child) => {
                await walkNode(child)
            })
        }
        else if (node.layer.image) {
            /*  skip hard-coded Procreate background layer  */
            if (node.path() === "Background")
                return
            verbose(`extracting layer: ${chalk.blue(node.path())}`)
            let path = node.path().replace(/\//g, "-").replace(/[^a-zA-Z0-9.]/g, "")
            let pngfile = `${tmpdir.name}/extracted-${path}.png`
            await node.layer.image.saveAsPng(pngfile)
            layers.push({
                file:    pngfile,
                name:    node.get("name"),
                path:    node.path(),
                opacity: node.layer.opacity,
                width:   node.layer.width,
                height:  node.layer.height
            })
        }
    }
    await walkNode(psd.tree())

    /*  skip some layers  */
    let regex1 = new RegExp(argv.skip)
    layers = layers.filter((item) => !item.path.match(regex1))

    /*  divide layers into canvas, slides and other layers  */
    let regex2 = new RegExp(`^${argv.canvas}\/`, "i")
    let layersCanvas = layers.filter((item)       =>  item.path.match(regex2))
    let layersSlides = layers.filter((item)       => !item.path.match(regex2))
    let layersOthers = layersSlides.filter((item) => !item.path.match(/^.+\/[^\/]+$/i))
    layersSlides     = layersSlides.filter((item) =>  item.path.match(/^.+\/[^\/]+$/i))

    /*  generate canvas  */
    let bottom = layersCanvas.slice(-1)[0]
    let w = bottom.width
    let h = bottom.height
    verbose(`generating canvas: ${layersCanvas.map((item) => chalk.blue(item.path)).join(", ")} (${w}x${h})`)
    let canvas = await Jimp.read(bottom.file)
    await canvas.opacity(bottom.opacity / 255)
    await bluebird.each(layersCanvas.slice(0, -1).reverse(), async (item) => {
        let src = await Jimp.read(item.file)
        await src.opacity(item.opacity / 255)
        await canvas.blit(src, 0, 0)
    })
    let canvasFilename = `${tmpdir.name}/canvas.png`
    await canvas.writeAsync(canvasFilename)

    /*  generate empty image  */
    const empty = await new Promise((resolve, reject) => {
        new Jimp(w, h, 0x00000000, (err, image) => {
            if (err) reject(err)
            else     resolve(image)
        })
    })

    /*  generate slides  */
    let pngs = []
    let n = 1
    let slide  = null
    let prefix = null
    await bluebird.each(layersSlides.reverse(), async (item) => {
        let p = item.path.replace(/^(.+\/)[^\/]+$/, "$1")
        if (prefix !== p) {
            prefix = p
            slide = empty.clone()
            verbose(`generating image: ${chalk.blue(item.path)} (scratch)`)
        }
        else
            verbose(`generating image: ${chalk.blue(item.path)} (merged)`)
        let src = await Jimp.read(item.file)
        await src.opacity(item.opacity / 255)
        await slide.blit(src, 0, 0)
        let out = `${tmpdir.name}/slide-${n++}.png`
        await slide.writeAsync(out)
        pngs.push({ file: out, path: item.path })
    })

    /*  generate PPTX out of PNG images  */
    let pptx = new PPTXGenJS()
    pptx.setLayout({ name: "Custom", width: 10, height: 10 * (h/w) })
    pptx.defineSlideMaster({
         title: "psd2pptx",
         bkgd:  "FFFFFF",
         objects: [
             { "image": { x: 0, y: 0, w: 10, h: 10 * (h/w), path: canvasFilename } }
         ]
    })
    for (let i = 0; i < pngs.length; i++) {
        verbose(`generating slide: ${chalk.blue(pngs[i].path)}`)
        let slide = pptx.addNewSlide("psd2pptx")
        slide.addImage({ path: pngs[i].file, x: 0, y: 0, w: 10, h: 10 * (h/w) })
    }
    let pptxfile = `${tmpdir.name}/slides.pptx`
    verbose("generating PPTX")
    await new Promise((resolve, reject) => {
        pptx.save(pptxfile, (filename) => {
            resolve()
        })
    })

    /*  post-adjust PPTX: add slide transition  */
    verbose("post-adjusting PPTX")
    let zip = fs.readFileSync(pptxfile)
    let out = await zipProcess(zip, {
        compression: "DEFLATE",
        extendOptions: { compressionOptions: { level: 9 } }
    }, {
        string: {
            filter: (relativePath, fileInfo) => {
                return relativePath.match(/^ppt\/slides\/slide\d+\.xml$/)
            },
            callback: (data, relativePath, zipObject) => {
                let transition =
                    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
                        '<mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">' +
                            '<p:transition spd="med" p14:dur="700">' +
                                '<p:fade/>' +
                            '</p:transition>' +
                        '</mc:Choice>' +
                        '<mc:Fallback>' +
                            '<p:transition spd="med">' +
                                '<p:fade/>' +
                            '</p:transition>' +
                        '</mc:Fallback>' +
                    '</mc:AlternateContent>'
                data = data.replace(/(<\/p:sld>)/, transition + "$1")
                return data
            }
        }
    })

    /*  write output file  */
    let pptxname = (argv.output !== "" ? argv.output : `${basename}.pptx`)
    verbose(`writing PPTX file: ${chalk.blue(pptxname)}`)
    fs.writeFileSync(pptxname, out)

    /*  delete temporary filesystem area  */
    tmpdir.removeCallback()

})().catch((err) => {
    /*  report error  */
    process.stderr.write(chalk.red(`** ERROR: ${err.stack}\n`))
    process.exit(1)
})

