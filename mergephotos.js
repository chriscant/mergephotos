// node mergephotos.js data/config.json
//
// Find photo filenames and add into matching specieslist spreadsheet rows
//
// CumbriaBotany/mergephotos
// https://www.npmjs.com/package/exceljs#reading-xlsx

import fs from 'fs'
import path from 'path'
import ExcelJS from 'ExcelJS'
import { fileURLToPath } from 'url'
const __dirname = path.dirname(fileURLToPath(import.meta.url))

let config = false
let outsheet = false

const COL_MIN = 0
const COL_NAME = 0
const COL_FULL_NAME = 1
const COL_PHOTOS = 6
const COL_MAX = 15

const packageJson = fs.readFileSync('./package.json')
const version = JSON.parse(packageJson).version || 0
const now = new Date()
const fullversion = 'mergephotos ' + version + ' - run at ' + now.toLocaleString()

function exiterror (r, ...msg) {
  console.log(r, ...msg)
  return 0
}

export async function run (argv) {
  const rv = 1
  try {
    // Display usage
    if (argv.length <= 2) {
      console.error('usage: node mergephotos.js <config.json>')
      return 0
    }
    console.log(fullversion)

    // Load config file and remove UTF-8 BOF and any comments starting with //
    let configtext = fs.readFileSync(path.resolve(__dirname, argv[2]), { encoding: 'utf8' })
    if (configtext.charCodeAt(0) === 65279) { // Remove UTF-8 start character
      configtext = configtext.slice(1)
    }
    while (true) {
      const dslashpos = configtext.indexOf('//')
      if (dslashpos === -1) break
      const endlinepos = configtext.indexOf('\n', dslashpos)
      if (endlinepos === -1) {
        configtext = configtext.substring(0, dslashpos)
        break
      }
      configtext = configtext.substring(0, dslashpos) + configtext.substring(endlinepos)
    }
    // console.log(configtext)
    try {
      config = JSON.parse(configtext)
    } catch (e) {
      console.error('config file not in JSON format')
      return 0
    }
    console.log(config)

    if (!('input' in config)) {
      console.error('config.input missing')
      return 0
    }
    if (!('output' in config)) {
      console.error('config.output missing')
      return 0
    }

    /// /
    const outbook = new ExcelJS.Workbook()
    outsheet = outbook.addWorksheet('FoC')
    outsheet.columns = [
      { header: 'Name', key: 'Name', width: 20 }, // COL_NAME
      { header: 'Full name', key: 'Full name', width: 20 }, // COL_FULL_NAME
      { header: 'Full name 2', key: 'Full name 2', width: 20 }, // COL_FULL_NAME_2
      { header: 'English name', key: 'English name', width: 20 }, // COL_ENGLISH_NAME
      { header: 'Modern text', key: 'Modern text', width: 30 }, // COL_MODERN_TEXT
      { header: 'Flora 1997 text', key: 'Flora 1997 text', width: 80 }, // COL_1997_TEXT
      { header: 'SpeciesPhoto1', key: 'SpeciesPhoto1', width: 20 },
      { header: 'SpeciesPhoto1Caption', key: 'SpeciesPhoto1Caption', width: 20 },
      { header: 'SpeciesPhoto2', key: 'SpeciesPhoto2', width: 20 },
      { header: 'SpeciesPhoto2Caption', key: 'SpeciesPhoto2Caption', width: 20 },
      { header: 'SpeciesPhoto3', key: 'SpeciesPhoto3', width: 20 },
      { header: 'SpeciesPhoto3Caption', key: 'SpeciesPhoto3Caption', width: 20 },
      { header: 'SpeciesPhoto4', key: 'SpeciesPhoto4', width: 20 },
      { header: 'SpeciesPhoto4Caption', key: 'SpeciesPhoto4Caption', width: 20 },
      { header: 'SpeciesPhoto5', key: 'SpeciesPhoto5', width: 20 },
      { header: 'SpeciesPhoto5Caption', key: 'SpeciesPhoto5Caption', width: 20 }
    ]

    /// /
    const inbook = new ExcelJS.Workbook()
    await inbook.xlsx.readFile(config.input)
    const insheet = inbook.worksheets[0]

    console.log('rowCount', insheet.rowCount)
    console.log('actualRowCount', insheet.actualRowCount)
    console.log('columnCount', insheet.columnCount)
    console.log('actualColumnCount', insheet.actualColumnCount)
    console.log()

    const species = []
    let photoSpeciesNotFound = 0
    let photosNotJPG = 0
    let maxPhotoColumn = 0

    for (let r = 2; r <= insheet.rowCount; r++) {
      // console.log('=== ROW ', r)
      const inrow = insheet.getRow(r)

      const outline = []
      for (let colno = 0; colno < insheet.columnCount; colno++) {
        const inrowcell = inrow.getCell(colno + 1)
        if (typeof inrowcell !== 'object') return exiterror(r, 'inrowcell unexpectedly not an object')
        outline.push(inrowcell.value)
        // console.log(typeof inrowcell.value)
      }
      if (outline.length < COL_PHOTOS) { // Ensure outline at least COL_PHOTOS long
        outline.push(...Array(COL_PHOTOS - outline.length).fill(''))
      }
      species.push(outline)
      // if (r === 6) break
    }

    const rootdirs = fs.readdirSync(config.photosdir, { withFileTypes: true })
    // const rootdirs = await glob(, { withFileTypes: true })
    for (const rootdir of rootdirs) {
      if (rootdir.isDirectory()) {
        // console.log(rootdir)
        const rootpath = path.resolve(config.photosdir, rootdir.name)
        const speciedirs = fs.readdirSync(rootpath, { withFileTypes: true })
        for (const speciedir of speciedirs) {
          if (speciedir.isDirectory()) {
            // Fix up speciesname
            let speciesname = speciedir.name
            const speciesnamelen = speciesname.length
            let undertodot = false
            let adddot = false
            const last3 = speciesname.substring(speciesnamelen - 3)
            const last4 = speciesname.substring(speciesnamelen - 4)
            if (last4 === 'agg_') undertodot = true
            if (last4 === 's.s_') undertodot = true
            if (last4 === 's.l_') undertodot = true
            if (last3 === 's.s') adddot = true
            if (undertodot) {
              speciesname = speciesname.substring(0, speciesnamelen - 1) + '.'
            }
            if (adddot) {
              speciesname += '.'
            }
            // Look for specie in spreadsheet
            let found = species.find(row => row[0] === speciesname)
            if (!found) {
              found = species.find(row => row[0] === speciesname + ' agg.')
              if (found) speciesname += ' agg.'
            }
            if (!found) {
              console.log(speciedir.name + ' NOT FOUND')
              photoSpeciesNotFound++
              // console.log('NOT FOUND')
            } else {
              // console.log(speciedir.name + ' FOUND')
              const speciepath = path.resolve(rootpath, speciedir.name)
              const photofiles = fs.readdirSync(speciepath, { withFileTypes: true })
              let insertpos = COL_PHOTOS
              for (const photofile of photofiles) {
                if (photofile.isFile()) {
                  // console.log(photofile.name)
                  let caption = photofile.name
                  const captionlength = caption.length
                  if (caption.substring(captionlength - 4).toLowerCase() === '.jpg') {
                    caption = caption.substring(0, captionlength - 4)
                  } else if (caption.substring(captionlength - 5).toLowerCase() === '.jpeg') {
                    caption = caption.substring(0, captionlength - 5)
                  } else {
                    console.log(photofile.name, 'NOT JPG/JPEG')
                    photosNotJPG++
                  }
                  const path = rootdir.name + '/' + speciedir.name + '/' + photofile.name
                  found[insertpos++] = path.replaceAll(' × ', ' x ')
                  found[insertpos++] = caption
                }
              }
              if (insertpos > maxPhotoColumn) maxPhotoColumn = insertpos
              // break
            }
          }
        }
        // break
      }
    }

    // Add updated rows to output sheet
    for (const specie of species) {
      outsheet.addRow(specie)
    }

    // Set alignment and wrap for all cells. Fix up multiplication signs
    for (let outrow = 1; outrow <= outsheet.rowCount; outrow++) {
      for (let col = COL_MIN; col < COL_MAX; col++) {
        const cell = outsheet.getRow(outrow).getCell(col + 1)
        cell.alignment = { vertical: 'top', wrapText: true }
        let cellv = cell.value
        if (cellv && typeof cellv === 'string') {
          if (col === COL_NAME) { // Name col: ensure all multiplication sign × are letter x
            cellv = cellv.replaceAll(' × ', ' x ')
          } else if (col >= COL_FULL_NAME && col < COL_PHOTOS) { // Other main cols: replace letter x with multiplication sign ×
            cellv = cellv.replaceAll(' x ', ' × ')
          }
          cellv = cellv.replaceAll(' </i> ', '</i> ')
          cell.value = cellv
        }
      }
    }
    outsheet.getRow(1).font = { bold: true }
    outsheet.views = [{ state: 'frozen', ySplit: 1 }]
    await outbook.xlsx.writeFile(config.output)

    console.log()
    console.log('photoSpeciesNotFound', photoSpeciesNotFound)
    console.log('photosNotJPG', photosNotJPG)
    console.log('maxPhotoColumn', maxPhotoColumn)

    if (rv) console.log('SUCCESS')
    return 1
  } catch (e) {
    console.error('run EXCEPTION', e)
    return 2
  }
}

/// ////////////////////////////////////////////////////////////////////////////////////
// If called from command line, then run now.
// If jest testing, then don't.
if (process.env.JEST_WORKER_ID === undefined) {
  run(process.argv)
}
