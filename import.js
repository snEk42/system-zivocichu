const Promise = require('bluebird')
const Database = require('better-sqlite3')
const xlsx = require('xlsx')
const axios = require('axios')
const _ = require('lodash')

const TAXONOMIC_UNITS = {
  Class: 'Třída',
  Order: 'Řád',
  Family: 'Čeleď',
}

const db = new Database('./ITIS.sqlite')

let unmatchedCount = 0

importClassAndFamily()

async function importClassAndFamily() {
  const { workbook, sheets } = await parseFile('./data.xlsx')
  const finalSheets = await Promise.mapSeries(sheets, processSheet)
  await writeFile(workbook, finalSheets, 'filled.xlsx')
  console.log('UNMATCHED: ' + unmatchedCount)
}

async function processSheet(sheet) {
  const rows = _.tail(sheet.data)
  const genusChunks = groupByGenus(rows)
  const filledChunks = await Promise.mapSeries(_.values(genusChunks), processGenusChunk)
  const finalData = [
    _.concat(_.head(sheet.data), _.values(TAXONOMIC_UNITS)),
    ..._.flatten(filledChunks),
  ]
  return {
    data: finalData,
    name: sheet.name,
  }
}

function groupByGenus(rows) {
  const capitalizedAndTrimmed = rows.map(row => {
    row[0] = _.trim(row[0])
    row[1] = _.upperFirst(_.trim(row[1]))
    return row
  })
  const rowsWithGenus = capitalizedAndTrimmed.map(row => {
    const latinNames = _.split(row[1], ' ')
    return {
      genus: _.head(latinNames),
      species: _.tail(latinNames).join(' '),
      sourceRow: row,
    }
  })
  const ordered = _.orderBy(rowsWithGenus, ['genus', 'species'], ['asc', 'asc'])
  return _.groupBy(ordered, 'genus')
}

async function processGenusChunk(chunk) {
  if (!chunk || chunk.length === 0) {
    return []
  }
  if (!chunk[0].genus) {
    console.log('MISSING GENUS: ' + _.map(chunk, 'sourceRow[0]').join(', '))
    return _.map(chunk, 'sourceRow')
  }
  const possibleAnimals = db.prepare(`
    SELECT *
    FROM longnames
    WHERE completename LIKE '${chunk[0].genus}%'
    ORDER BY completename
  `).all()
  const filledChunkRows = await Promise.map(chunk, row => addTaxonomicUnits(possibleAnimals, row))
  return filledChunkRows
}

function addTaxonomicUnits(possibleAnimals, row) {
  if (!possibleAnimals || possibleAnimals.length === 0) {
    console.log('NO MATCH: ' + row.sourceRow[1])
    unmatchedCount++
    return row.sourceRow
  }

  // Handle full match
  const fullMatch = _.find(possibleAnimals, { completename: `${row.genus} ${row.species}` })
  if (fullMatch) {
    const taxonomicUnits = getTaxonomicUnitsByTSN(fullMatch.tsn)
    return _.concat(row.sourceRow, selectWanted(taxonomicUnits))
  }

  // Handle genus name match
  let animalIndex = 0
  let taxonomicUnits = []
  while (!_.find(taxonomicUnits, { name: 'Animalia', type: 'Kingdom' }) && possibleAnimals.length > animalIndex) {
    taxonomicUnits = getTaxonomicUnitsByTSN(possibleAnimals[animalIndex].tsn)
    animalIndex++
  }
  if (!_.find(taxonomicUnits, { name: 'Animalia', type: 'Kingdom' })) {
    console.log('NO ANIMALIA MATCH: ' + row.sourceRow[1])
    return _.keys(TAXONOMIC_UNITS).map(() => '')
  }
  return _.concat(row.sourceRow, selectWanted(taxonomicUnits))
}

function getTaxonomicUnitsByTSN(tsn) {
  let hierarchy = db.prepare('SELECT * FROM hierarchy WHERE tsn=?').get(tsn)
  if (!hierarchy) {
    const synonym = db.prepare('SELECT * FROM synonym_links WHERE tsn=?').get(tsn)
    hierarchy = db.prepare('SELECT * FROM hierarchy WHERE tsn=?').get(synonym.tsn_accepted)
    if (!hierarchy) {
      console.log('UNMATCHED HIERARCHY: ' + tsn)
      throw new Error('Unmatched hierarchy')
    }
  }
  const parentTSNs = _.split(hierarchy.hierarchy_string, '-')
  const parentUnits = db.prepare(`
    SELECT
      units.unit_name1 AS name,
      types.rank_name AS type
    FROM taxonomic_units AS units
      INNER JOIN taxon_unit_types AS types ON units.rank_id = types.rank_id AND types.kingdom_id = 5
    WHERE units.tsn IN (${parentTSNs.map(tsn => `'${tsn}'`).join(', ')})
  `).all()
  return parentUnits
}

function selectWanted(taxonomicUnits) {
  return _.keys(TAXONOMIC_UNITS).map(type => {
    const taxonomicUnit = _.find(taxonomicUnits, { type })
    return taxonomicUnit ? taxonomicUnit.name : ''
  })
}

async function parseFile(fileName) {
  const workbook = await xlsx.readFile(fileName)
  const sheets = await Promise.mapSeries(workbook.SheetNames, async sheetName => {
    return {
      data: await xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 }),
      name: sheetName,
    }
  })
  return { workbook, sheets }
}

async function writeFile(workbook, sheets, fileName) {
  await Promise.mapSeries(sheets, async sheet => {
    workbook.Sheets[sheet.name] = await xlsx.utils.aoa_to_sheet(sheet.data)
  })
  await xlsx.writeFile(workbook, fileName)
}


