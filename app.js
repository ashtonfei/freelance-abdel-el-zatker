const APP_NAME = "Signage App"

const SN_MARGINS = "Margins"  // sheet name of margins
const SN_IMPORT = "Import Demo" // sheet name of import (ouput)

const CN_DESIGN_NAME = "B"  // column Name of design name
const CN_SIZE = "D" // column Name of szie
const CN_SINGLE = "AL" // column Name of signle color
const CN_MULTI = "AU" // column Name of multi color


const ROW_START = 5 // row starts (margins)

const CUT_TO_SHAPE = "Cut To Shape"
const FULL_BOARD = "Full Board"
const SINGLE_COLOR = "Single Colour"
const MULTI_COLOR = "Multi-Colour"
const NEON_SIGN = "Neon Sign"
const SEP_HANDLE = "-"
const ACTIVE = "ACTIVE"

const CN_HANDLE = "A" // column name of handle
const CN_TITLE = "B" // column name of title
const CN_OPTION_1 = "I" // column name of option 1
const CN_OPTION_2 = "K" // column name of option 2
const CN_OPTION_3 = "M" // column name of option 3

const CN_OPTION_NAME_1 = "H" // column name of option name 1
const CN_OPTION_NAME_2 = "J" // column name of option name 2
const CN_OPTION_NAME_3 = "L" // column name of option name 3

const CN_PRICE = "T" // column name of option 1 variant price
const CN_STATUS = "AV" // column name of status

const CN_VENDOR_NAME = "D"
const CN_TYPE = "E"
const CN_POLICY = "R"
const CN_SERVICE = "S"

const OPTION_NAME_1 = "Size (Length of Longest Size)"
const OPTION_NAME_2 = "Backing Style"
const OPTION_NAME_3 = "Colour Type"
const VENDOR_NAME = "Neon Rooms"
const TYPE = "Neon Signs"
const POLICY = "continue"
const SERVICE = "manual"


// headers for the import sheet
const HEADERS = [ 
  'Handle',
  'Title',
  'Body (HTML)',
  'Vendor',
  'Type',
  'Tags',
  'Published',
  'Option1 Name',
  'Option1 Value',
  'Option2 Name',
  'Option2 Value',
  'Option3 Name',
  'Option3 Value',
  'Variant SKU',
  'Variant Grams',
  'Variant Inventory Tracker',
  'Variant Inventory Qty',
  'Variant Inventory Policy',
  'Variant Fulfillment Service',
  'Variant Price',
  'Variant Compare At Price',
  'Variant Requires Shipping',
  'Variant Taxable',
  'Variant Barcode',
  'Image Src',
  'Image Position',
  'Image Alt Text',
  'Gift Card',
  'SEO Title',
  'SEO Description',
  'Google Shopping / Google Product Category',
  'Google Shopping / Gender',
  'Google Shopping / Age Group',
  'Google Shopping / MPN',
  'Google Shopping / AdWords Grouping',
  'Google Shopping / AdWords Labels',
  'Google Shopping / Condition',
  'Google Shopping / Custom Product',
  'Google Shopping / Custom Label 0',
  'Google Shopping / Custom Label 1',
  'Google Shopping / Custom Label 2',
  'Google Shopping / Custom Label 3',
  'Google Shopping / Custom Label 4',
  'Variant Image',
  'Variant Weight Unit',
  'Variant Tax Code',
  'Cost per item',
  'Status' 
]

class App {
  constructor() {
    this.ss = SpreadsheetApp.getActive()
    this.designNameIndex = this.getColumnIndex(CN_DESIGN_NAME)
    this.sizeIndex = this.getColumnIndex(CN_SIZE)
    this.singleIndex = this.getColumnIndex(CN_SINGLE)
    this.multiIndex = this.getColumnIndex(CN_MULTI)
    
    this.handleIndex = this.getColumnIndex(CN_HANDLE)
    this.titleIndex = this.getColumnIndex(CN_TITLE)
    this.option1Index = this.getColumnIndex(CN_OPTION_1)
    this.option2Index = this.getColumnIndex(CN_OPTION_2)
    this.option3Index = this.getColumnIndex(CN_OPTION_3)
    this.optionName1Index = this.getColumnIndex(CN_OPTION_NAME_1)
    this.optionName2Index = this.getColumnIndex(CN_OPTION_NAME_2)
    this.optionName3Index = this.getColumnIndex(CN_OPTION_NAME_3)
    this.priceIndex = this.getColumnIndex(CN_PRICE)
    this.statusIndex = this.getColumnIndex(CN_STATUS)
    
    this.vendorNameIndex = this.getColumnIndex(CN_VENDOR_NAME)
    this.typeIndex = this.getColumnIndex(CN_TYPE)
    this.policyIndex = this.getColumnIndex(CN_POLICY)
    this.serviceIndex = this.getColumnIndex(CN_SERVICE)
  }
  
  getColumnIndex(name) {
    return this.ss.getActiveSheet().getRange(`${name.trim().toUpperCase()}1`).getColumn()
  }
  
  getMargins() {
    const ws = this.ss.getSheetByName(SN_MARGINS)
    if (!ws) return null
    const dataRange = ws.getDataRange()
    const values = dataRange.getValues()
    const margins = []
    let designName
    values.slice(ROW_START - 1).forEach(value => {
      designName = value[this.designNameIndex - 1].toString().trim() || designName
      const handle = designName.toLowerCase().replace(/\s/g, SEP_HANDLE)
      const title = `'${designName}' ${NEON_SIGN}`
      const size = value[this.sizeIndex - 1].toString().trim()
      const single = value[this.singleIndex - 1]
      const multi = value[this.multiIndex - 1]
      if (size !== ""){
        const margin = HEADERS.map(h => null)
        
        margin[this.handleIndex - 1] = handle
        margin[this.titleIndex - 1] = title
        margin[this.option1Index - 1] = size
        margin[this.statusIndex - 1] = ACTIVE
        margin[this.optionName1Index - 1] = OPTION_NAME_1
        margin[this.optionName2Index - 1] = OPTION_NAME_2
        margin[this.optionName3Index - 1] = OPTION_NAME_3
        
        margin[this.vendorNameIndex - 1] = VENDOR_NAME
        margin[this.typeIndex - 1] = TYPE
        margin[this.policyIndex - 1] = POLICY
        margin[this.serviceIndex - 1] = SERVICE
        
        const margin1 = [...margin]
        const margin2 = [...margin]
        const margin3 = [...margin]
        const margin4 = [...margin]
        
        margin1[this.option2Index - 1] = CUT_TO_SHAPE
        margin1[this.option3Index - 1] = SINGLE_COLOR
        margin1[this.priceIndex - 1] = single
        margins.push(margin1)
        
        margin2[this.option2Index - 1] = CUT_TO_SHAPE
        margin2[this.option3Index - 1] = MULTI_COLOR
        margin2[this.priceIndex - 1] = multi
        margins.push(margin2)
        
        margin3[this.option2Index - 1] = FULL_BOARD
        margin3[this.option3Index - 1] = SINGLE_COLOR
        margin3[this.priceIndex - 1] = single
        margins.push(margin3)
        
        margin4[this.option2Index - 1] = FULL_BOARD
        margin4[this.option3Index - 1] = MULTI_COLOR
        margin4[this.priceIndex - 1] = multi
        margins.push(margin4)
      }
    })
    margins.unshift(HEADERS)
    return margins
  }
  
  writeToSheet(values) {
    const ws = this.ss.getSheetByName(SN_IMPORT) || this.ss.insertSheet(SN_IMPORT)
    ws.clear()
    ws.getRange(1, 1, values.length, values[0].length).setValues(values)
    ws.activate()
  }
  
  run() {
    const margins = this.getMargins()
    this.writeToSheet(margins)
  }
}

function run() {
  const ss = SpreadsheetApp.getActive()
  ss.toast("Running...", APP_NAME)
  const startTime = new Date().getTime()
  try {
    const app = new App()
    app.run()
    const endTime = new Date().getTime()
    const usedTime = Math.floor((endTime - startTime) / 1000)
    ss.toast(`Done. Used time in seconds ${usedTime}.`, APP_NAME)
  } catch(e) {
    ss.toast(e.message, APP_NAME, 30)
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu(APP_NAME)
  menu.addItem("Run", "run")
  .addToUi()
}