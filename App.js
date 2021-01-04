const APP_NAME = "Invoice App"
const SN_ORDERS = "ORDERS"
const SN_INVOICE = "INVOICE"
const ID_PDF_FOLDER = "11aTFa9Ol0HsbdnCg66oy4HlVARNTLznP" // replace this folder id with yours
const FN_PDF_INVOICES = "Invoice PDFs" // default folder name of the folder

const STATUS_SUCCESS = "PDF GENERATED"
const HEADER_STATUS = "BEMERKUNGEN"
const HEADER_ORDER_ID = "order_id"

const ORDER_ID_RANGE_NAME = "B23"

class App{
  constructor(){
    this.ss = SpreadsheetApp.getActive()
    this.id = this.ss.getId()
    this.ui = SpreadsheetApp.getUi()
    try{
      this.folder = DriveApp.getFolderById(ID_PDF_FOLDER)
    }catch(e){
      this.folder = this.getInvoicesFolder()
    }
    
    this.wsInvoice = this.ss.getSheetByName(SN_INVOICE)
    this.sheetId = this.wsInvoice.getSheetId()
    this.rangeOrderId = this.wsInvoice.getRange(ORDER_ID_RANGE_NAME)
    
    this.wsOrders = this.ss.getSheetByName(SN_ORDERS)
    
    this.pdfUrl = `https://docs.google.com/spreadsheets/d/${this.id}/export?format=pdf&size=7&gridlines=false&gid=${this.sheetId}`
  }
  
  updateStatus(orders){
    const values = this.wsOrders.getDataRange().getValues()
    const headers = values.shift()
    const indexOfOrderId = headers.indexOf(HEADER_ORDER_ID)
    const indexOfStatus = headers.indexOf(HEADER_STATUS)
    values.forEach((value, row) => {
      const id = value[indexOfOrderId]
      const status = orders[id]
      if (status) {
      if (status.indexOf("http") === 0){
        this.wsOrders.getRange(row + 2, indexOfStatus + 1).setFormula(`=HYPERLINK("${status}", "${STATUS_SUCCESS}")`)
      }else{
        this.wsOrders.getRange(row + 2, indexOfStatus + 1).setValue(status)
      }
      }
    })
  }
  
  getInvoicesFolder(){
    const currentFolder = DriveApp.getFileById(this.id).getParents().next()
    const folders = currentFolder.getFoldersByName(FN_PDF_INVOICES)
    if (folders.hasNext()) return folders.next()
    return currentFolder.createFolder(FN_PDF_INVOICES)
  }
  
  createPDF(filename){
    const blob = UrlFetchApp.fetch(this.pdfUrl, {method: "GET", headers: {Authorization: "Bearer " + ScriptApp.getOAuthToken()}}).getBlob()
    blob.setName(filename)
    const pdf = this.folder.createFile(blob)
    return pdf
  }
  
  createInvoice(id){
    this.rangeOrderId.setValue(id)
    SpreadsheetApp.flush()
    const fileName = `${id}.pdf`
    try{
      const pdf = this.createPDF(fileName)
      return pdf.getUrl()
    }catch(e){
      return e.message
    }
  }
  
  getOrderIds(){
    const values = this.wsOrders.getDataRange().getValues()
    const headers = values.shift()
    const indexOfOrderId = headers.indexOf(HEADER_ORDER_ID)
    const indexOfStatus = headers.indexOf(HEADER_STATUS)
    
    let ids = []
    values.forEach(value => {
      const id = value[indexOfOrderId]
      const status = value[indexOfStatus]
      if (status !== STATUS_SUCCESS) ids.push(id)
    })
    ids = [... new Set(ids)]
    return ids
  }

  run(){
    const input = this.ui.prompt(APP_NAME, "How many invoices do you want to create?", this.ui.ButtonSet.OK_CANCEL)
    const numberOfInvoices = Number(input.getResponseText())
    const confirmButton = input.getSelectedButton()
    const startTime = new Date().getTime()
    if (numberOfInvoices && confirmButton === this.ui.Button.OK) {
      const ids = this.getOrderIds()
      const orders = ids.slice(0, numberOfInvoices)
      this.ss.toast(`${orders.length} invoices will be created.`, APP_NAME, 15)
      const results = {}
      orders.forEach(order => results[order] = this.createInvoice(order))
      this.updateStatus(results)
      const endTime = new Date().getTime()
      const usedTime = Math.floor((endTime - startTime) / 1000)
      this.ss.toast(`Done! Used time in seconds ${usedTime}.`, APP_NAME + " - Success", 15)
    }else{
      this.ss.toast(`${numberOfInvoices} invoices will be canceled.`, APP_NAME + " - Error", 15)
    }
  }
}

function run() {
  const app = new App()
  app.run()
}

function onOpen(){
  SpreadsheetApp.getUi().createMenu(APP_NAME).addItem("Create", "run").addToUi()
}