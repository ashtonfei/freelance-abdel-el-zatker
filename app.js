const APP_NAME = "PointeWest"
const SN_DATA = "Data"
const SN_DEMO = "Demo"

const STATUS_ACTIVE = "ACTIVE"
const STATUS_INACTIVE = "INACTIVE"

const HEADERS =  [ 
  'Status',
  'Listing ID',
  'Details Page - URL',
  'Address',
  'Unit',
  'City',
  'State',
  'Zip',
  'Headline 1',
  'Listing Detail Title',
  'Listing Detail Descritpion',
  'Rent Value',
  'Date Available',
  'Bedrooms Value',
  'Baths Value',
  'Image URL' 
]

function createTrigger(){
  const functionName = "run"
  const triggers = ScriptApp.getProjectTriggers()
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger))
  ScriptApp.newTrigger("run").timeBased().everyMinutes(10).create()
  SpreadsheetApp.getActive().toast(`New trigger has been created.`, APP_NAME)
}

function onOpen(){
  const menu = SpreadsheetApp.getUi().createMenu(APP_NAME)
  menu.addItem("Run", "run")
  menu.addItem("Add trigger", "createTrigger")
  menu.addToUi()
}

function run() {
  const startTime = new Date().getTime()
  const ss = SpreadsheetApp.getActive()
  try{
  ss.toast("Fetching data from website...", APP_NAME)
  const app = new App()
  app.run()
  const endTime = new Date().getTime()
  const usedTime = Math.floor((endTime - startTime) / 1000)
  ss.toast(`Done. Used time ${usedTime}s.`, APP_NAME)
//  app.getDetails("https://pointewest.appfolio.com/listings/detail/727a113a-fdc8-4d75-8919-096f90cff213")
  }catch(e){
    ss.toast(e.message, APP_NAME, 30)
  }
}

class App {
  constructor(){
    this.ss = SpreadsheetApp.getActive()
    this.ws = this.ss.getSheetByName(SN_DATA)
    this.wsDemo = this.ss.getSheetByName(SN_DEMO)
    this.baseUrl = "https://pointewest.appfolio.com"
    this.listUrl = "https://pointewest.appfolio.com/listings"
  }
  
  getHtmlContentByUrl(url){
    const response = UrlFetchApp.fetch(url)
    const code = response.getResponseCode()
    if (code !== 200) return null
    return response.getContentText()
  }
  
  findAllMatches(re, content){
    const matches = []
    let match
    while ((match = re.exec(content)) != null) {
      matches.push(match[1])
    }
    return matches
  }
  
  findMatch(re, content){
    const match = re.exec(content)
    if (!match) return null
    return match[1]
  }
  
  parseAddress(fullAddress){
    const items = fullAddress.split(",")
    let address = null
    let unit = null
    let city = null
    let state = null
    let zip = null
    if (items.length === 4) {
      address = items[0].trim()
      unit = items[1].trim()
      city = items[2].trim()
      let stateZip = items[3].trim().split(" ")
      state = stateZip[0]
      zip = stateZip[1]
    } else if (items.length === 3) {
      const addressValue = items[0].trim()
      if (addressValue.indexOf("Unit") !== -1){
        address = addressValue.split("Unit")[0]
        unit = "Unit" + addressValue.split("Unit")[1]
      }else if(addressValue.indexOf("#") !== -1){
        address = addressValue.split("#")[0].trim()
        unit = "#" + addressValue.split("#")[1]
      }else if (addressValue.indexOf("-") !== -1) {
        address = addressValue.split("-")[0].trim()
        unit = addressValue.split("-")[1].trim()
      }else{
        address = addressValue
      }
      city = items[1].trim()
      let stateZip = items[2].trim().split(" ")
      state = stateZip[0]
      zip = stateZip[1]
    }else{
      address = fullAddress
    }
    return {address, unit, city, state, zip}
  }
  
  getDetails(url, imageUrl){
    let content = this.getHtmlContentByUrl(url)
    if (!content) return null
    content = content.replace(/<br \/*>/gmi, "")
    const addressRegex = /js-show-title">\s*(.*)\s*<a/mgi
    const fullAddress = this.findMatch(addressRegex, content)
    const {address, unit, city, state, zip} = this.parseAddress(fullAddress)
            
    const bbRegex = /<h3 class="font-weight-normal">\s*(.*)\s*<\/h3>/mgi
    let bedAndBath = this.findMatch(bbRegex, content)
    if (bedAndBath) bedAndBath = bedAndBath.replace("bd", "").replace("ba", "")
    const [bed, bath] = bedAndBath ? bedAndBath.trim().split("/") : [null, null]
     
            
    const titleRegex = /<h2 class="listing-detail__title">(.*)<\/h2>/mgi
    const title = this.findMatch(titleRegex, content)
    
    const detailsRegex = /<p class="listing-detail__description hand-hidden tablet-hidden font-weight-light">(.*\s*)/gmi
    const details = this.findMatch(detailsRegex, content)
    
    const rentRegex = /<li class="list__item">Rent:\s*(.*)<\/li>/gmi
    const rent = this.findMatch(rentRegex, content)
            
    const availableRegex = /<li class="list__item">Available\s*(.*)<\/li>/gmi
    const available = this.findMatch(availableRegex, content)
    
    const id = url.split("/detail/")[1]
    return [STATUS_ACTIVE, id, url,address, unit, city, state, zip, '', title, details, rent, available, bed, bath, imageUrl]
  }
  
  list(){
    const urlRegex = /container">\s*<a href="(\/.*?)"/gmi
    const imageRegex = /original="(http.*?)"/gmi
    const content = this.getHtmlContentByUrl(this.listUrl)
    const urls = this.findAllMatches(urlRegex, content)
    const imageUrls = this.findAllMatches(imageRegex, content)
    const values = urls.map((url, i) => {
      const itemUrl = `${this.baseUrl}${url}`
      const details = this.getDetails(itemUrl, imageUrls[i])
      return details
    })
    return values
  }
  
  check(list){
    const dataRange = this.ws.getDataRange()
    const values = dataRange.getValues()
    values.forEach((v, i) => {
      const [status, id, url] = v
      const imageUrl = v[v.length - 1]
      if (status === STATUS_ACTIVE) {
        const match = list.find(listValue => listValue[1] === id)
        if (match) {
          values[i] = match
        }else{
          const details = this.getDetails(url, imageUrl)
          if (details) {
            values[i] = details
          } else {
            values[i][0] = STATUS_INACTIVE
          }
        }
      }
    })
    
    list.forEach(item => {
      const match = values.find(v => v[1] === item[1])
      if (!match) values.push(item)
    })
    return values
  }
  
  run(){
    const list = this.list()
    const values = this.check(list)
    values[0] = HEADERS
    this.ws.clearContents().getRange(1, 1, values.length, values[0].length).setValues(values)
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM hh:mm:ss')
    this.ws.getRange("A1").setNote(`${APP_NAME} - Last Refresh:\n${timestamp}`)
    this.ws.activate()
  }
}