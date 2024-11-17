let ENTITY
let ENTITY_IDS
let FIELDS = []
let RELATED_LISTS = []
let RECORD
let MODULE
let Z
let ZS
let TEMPLATE_SELECTOR
let DBG = false
let workbook

let zSheetTemplate
let zSingleTemplateSheetId
let zMultiTemplateSheetId
let zSingleTemplateContents
let zMultiTemplateContents
let zCurrentTemplateContents
let worksheetContents = []

let gather

let widgetData

let progress_steps = 100
let current_steps = 0
let perCellProgressStep, perSheetProgressStep, perRecordProgressStep	

let progressTotal = 100
let progressSteps = 0
let progressCurrent = 0
function progressAddStep(step){
	progressSteps = progressSteps + step
	progressGetCurrent()
}
function progressNext(step){
	progressCurrent++
	progressGetCurrent()
}
function progressGetCurrent(){
	let current = (progressTotal/progressSteps) * progressCurrent + "%"
	document.getElementById("progressBar").style.width = current
	console.log(current)
}

//###############　負荷テスト用コード　################
let debugRecordId = ""
//###############　負荷テスト用コード　################


const PRDUCTION_ORGID = "90001619930"
const TEMPLATE_CRMVAR = "Zs_Template"
let ENVIROMENT = "production"
let ApiDomain = "https://www.zohoapis.jp"
let fileNameAddition = ""

let WORKING_BOOK_ID

let GatherSalesNumbers = ""

let API_COUNT = {}

let IP

// fetch('https://ipinfo.io?callback')
// .then(res => res.json())
// .then(json => IP = json.ip)