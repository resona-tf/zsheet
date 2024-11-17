/*
 * Version: 1.1.1
 * Update: プログレスバーの制御機能を追加
 */

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

// プログレスバーの制御用変数
let currentProgress = 0
let totalRecords = 0

function initProgress(total) {
    currentProgress = 0
    totalRecords = total
    const progressBar = document.getElementById('progressBar')
    progressBar.style.width = '0%'
    progressBar.setAttribute('aria-valuenow', '0')
}

function progressNext() {
    currentProgress++
    const percentage = (currentProgress / totalRecords) * 100
    const progressBar = document.getElementById('progressBar')
    progressBar.style.width = percentage + '%'
    progressBar.setAttribute('aria-valuenow', percentage)
}

//###############　負荷テスト用コード　################
let debugRecordId = ""
//###############　負荷テスト用コード　################


const PRDUCTION_ORGID = "90001619930"
const TEMPLATE_CRMVAR = new URLSearchParams(window.location.search).get('varName')
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
