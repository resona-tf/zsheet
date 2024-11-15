let ENTITY
let ENTITY_ID
let FIELDS
let RECORD
let MODULE
let XLSXSELECTOR
let DBG = false
const workbook = new ExcelJS.Workbook()

window.onload = async function(){
// debugger
	XLSXSELECTOR = document.getElementById("xlsxSelector")

	ZOHO.embeddedApp.on("PageLoad", async function(data){

		ZOHO.CRM.UI.Resize({height:"300", width:"540"})

		ENTITY = data.Entity
		ENTITY_ID = data.EntityId[0]

		FIELDS = await ZOHO.CRM.META.getFields({Entity:ENTITY})

		//レコード情報を取得
		RECORD = await ZOHO.CRM.API.getRecord({Entity:ENTITY, RecordID:ENTITY_ID})
		console.log(RECORD)
		//debugger

		//Excelテンプレートの全レコードを取得
		let Templates = await ZOHO.CRM.API.getAllRecords({Entity:"Excel_Templates"})
		console.log(Templates)
		
		//ENTITYのモジュール情報を取得
		let modules = await ZOHO.CRM.META.getModules()
		for(let idx in modules.modules){
			if(modules.modules[idx].api_name == ENTITY){
				MODULE = modules.modules[idx]
				break
			}
		}

		//Excelテンプレートの利用タブ名にENTITYモジュールの名前が含まれていれば、OPTIONに追加する
		for(let idx1 in Templates.data){
			let tabnames = Templates.data[idx1].Target_TabName
			let tabs = tabnames.split(",")
			for(let idx2 in tabs){
				if(MODULE.singular_label == tabs[idx2]){
					let op = document.createElement("option")
					op.value = Templates.data[idx1].id
					op.innerHTML = Templates.data[idx1].Name
					XLSXSELECTOR.appendChild(op)
				}
			}
		}
		let genBtn = document.getElementById("generateBtn")
		genBtn.addEventListener("click", function(){ generateXLSX() })
	})

	ZOHO.embeddedApp.init();

	async function generateXLSX(){

		// debugger
		//let a = await replaceZohoFieldVariables("置き換えテスト<Z%サブフォーム 1%Z>と<Z%ギャラ件名%Z>")
		//return

		//テンプレートExcelファイルを取得
		let template = await ZOHO.CRM.API.getRecord({Entity:"Excel_Templates", RecordID:XLSXSELECTOR.value})
		let xlsxFileId = template.data[0].XLSX[0].file_Id
		let xlsx = await ZOHO.CRM.API.getFile({id : xlsxFileId})
		let xlsxData = await xlsx.arrayBuffer()
		
		//ファイル名の置換変数を置き換え
		let r = await replaceZohoFieldVariables(template.data[0].XLSX[0].file_Name)
		let xlsxFileName = r[0].replace("/","-")

		//Excelファイルの内容を取得
		debugger
		await workbook.xlsx.load(xlsxData)
		workbook.removeWorksheet(1)
		let sheetToClone = workbook.getWorksheet(2)
		let copySheet = workbook.addWorksheet("Sheet")
		copySheet.model = Object.assign(sheetToClone.model, { mergeCells: sheetToClone.model.merges })
		// copySheet.name = "THE REAL ID/NAME YOU WANT FOR THIS SHEET"

		// debugger





		//Excel内のセル内容を取得
		let worksheetData = {}
		let promises = new Promise(function(resolve, reject){
			workbook.eachSheet(function(ws,wsid){
				//debugger
				worksheetData[ws.name] = {}
				ws.eachRow(function(r,rn){
					worksheetData[ws.name][rn] = {}
					r.eachCell(function(c,cn){
						worksheetData[ws.name][rn][cn] = c.value
					})
				})
			})
			resolve();
		})
		Promise.all([promises])


		for(let sheetname in worksheetData){
			for(let rn in worksheetData[sheetname]){
				for(let cn in worksheetData[sheetname][rn]){
					//debugger
					let ws = workbook.getWorksheet(sheetname)
					let origCellValue = worksheetData[sheetname][rn][cn]
					if(typeof origCellValue != "string"){ continue }
					if(origCellValue.match(/<Z%[^%]*%Z>/g)){
						// debugger
						let replace = await replaceZohoFieldVariables(origCellValue)
						for(let idx in replace){
							let row = ws.getRow(Number(rn) + Number(idx))
							let cell = row.getCell(Number(cn))
							cell.value = replace[idx]
						}
					}
				}
			}
		}


		//置き換え済みのExcelをBlob化
		let newXlsx = await workbook.xlsx.writeBuffer()
		let newXlsxData = new Blob([newXlsx])

		//Blob化したExcelファイルをレコードに添付する
		let res = await ZOHO.CRM.API.attachFile({
			Entity: ENTITY,
			RecordID : ENTITY_ID,
			File:{Name:xlsxFileName, Content:newXlsxData}
		})

		ZOHO.CRM.UI.Popup.closeReload()
	}
	async function replaceZohoFieldVariables(srcText){
		if(typeof srcText === "object"){ return srcText }

		let returnText = [srcText.replace("<Z%実行日時%Z>", moment().format("YYYYMMDD-hhmm"))]
	
		let match = returnText[0].match(/<Z%[^%]*%Z>/g)
		for(idx1 in match){
			let searchText = match[idx1]
			let fieldLabel = match[idx1].replace("<Z%","").replace("%Z>","").split("~")
			let replaceText

			let field = FIELDS.fields.find(
				function(data){ return (data.field_label == fieldLabel[0] || data.api_name == fieldLabel[0]) ? data : false }
			)
			if(field){
				switch(field.json_type){
					default:
						replaceText = RECORD.data[0][field.api_name]
						returnText[0] = returnText[0].replace(searchText, replaceText)   
						break
		
					case "jsonarray":
						

						let lookupModuleFieldApiName = ""
						let lookupModuleApiName = ""
						let lookupModuleRecord

						if(fieldLabel.length > 1){
							lookupModuleApiName = field.api_name
							let lookupModuleFields = await ZOHO.CRM.META.getFields({Entity:lookupModuleApiName})
							let lookupModuleField = lookupModuleFields.fields.find(
								function(data){ return (data.field_label == fieldLabel[1] || data.api_name == fieldLabel[1]) ? data : false }
							)
							if(lookupModuleField){
								lookupModuleFieldApiName = lookupModuleField.api_name
							}
						}

						let arrayField = RECORD.data[0][field.api_name]
						

						//請求書・発注書・見積書タブの特例措置
						let fieldProperty = "field"
						if(field.api_name == "Product_Details"){
							fieldProperty = "product"
						}
						
						for(let idx2 in arrayField){
							if(lookupModuleFieldApiName == ""){
								replaceText = arrayField[idx2][fieldProperty].name
							}else{
								replaceText = arrayField[idx2][lookupModuleFieldApiName]
							}


							if(idx2 == 0){
								returnText[0] = returnText[0].replace(searchText, replaceText)
							}else{
								returnText.push(replaceText)
							}
						}
						break
		
					case "jsonobject":
						replaceText = RECORD.data[0][field.api_name].name
						returnText[0] = returnText[0].replace(searchText, replaceText)  
						break
				}
			}


		}
		//debugger
		return returnText
	}

}