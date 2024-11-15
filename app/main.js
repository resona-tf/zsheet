let ENTITY
let ENTITY_IDS
let FIELDS = []
let RELATED_LISTS = []
let RECORD
let MODULE
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


const PRDUCTION_ORGID = "90000685150"
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

window.onload = async function(){
	ZOHO.embeddedApp.on("PageLoad", async function(data){
		let orgInfo = await ZOHO.CRM.CONFIG.getOrgInfo()
		const orgId = orgInfo.org[0].zgid
		ApiDomain = orgId == PRDUCTION_ORGID ? "https://www.zohoapis.jp" : "https://crmsandbox.zoho.jp"
		CrmUrl =  orgId == PRDUCTION_ORGID ? "https://crm.zoho.jp/crm/" + orgInfo.org[0].domain_name : "https://crmsandbox.zoho.jp/crm/" + orgInfo.org[0].domain_name

		fileNameAddition = orgId == PRDUCTION_ORGID ? "" : "_Sandbox"
		widgetData = data

		await waitFor("#loadCheck")
		const UrlParams = new URLSearchParams(window.location.search)
		const mode = UrlParams.get("mode")
		const paymentDocsOnly = UrlParams.get("docsonly")

		let templateUrl



		function apiCounter(api){
			if(API_COUNT[api]){
				API_COUNT[api]++
			}else{
				API_COUNT[api] = 1
			}
			console.log(`## API ## ${api} : ${API_COUNT[api]}`)
		}
		//レコード情報キャッシュ用オブジェクト
		Z = {
			records:{},
			fields:{},
			module:{},
			search:{},
			apiCount:{},
			searchRecord:async function(entity, criteria){
				if(Z.search[entity]){
					if(Z.search[entity][criteria]){
						return Z.search[entity][criteria]
					}
				}else{
					Z.search[entity] = {}
				}
				apiCounter("ZOHO.CRM.API.searchRecords")
				// console.log(`ZOHO.CRM.API.searchRecords({Entity:${entity}, Criteria:${criteria}}})`)
				let res = await ZOHO.CRM.API.searchRecord({Entity:entity, Type:"criteria", Query:criteria})
				Z.search[entity][criteria] = res.data

				for(let record of res.data){
					if(!Z.records[entity]){ Z.records[entity] = {} }
					if(!Z.records[entity][record.id]){ Z.records[entity][record.id] = {} }
					Z.records[entity][record.id].data = record
				}
				return res.data
			},
			getRecord:async function(entity, id){
				if(Z.records[entity]){
					if(Z.records[entity][id]){
						if(Z.records[entity][id].data){ return Z.records[entity][id].data }
					}else{
						Z.records[entity][id] = {}
					}
				}else{
					Z.records[entity] = {}
					Z.records[entity][id] = {}
				}
				apiCounter("ZOHO.CRM.API.getRecord")
				//console.log(`ZOHO.CRM.API.getRecord({Entity:${entity}, RecordID:${id}}})`)
				let res = await ZOHO.CRM.API.getRecord({Entity:entity, RecordID:id})
				let record = res.data[0]
				Z.records[entity][id].data = record
				return record
			},
			getRelatedRecords:async function(entity, id, related){	
				if(Z.records[entity]){
					if(Z.records[entity][id]){
						if(Z.records[entity][id].related){
							if(Z.records[entity][id].related[related]){ return Z.records[entity][id].related[related] }
						}else{
							Z.records[entity][id].related = {}
						}
					}else{
						Z.records[entity][id] = {}
						Z.records[entity][id].related = {}
					}
				}else{
					Z.records[entity] = {}
					Z.records[entity][id] = {}
					Z.records[entity][id].related = {}
				}
				apiCounter("ZOHO.CRM.API.getRelatedRecords")
				//console.log(`ZOHO.CRM.API.getRelatedRecords({Entity:${entity}, RecordID:${id}}}),RelatedList:${related}`)
				let res = await ZOHO.CRM.API.getRelatedRecords({Entity:entity, RecordID:id, RelatedList:related})
				let records = res.data
				Z.records[entity][id].related[related] = records
				return records
			},
			getSubform:async function(entity, id, subform){
				if(Z.fields[`${entity}__subform__${subform}`]){
					return Z.fields[`${entity}__subform__${subform}`]
				}
				apiCounter("ZOHO.CRM.API.getFields")
				//console.log(`ZOHO.CRM.META.getFields({Entity:${entity}, RelatedList:${subform}})`)
				let res = await ZOHO.CRM.META.getFields({Entity:subform})
				// let result = await ZOHO.CRM.CONNECTION.invoke("crmapiconnection",{
				// 	"url": `https://www.zohoapis.jp/crm/v4/${zApiScheduleTab}?ids=${deteleScheduleIds.join(",")}`,
				// 	"method" : "DELETE",
				// 	"param_type" : 1
				// })

			},

			getAllRecords:async function(entity){
				apiCounter("ZOHO.CRM.API.getAllRecords")
				//console.log(`ZOHO.CRM.API.getAllRecords({Entity:${entity}})`)
				let module = await Z.getModule(entity)
				let res = await ZOHO.CRM.API.getAllRecords({Entity:entity})
				let records = res.data
				for( record of records){
					if(!Z.records[entity]){ Z.records[entity] = {} }
					if(!Z.records[entity][record.id]){ Z.records[entity][record.id] = {} }
					Z.records[entity][record.id].data = record
				}
				return records
			},
			getFields: async function(entity){
				if(Z.fields[entity]){ return Z.fields[entity] }
				apiCounter("ZOHO.CRM.API.getFields")
				//console.log(`ZOHO.CRM.META.getFields({Entity:${entity}})`)
				let res = await ZOHO.CRM.META.getFields({Entity:entity})
				let fields = res.fields
				Z.fields[entity] = fields
				return fields
			},
			getRelatedList: async function(entity){
				if(Z.fields[entity]?.related_lists){ return Z.fields[entity].related_lists}
				apiCounter("ZOHO.CRM.API.getRelatedList")
				//console.log(`ZOHO.CRM.META.getRelatedList({Entity:${entity}})`)
				let res = await ZOHO.CRM.META.getRelatedList({Entity:entity})
				Z.fields[entity].related_lists = res.related_lists
				return res.related_lists
			},
			getModule: async function(entity){
				if(Z.module[entity]){ return Z.module[entity] }
				apiCounter("ZOHO.CRM.API.getModules")
				//console.log(`ZOHO.CRM.META.getModules()`)
				let res = await ZOHO.CRM.META.getModules()
				let modules = res.modules
				for( module of modules){
					if(module.api_name == entity){
						Z.module[entity] = module
						return module
					}
				}
			},
			getAllModules: async function(){
				if(Z.modules){ return Z.modules }
				apiCounter("ZOHO.CRM.API.getModules")
				//console.log(`ZOHO.CRM.META.getModules()`)
				let res = await ZOHO.CRM.META.getModules()
				let modules = res.modules
				Z.modules = modules
				return modules
			}
		}

		ZS = {
			copySheetApiCount:0,
			sheetNames:[],
			sheetContents:[],
			timer : new Worker("timer.js"),
			zsApi:async function(url, method, param_type, parameters){
				return new Promise(async function(resolve, reject){
					try{
						ZOHO.CRM.CONNECTION.invoke("zohooauth",{
							"url": url,
							"method" : method,
							"param_type" : param_type,
							"parameters" : parameters
						}).then(function(result){
							console.log(parameters)
							console.log(result)
							if(!result.details || result.details.statusMessage.error_code){
								// debugger
								if(result.details.statusMessage.error_code == "2950"){
									alert("APIの呼び出し回数が上限に達しました。処理件数を減らして再度実行してください。")
									ZOHO.CRM.UI.Popup.close()
									return
								}else{
									debugger
									alert("エラーが発生しました。再度実行してください。")
									ZOHO.CRM.UI.Popup.close()
									return
								}
								// console.log(`## Api Result -> ${JSON.stringify(result)} ##`)
								// console.log(`## Wait -> ${JSON.stringify(parameters)} ##`)
								// if(!ZS.timer.onmessage){
								// 	let counter = 0
								// 	ZS.timer.onmessage = async function(e) {
								// 		console.log(`## Waiting... ${counter*10}sec`)
								// 		if(counter > 31){
								// 			console.log(`## Retry -> ${JSON.stringify(parameters)} ##`)
								// 			let r = await ZS.zsApi(url, method, param_type, parameters)
								// 			resolve(r)
								// 		}
								// 		counter++
								// 	}
								// }
								// ZS.timer.postMessage(10000);
							}else{
								resolve( result.details.statusMessage )
							}
						}).catch(function(error){
							console.log(error)
							alert("APIの呼び出し回数が上限に達しました。処理件数を減らして再度実行してください。")
							reject(error)
						})
					} catch(error){
						alert("エラーが発生しました。再度実行してください。" + JSON.stringify(error))
						ZOHO.CRM.UI.Popup.close()
						return
					}
				})
			},
			getSheetContents: async function(wbid,wsid,force=false){
				if(ZS.sheetContents[wbid]?.[wsid] && force != true){ return ZS.sheetContents[wbid][wsid] }
				if(!ZS.sheetContents[wbid]){ ZS.sheetContents[wbid] = {} }
				if(!ZS.sheetContents[wbid][wsid]){ ZS.sheetContents[wbid][wsid] = {} }
				// //使用範囲を取得
				// console.log(`worksheet.usedarea -> ${wbid}.${wsid}`)
				// let result = await ZS.zsApi(
				// 	`https://sheet.zoho.jp/api/v2/${wbid}`,"POST",1,
				// 	{
				// 		method:"worksheet.usedarea",
				// 		worksheet_id:wsid
				// 	}
				// )
				// maxCol = result.used_column_index
				// maxRow = result.used_row_index

				//プログレスバーの設定
				// progressAddTask( maxRow * maxCol )
				// perCellProgressStep = perSheetProgressStep / ( maxRow * maxCol )


				//使用範囲のコンテンツを取得
				apiCounter("worksheet.content.get")
				//console.log(`worksheet.content.get -> ${wbid}.${wsid}`)
				result = await ZS.zsApi(
					`https://sheet.zoho.jp/api/v2/${wbid}`,"POST",1,
					{
						method:"worksheet.content.get",
						worksheet_id:wsid,
						start_row:1,
						start_column:1,
						visible_rows_only:false,
						visible_columns_only:false
					}
				)
				let contents = result.range_details
				let maxCol = result.used_column
				let maxRow = result.used_row

				//コンテンツのない行を補完
				for(let r=0; r<maxRow; r++){
					let contentRow = contents.find( (row) => row.row_index == r+1 )
					if(!contentRow){
						contentRow = { row_index:r+1, row_details:[] }
						contents.push(contentRow)
					}
				}
				// contentsをrowIndexでソート
				contents.sort( (a,b) => a.row_index - b.row_index )

				//コンテンツのない列を補完
				for(let r=0; r<maxRow; r++){
					let contentRow = contents.find( (row) => row.row_index == r+1 )
					for(let c=0; c<maxCol; c++){
						let contentCol = contentRow.row_details.find( (col) => col.column_index == c+1 )
						if(!contentCol){
							contentCol = { column_index:c+1, content:"" }
							contentRow.row_details.push(contentCol)
						}
					}
					// row_detailsをcolumnIndexでソート
					contentRow.row_details.sort( (a,b) => a.column_index - b.column_index )
				}

				ZS.sheetContents[wbid][wsid] = contents
				return contents
			},
			getWorksheetList: async function(wbid, force=false){
				if(ZS.sheetNames[wbid] && force != true){ return ZS.sheetNames[wbid] }
				if(!ZS.sheetNames[wbid]){ ZS.sheetNames[wbid] = {} }
				apiCounter("worksheet.list")
				// console.log(`worksheet.list -> ${wbid}`)
				let res = await ZS.zsApi(
					`https://sheet.zoho.jp/api/v2/${wbid}`,"POST",1,
					{ method:"worksheet.list" }
				)

				let worksheets = res.worksheet_names
				ZS.sheetNames[wbid] = worksheets
				return worksheets
			},
			deleteRows: async function(wbid,wsid,rows){
				apiCounter("worksheet.rows.delete")
				//console.log(`worksheet.rows.delete -> ${wbid}.${wsid}`)
				let result = await ZS.zsApi(
					`https://sheet.zoho.jp/api/v2/${wbid}`,"POST",1,
					{
						method:"worksheet.rows.delete",
						worksheet_id:wsid,
						row_index_array:rows
					}
				)
				return result
			},
			updateSheetViaCsv: async function(wbid,wsid,csv){
				apiCounter("worksheet.csvdata.set")
				//console.log(`worksheet.csvdata.set -> ${wbid}.${wsid}`)
				let result = await ZS.zsApi(
					`https://sheet.zoho.jp/api/v2/${wbid}`,"POST",1,
					{
						method:"worksheet.csvdata.set",
						worksheet_id:wsid,
						row:1,
						column:1,
						ignore_empty:true,
						data:csv
					}
				)
				console.log("### worksheet.csvdata.set result ###")
				console.log(result)
				return result
			},
			copySheet: async function(wbid,origWsid,newWsName){
				if(ZS.copySheetApiCount == 30){
					debugger
					apiCounter("workbook.copy")
					//console.log(`workbook.copy -> ${wbid}.${origWsid}.${newWsName}`)
					let result = await ZS.zsApi(
						`https://sheet.zoho.jp/api/v2/copy`,"POST",1,
						{
							method:"workbook.copy",
							resource_id:wbid,
						}
					)
					WORKING_BOOK_ID = result.resource_id
					ZS.sheetContents[WORKING_BOOK_ID] = ZS.sheetContents[wbid]
					wbid = WORKING_BOOK_ID
					ZS.copySheetApiCount = 0
					await deleteFile(wbid)
				}
				apiCounter("worksheet.copy")
				// console.log(`worksheet.copy -> ${wbid}.${origWsid}.${newWsName}`)
				let result = await ZS.zsApi(
					`https://sheet.zoho.jp/api/v2/${wbid}`,"POST",1,
					{
						method:"worksheet.copy",
						worksheet_id:origWsid,
						new_worksheet_name:newWsName
					}
				)
				ZS.copySheetApiCount++
				return result
			},
			deleteSheet: async function(wbid,wsid){
				ZS.sheetContents[wbid][wsid] = null
				apiCounter("worksheet.delete")
				// console.log(`worksheet.delete -> ${wbid}.${wsid}`)
				let result = await ZS.zsApi(
					`https://sheet.zoho.jp/api/v2/${wbid}`,"POST",1,
					{
						method:"worksheet.delete",
						worksheet_id:wsid
					}
				)
				return result
			}
		


		}





		if(mode == "payment"){
			progressAddStep(2)
			let createPaymentDataFunc = `${ApiDomain}/crm/v2/functions/y_test02/actions/execute?auth_type=apikey&zapikey=1003.d89544f13ce46973dab74091db544213.e1e53e7f681a1b2b5a1cdd5243e1e74f`
			let progressMsg = document.getElementById("progressMsg")
			let closeBtn = document.getElementById("closeBtn")
			closeBtn.addEventListener("click", async function(){
				// debugger
				await ZOHO.CRM.UI.Record.open({Entity:data.Entity,RecordID:0})
				await ZOHO.CRM.UI.Record.open({Entity:data.Entity,RecordID:data.EntityId[0]})
				// let ORG = await ZOHO.CRM.CONFIG.getOrgInfo()
				// let Modules = await ZOHO.CRM.META.getModules()
				// let moduleInfo = Modules.modules.find((module) => { return module.api_name == data.Entity })
				// window.open(`https://crm.zoho.jp/crm/${ORG.org[0]["domain_name"]}/tab/${moduleInfo["module_name"]}/${data.EntityId}`, "_blank")
				// parent.close()
			})

			document.querySelector("#invoice").style.display = "none"
			document.querySelector("#payment").style.display = "block"
			document.querySelector("#estimate").style.display = "none"
			
			templateUrl = await ZOHO.CRM.API.getOrgVariable("Payment_Template_Url")
			if(!templateUrl.Success){
				ZOHO.CRM.UI.Popup.close()
				return
			}
			progressNext()



			//####支払明細作成だけ走らせる場合はここから####
			// progressMsg.innerHTML = "支払明細書を作成中"
			// progressNext()
			// templateSelectoerSetup(templateUrl.Success.Content, document.querySelector("#payTemplateSelect"))
			// zSheetTemplate = document.getElementById("payTemplateSelect").value
			// let workbookUrl = await createPaymentViaSheets(widgetData)
			//####支払明細作成だけ走らせる場合はここまで####




			progressMsg.innerHTML = "支払データを作成中"
			// debugger
			if(paymentDocsOnly != "yes"){
				ZOHO.CRM.FUNCTIONS.execute(
					"y_test02",
					{
						"arguments":JSON.stringify({
							"id":String(data.EntityId)
						})
					}
				).then(async function(funcRecult){
					await createDocs(funcRecult)
				})
			}else{
				createDocsOnly()
			}

			async function createDocsOnly(){
				progressMsg.innerHTML = "支払明細書を作成中"
				progressNext()
				templateSelectoerSetup(templateUrl.Success.Content, document.querySelector("#payTemplateSelect"))
				zSheetTemplate = document.getElementById("payTemplateSelect").value
				let workbookUrl = await createPaymentViaSheets(widgetData)

				if(IP != "61.200.96.103"){
					let updateData = {
						"id":data.EntityId[0],
						"Payment_Doc_Url":workbookUrl
					}	
					let updateResult = await ZOHO.CRM.API.updateRecord({
						Entity:"Gyara_Payment",
						APIData:updateData
					})
				}

				progressMsg.innerHTML = `支払データ作成処理が完了しました。<br>📑<a target="_blank" href="${workbookUrl}">支払明細書を開く</a>`
				document.getElementById("progress").style.display = "none"
				document.getElementById("closeBtnArea").style.display = "flex"
			}

			async function createDocs(funcRecult){
				if(funcRecult?.details?.output != "支払データの作成が完了しました"){
					progressMsg.innerHTML = funcRecult?.details?.output ? funcRecult?.details?.output : "支払データの作成中にエラーが発生しました<br><font size=-1>関数エラー："+funcRecult.message+"</font>"
					document.getElementById("progress").style.display = "none"
					document.getElementById("closeBtnArea").style.display = "flex"
					return
				}
				progressMsg.innerHTML = "支払明細書を作成中"
				progressNext()
				templateSelectoerSetup(templateUrl.Success.Content, document.querySelector("#payTemplateSelect"))
				zSheetTemplate = document.getElementById("payTemplateSelect").value
				let workbookUrl = await createPaymentViaSheets(widgetData)
				debugger
				let userMessages = funcRecult?.details?.userMessage
				for(let msg of userMessages){
					let msgInfo = JSON.parse(msg)
					if(msgInfo.Payment_Info){
						// let updateData = {
						// 	"id":msgInfo.Payment_Info.id,
						// 	"Payment_Doc_Url":workbookUrl
						// }	
						// let updateResult = await ZOHO.CRM.API.updateRecord({
						// 	Entity:"Payments",
						// 	APIData:updateData
						// })
						continue
					}
					if(msgInfo.Gyara_Payment_Info){
						let GyaraPaymentInfo = await Z.getRecord("Gyara_Payment", msgInfo.Gyara_Payment_Info.id)
						let payDestinationInfo = await Z.getRecord("Payment_Destination_M", GyaraPaymentInfo.Payment_Destination_Name.id)
						let DocMethod = payDestinationInfo.Doc_Send_Method
						
						let updateData = {
							"id":msgInfo.Gyara_Payment_Info.id,
							"Payment_Doc_Url":workbookUrl,
							"Doc_Send_Method":DocMethod,
							"Payment_Doc_Status":"未処理",
							"Email":payDestinationInfo.Email,
						}	
						let updateResult = await ZOHO.CRM.API.updateRecord({
							Entity:"Gyara_Payment",
							APIData:updateData
						})
						continue
					}

				}

				progressMsg.innerHTML = `支払データ作成処理が完了しました。<br>📑<a target="_blank" href="${workbookUrl}">支払明細書を開く</a>`
				document.getElementById("progress").style.display = "none"
				document.getElementById("closeBtnArea").style.display = "flex"

			}





		}

		if(mode == "invoice"){
			// debugger
			document.querySelector("#invoice").style.display = "block"
			document.querySelector("#payment").style.display = "none"
			document.querySelector("#estimate").style.display = "none"
			templateUrl = await ZOHO.CRM.API.getOrgVariable("Invoice_Template_Url")
			if(!templateUrl.Success){
				alert("テンプレートURLが設定されていません。")
				ZOHO.CRM.UI.Popup.close()
				return
			}
			templateSelectoerSetup(templateUrl.Success.Content, document.querySelector("#invTemplateSelect"))
		}

		if(mode == "estimate"){
			document.querySelector("#invoice").style.display = "none"
			document.querySelector("#payment").style.display = "none"
			document.querySelector("#estimate").style.display = "block"
			templateUrl = await ZOHO.CRM.API.getOrgVariable("Estimate_Template_Url")
			if(!templateUrl.Success){
				alert("テンプレートURLが設定されていません。")
				ZOHO.CRM.UI.Popup.close()
				return
			}
			templateSelectoerSetup(templateUrl.Success.Content, document.querySelector("#estTemplateSelect"))
		}

		function templateSelectoerSetup(v, elm){
			// debugger
			let variable = v
			let zTemplateName
			let templateOptionHtml = ""
			if(variable.match(/\n/g)){
				let templateUrlArray = variable.split("\n")
				for(let url of templateUrlArray){
					zTemplateName = url.split("#").shift()
					zSheetTemplate = url.split("/").pop()
					if(zSheetTemplate.match(/\?/g)){
						zSheetTemplate = zSheetTemplate.split("?").shift()
					}
					
					templateOptionHtml += `<option value="${zSheetTemplate}">${zTemplateName}</option>`
				}
				elm.innerHTML = templateOptionHtml
			}else{
				zTemplateName = variable.split("#").shift()
				zSheetTemplate = variable.split("/").pop()
				if(zSheetTemplate.match(/\?/g)){
					zSheetTemplate = zSheetTemplate.split("?").shift()
				}
				templateOptionHtml += `<option value="${zSheetTemplate}" selected>${zTemplateName}</option>`
				elm.innerHTML = templateOptionHtml
				elm.disabled = true
			}
		}



		// criteriaCheck(await Z.getRecord(data.Entity, data.EntityId[0]), await Z.getFields(data.Entity), "ギャラ支払入力の担当者??ギャラ支払NO=43&&(タレントCD=876568||支払先CD=99999)")




// debugger
		
		ZOHO.CRM.UI.Resize({height:"200", width:"540"})

		ENTITY = data.Entity
		ENTITY_IDS = data.EntityId

		FIELDS[ENTITY] = await Z.getFields(ENTITY)

		
		//レコード情報を取得
		//RECORD = await Z.getRecord(ENTITY, ENTITY_ID)
		//debugger

		//Excelテンプレートの全レコードを取得
		// let Templates = await Z.getAllRecords("Excel_Templates")
		// console.log(Templates)
		
		//ENTITYのモジュール情報を取得
		let modules = await Z.getAllModules()
		for(let idx in modules){
			if(modules[idx].api_name == ENTITY){
				MODULE = modules[idx]
				break
			}
		}

		if(mode == "invoice"){initInvoiceUI()}
		if(mode == "payment"){initPaymentUI()}
		if(mode == "estimate"){initEstimateUI()}


	})

	ZOHO.embeddedApp.init();
	// loadHtml("https://air.southernwave.net/ZohoWidgetDevelopment/RecordToExcel/widget.html")



	async function createPaymentViaSheets(data){
		progressAddStep(1)
		progressAddStep(data.EntityId.length)

		// labels = "<Z%日程->日程名%Z>"
		// let rrr = replaceZohoFieldVariables(labels, "Bookings", "10796000013735411")
		// return

		document.getElementById("pay-generateBtnText").style.display = "none"
		document.getElementById("pay-generateBtnInProgress").style.display = "block"
		document.getElementById("pay-generateBtn").setAttribute("disabled", true)

		//保存先フォルダの確認・生成
		baseDirectoryId = "ea7bbc056ac8b1816437798bd5794986ee669"
		let dirList = await getDirList(baseDirectoryId)
		let directoryId, directoryUrl

		//年フォルダがなければ作成
		let yearDirectory = dirList.find( (d) => d.attributes.name == moment().format("YYYY年") )
		if( ! yearDirectory ){
			let mkdirRes = await createFolderOnWorkDrive(moment().format("YYYY年"), baseDirectoryId)
			mkdirOutput = JSON.parse(mkdirRes.details.output)
			directoryId = mkdirOutput.data.id
		}else{
			directoryId = yearDirectory.id
		}

		dirList = await getDirList(directoryId)
		//月フォルダがなければ作成
		// debugger
		let monthDirectory = dirList.find( (d) => d.attributes.name == moment().format("YYYY年MM月") )
		if( ! monthDirectory ){
			let mkdirRes = await createFolderOnWorkDrive(moment().format("YYYY年MM月"), directoryId)
			mkdirOutput = JSON.parse(mkdirRes.details.output)
			directoryId = mkdirOutput.data.id
			directoryUrl = mkdirOutput.data.attributes.permalink
		}else{
			directoryId = monthDirectory.id
			directoryUrl = monthDirectory.attributes.permalink
		}

		progressNext()

		
		// let record = await Z.getRecord("Gyara_Payment", e)
		// let talentInfo = await Z.getRecord("Talents", record.Talent_Name.id)
		// let folderName = talentInfo.Talent_Cd + "_" + talentInfo.Name

		//Sheetテンプレートから新規作成
		// debugger
		let recordData = await Z.getRecord("Gyara_Payment", data.EntityId[0])
		if(IP == "61.200.96.103"){ fileNameAddition += "TEST" }
		let WorkbookName = `${recordData.Talent_Cd}_${recordData.Talent_Name.name}_${fileNameAddition}`

		let createBookRes = await createSheetFromTemplate(WorkbookName, zSheetTemplate)
		createdWorkbookId = createBookRes.details.statusMessage.resource_id
		workbookUrl = createBookRes.details.statusMessage.workbook_url

		WORKING_BOOK_ID = createdWorkbookId

		//SheetをWorkDriveに移動
		// if(IP != "61.200.96.103"){
			let mvRes = await moveFile(createdWorkbookId, directoryId)
		// }
		
		let worksheets = await ZS.getWorksheetList(createdWorkbookId)
		//debugger
		await generateSheet(WORKING_BOOK_ID, "Gyara_Payment", data.EntityId)
		
		// contents = await replaceSheetVariables(createdWorkbookId, worksheets[0].worksheet_id, "Gyara_Payment", record.id)
		// await clearingRows(createdWorkbookId, worksheets[0].worksheet_id)
		// progressNext()

		return workbookUrl
		window.open(directoryUrl, "_blank")

		// debugger

	
		ZOHO.CRM.UI.Popup.close()

	}


	async function createInvoiceViaSheets(data){
		// debugger
		progressAddStep(1)

		document.getElementById("generateBtnText").style.display = "none"
		document.getElementById("generateBtnInProgress").style.display = "block"
		document.getElementById("generateBtn").setAttribute("disabled", true)
		// document.getElementById("generateGatherBtn").setAttribute("disabled", true)
		
		let remindElm = document.getElementById("remindCheck")
		let remindReason = document.getElementById("remindReason")

		let parentDirectoryId = "s43w5ba4a8c6bc4234c0999c5179e2a9d8752"
		let createdDirectoryId
		let createdWorkbookId
		let res


		let targetRecords = []
		let targetSalesNos = []
		for(let idx in data.EntityId){
			let record = await Z.getRecord(ENTITY, data.EntityId[idx])
			if(record.Sales_Type_Cd != "通常" && record.Sales_Type_Cd != null){ continue }
			if(record.Billing_Statement_Type == "不要／自動入金" /* || record.Billing_Statement_Type == "先方指定のフォーマットで発行" */ ){ continue }
			if(record.Stage == "交渉中" || record.Stage == "確定承認依頼中" || record.Stage == "確定承認済み" || record.Stage == "経理承認依頼中"){ continue }

			targetRecords.push(data.EntityId[idx])
		}
		if(targetRecords.length == 0){
			alert("請求書を作成できる仕事カードがありません")
			ZOHO.CRM.UI.Popup.close()
			return
		}

		GatherSalesNumbers = targetSalesNos.join(",")

		progressAddStep(targetRecords.length)


		// //WorkDriveにフォルダ作成
		let now = moment().format("YYYYMMDD-HHmm")

		res = await createFolderOnWorkDrive(now, parentDirectoryId)
		let output = JSON.parse(res.details.output)
		createdDirectoryId = output.data.id
		directoryUrl = output.data.attributes.permalink

		let WorkbookName = "請求書_" + now + fileNameAddition

		//Sheetテンプレートから新規作成
		res = await createSheetFromTemplate(WorkbookName, zSheetTemplate)
		createdWorkbookId = res.details.statusMessage.resource_id
		WORKING_BOOK_ID = createdWorkbookId
		workbookUrl = res.details.statusMessage.workbook_url
		console.log(workbookUrl)


		let generateResult = await generateSheet(WORKING_BOOK_ID, ENTITY, targetRecords)

		//SheetをWorkDriveに移動
		res = await moveFile(WORKING_BOOK_ID, createdDirectoryId)
		createdDirectoryId = res.details.output

		

		let invoiceApiName = "Invoice_Histories"
		
		let LinkApiName = "X1"
		let LinkDealApiName = "field18_1"
		let LinkRecipientApiName = "field18"


		// debugger


		//###############　負荷テスト用コード　################
		if(data.EntityId[0] != debugRecordId){
		//###############　負荷テスト用コード　################

			let ApiDatas = []
			for(let target of targetRecords){
				ApiDatas.push({
					"Billing_Statement_Date":moment().format("YYYY-MM-DD"),
					"Invoice_Url":workbookUrl,
					"Re_Billing_Reason":remindElm.checked ? remindReason.value : "",
					"Deal":{ "id":target },
					"Trigger":"workflow"
				})
			}
			let invoiceId = (await ZOHO.CRM.API.insertRecord({
				Entity:invoiceApiName,
				APIData:ApiDatas

			})).data[0].details.id


		
			progressNext()

		//###############　負荷テスト用コード　################
		}
		//###############　負荷テスト用コード　################

		ZOHO.CRM.UI.Popup.close()
		// ZOHO.CRM.UI.Popup.closeReload()
		
		// window.open(workbookUrl, "_blank")
		window.open(workbookUrl, "_blank")
		console.log(res)


	}

	async function createEstimateViaSheets(data){

		progressAddStep(1)
		progressAddStep(data.EntityId.length)
	
		document.getElementById("est-generateBtnText").style.display = "none"
		document.getElementById("est-generateBtnInProgress").style.display = "block"
		document.getElementById("est-generateBtn").setAttribute("disabled", true)
		
		let parentDirectoryId = "ea7bb01dbb6e21ef04ae89b5e1e4059b12487"
		let createdDirectoryId
		let createdWorkbookId
		let res
// debugger
		// //WorkDriveにフォルダ作成
		let now = moment().format("YYYYMMDD-HHmm")
		let dirList = await getDirList(parentDirectoryId)

		let username = await ZOHO.CRM.CONFIG.getCurrentUser().then( (user) => user.users[0].full_name )
		// debugger

		//タレント用のフォルダがなければ作成
		let userDirectory = dirList.find( (d) => d.attributes.name == username )
		if( ! userDirectory ){
			let mkdirRes = await createFolderOnWorkDrive(username, parentDirectoryId)
			mkdirOutput = JSON.parse(mkdirRes.details.output)
			directoryId = mkdirOutput.data.id
			directoryUrl = mkdirOutput.data.attributes.permalink
		}else{
			directoryId = userDirectory.id
			directoryUrl = userDirectory.attributes.permalink
		}

		let WorkbookName = "見積書_" + now

		//Sheetテンプレートから新規作成
		res = await createSheetFromTemplate(WorkbookName, zSheetTemplate)
		createdWorkbookId = res.details.statusMessage.resource_id
		workbookUrl = res.details.statusMessage.workbook_url

		WORKING_BOOK_ID = createdWorkbookId
		let generateResult = await generateSheet(WORKING_BOOK_ID, ENTITY, ENTITY_IDS, true)

		//SheetをWorkDriveに移動
		res = await moveFile(createdWorkbookId, directoryId)
		createdDirectoryId = res.details.output

		res = await shareFile(createdWorkbookId)
		// debugger
		

		let estimateApiName = "Estimate_Histories"
		let ApiDatas = []
		for(let target of data.EntityId){
			ApiDatas.push({
				"Estimate_Date":moment().format("YYYY-MM-DD"),
				"Estimate_Url":workbookUrl,
				"Deal":{ "id":target },
				"Trigger":"workflow"
			})
		}
		let estimateId = (await ZOHO.CRM.API.insertRecord({
			Entity:estimateApiName,
			APIData:ApiDatas

		})).data[0].details.id


		
		progressNext()
		// debugger


		ZOHO.CRM.UI.Popup.close()
		// ZOHO.CRM.UI.Popup.closeReload()
		
		// window.open(workbookUrl, "_blank")
		window.open(workbookUrl, "_blank")
		console.log(res)

	}


	async function generateSheet(workbookid, moduleApiName, recordId = [], forceGather = false){

		let maxCol, maxRow
		let currentWsName, currentWsId
		let contents

		let gatherBaseSheetId
		let gatherBaseRecordId

		let sheetInfos = []

		let worksheets = await ZS.getWorksheetList(WORKING_BOOK_ID)
		let singleTemplate = await ZS.copySheet(WORKING_BOOK_ID, worksheets[0].worksheet_id, "single")
		await clearingRows(WORKING_BOOK_ID, singleTemplate.worksheet_names.find(s => s.worksheet_name == singleTemplate.new_worksheet_name).worksheet_id)

		worksheets = await ZS.getWorksheetList(WORKING_BOOK_ID, true)

		zSingleTemplateSheetId = singleTemplate.worksheet_names.find(s => s.worksheet_name == singleTemplate.new_worksheet_name).worksheet_id
		zMultiTemplateSheetId = worksheets[0].worksheet_id
		zSingleTemplateContents = await ZS.getSheetContents(WORKING_BOOK_ID, zSingleTemplateSheetId, true)
		zMultiTemplateContents = await ZS.getSheetContents(WORKING_BOOK_ID, zMultiTemplateSheetId, true)

		//プログレスバーの設定
		perSheetProgressStep = progress_steps / recordId.length

		//テンプレートシートのIDを取得
		//zTemplateSheetId = worksheets[0].worksheet_id

		// debugger
		//###############　負荷テスト用コード　################
		let debug = false
		// if(recordId[0] == debugRecordId){
		// 	// debugger
		// 	recordId = [
		// 		debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,
		// 		debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,
		// 		debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,
		// 		debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId,debugRecordId
		// 	]
		// 	debug = true
		// 	progressAddStep(50)
		// }
		//###############　負荷テスト用コード　################

		for( rId of recordId ){
			let result
			let record = await Z.getRecord(moduleApiName, rId)
			let newSheetName
			if(moduleApiName != "Gyara_Payment"){
				newSheetName = debug ? record.Sales_No + "_" + Date.now() : record.Sales_No
			}else{
				newSheetName = `${record.Talent_Cd}_${record.Talent_Name.name}`
			}

			//テンプレートにするシートの選択
			let templateSheetId
			if(moduleApiName != "Gyara_Payment"){
				let childs = await Z.getRelatedRecords("Deals", rId, "Child_Deals")
				// debugger
				if(childs){
					templateSheetId = zMultiTemplateSheetId
					zCurrentTemplateContents = zMultiTemplateContents
				}else{
					templateSheetId = zSingleTemplateSheetId
					zCurrentTemplateContents = zSingleTemplateContents
				}
			}else{
				templateSheetId = zMultiTemplateSheetId
				zCurrentTemplateContents = zMultiTemplateContents
			}

			//ワークシートをコピー
			result = await ZS.copySheet(WORKING_BOOK_ID, templateSheetId, newSheetName)
			currentWsName = newSheetName
			currentWsId = result.worksheet_names.find( (ws) => ws.worksheet_name == currentWsName ).worksheet_id
			sheetInfos.push({sheetId: currentWsId, sheetName: currentWsName, recordId: rId})

			if(!ZS.sheetContents[WORKING_BOOK_ID]){ZS.sheetContents[WORKING_BOOK_ID] = {}}
			if(!ZS.sheetContents[WORKING_BOOK_ID][currentWsId]){ZS.sheetContents[WORKING_BOOK_ID][currentWsId] = {}}
			// debugger
			let replacedContents = await replaceSheetVariables(WORKING_BOOK_ID, currentWsId, templateSheetId, moduleApiName, rId)

			let sheetContents = replacedContentsToSheetContents(replacedContents)
			ZS.sheetContents[WORKING_BOOK_ID][currentWsId] = sheetContents
			worksheetContents.push({sheetId: currentWsId, contents:contents})
			// await clearingRows(workbookid, currentWsId)
			progressNext()
		}
		
		// debugger

		if(moduleApiName != "Gyara_Payment"){
			if(forceGather){
				if(sheetInfos.length > 1){
					sheetInfos.sort( (a,b) => a.sheetName > b.sheetName ? 1 : -1 )
					let replacedContents = await gatheringSheets(workbookid, sheetInfos[0].sheetId, sheetInfos)
					// let sheetContents = replacedContentsToSheetContents(replacedContents)
					// ZS.sheetContents[workbookid][currentWsId] = sheetContents
					// debugger
				}
				for(let sheet of sheetInfos){
					//await ZS.deleteSheet(workbookid, sheet.sheetId)
				}
			}else{
				for(let sheetInfo of sheetInfos){
					let childs = await Z.getRelatedRecords("Deals", sheetInfo.recordId, "Child_Deals")
					//*****
					//売上種別に「社内請求」含まれない場合の条件を追加
					//*****
					if(childs){
						let gatherSheets = []
						gatherBaseSheetId = sheetInfo.sheetId
						gatherBaseRecordId = sheetInfo.recordId
						for(let child of childs){
							if(gatherBaseRecordId == child.id){ continue }
							let childSheet = sheetInfos.find( (sheet) => sheet.recordId == child.id )
							if(childSheet){
								gatherSheets.push(childSheet)
							}
						}
						if(gatherSheets.length > 0){
							gatherSheets.push( sheetInfos.find( (sheet) => sheet.sheetId == gatherBaseSheetId ) )
							//gatherSheets内のオブジェクトプロパティsheetNameで順ソート
							gatherSheets.sort( (a,b) => a.sheetName > b.sheetName ? 1 : -1 )
							let replacedContents = await gatheringSheets(workbookid, gatherBaseSheetId, gatherSheets)
							// let sheetContents = replacedContentsToSheetContents(replacedContents)
							// ZS.sheetContents[workbookid][currentWsId] = sheetContents
						}
					}
				}
			}
		}


		// if(gather){ await gatheringSheets() }



		//雛形シートを削除
		await ZS.deleteSheet(WORKING_BOOK_ID, zSingleTemplateSheetId)
		await ZS.deleteSheet(WORKING_BOOK_ID, zMultiTemplateSheetId)
		// debugger
		for(let sheetid in ZS.sheetContents[WORKING_BOOK_ID]){
			if(!ZS.sheetContents[WORKING_BOOK_ID][sheetid]){ continue }
			await clearingRows(WORKING_BOOK_ID, sheetid)
		}

		progressNext()
	}

	async function gatheringSheets(workbookid, baseSheetId, gatherSheetIds = []){
		// debugger
		let salesNos = []
		for(let r of gatherSheetIds){
			let record = await Z.getRecord("Deals", r.recordId)
			salesNos.push(record.Sales_No)
		}
		let gatherSheetName = salesNos.join(",")
		// debugger
		if(gatherSheetName.length > 31){
			gatherSheetName = gatherSheetName.substring(0,29) + "…"
		}
		//ワークシートをコピー
		let result = await ZS.copySheet( workbookid, baseSheetId, gatherSheetName )

		let currentWsName = result.new_worksheet_name
		let currentWsId = result.worksheet_names.find( (ws) => ws.worksheet_name == currentWsName ).worksheet_id

		
		let gatherSheetContents = await ZS.getSheetContents(workbookid, currentWsId)

		//合算シートのコンテンツを１列目以外を空にする
		for(let row of gatherSheetContents){
			for(let col of row.row_details){
				if(col.column_index == 1){ continue }
				col.content = ""
			}
		}


		//生成済みのシートのコンテンツを取得
		let gatheredRow = []
		// debugger
		for(let sheetIdInfo of gatherSheetIds){
			let sheetId = sheetIdInfo.sheetId
			if(sheetId == zSingleTemplateSheetId){ continue }
			if(sheetId == zMultiTemplateSheetId){ continue }
			// if(sheetId == baseSheetId ){ continue }
			if(sheetId == currentWsId ){ continue }

			let wsContents = ZS.sheetContents[workbookid][sheetId]

			//各シートの行をループ
			for(let contentsRowIdx in wsContents){

				//１列目が空ならスキップ
				if(wsContents[contentsRowIdx].row_details[0].content == ""){ continue }

				//1列目が繰り返し終了キーならスキップ
				if(wsContents[contentsRowIdx].row_details[0].content.slice(-1) == "/"){ continue }

				//空でなければ、繰り返しキーを取得
				let key = wsContents[contentsRowIdx].row_details[0].content

				//行のコンテンツが空ならスキップ
				let isEmptyRow = true
				for(let isEmptyCol of wsContents[contentsRowIdx].row_details){
					let c = isEmptyCol.content
					if(
						c != "" &&
						c != key &&
						c != key + "/" &&
						c != "☆" &&
						c != "※" &&
						c != '" "&CHAR(10)&" "' &&
						!c.match(/[ ]+/)
					){
						// debugger
						// for(let s of c){
						// 	console.log(s.charCodeAt(0))
						// }
						isEmptyRow = false
						break
					}
				}
				if(isEmptyRow){ continue }



				//合算シートを行ごとにループ
				for(let gatherRow of gatherSheetContents){
					//繰り返しキーが合致したら
					if(gatherRow.row_details[0].content == key){
						//転記済みの行ならスキップ
						if(gatheredRow.find( r => r == gatherRow.row_index)){ continue }
						//合算シートに行をコピー
						debugger
						gatherRow.row_details = wsContents[contentsRowIdx].row_details
						//行番号を転記済みとして記録
						gatheredRow.push(gatherRow.row_index)

						break
					}
				}
			}

			await ZS.deleteSheet(workbookid, sheetId)
		}
		// debugger

		//特例措置：請求Noは決め打ちで記入
		gatherSheetContents.find( row => row.row_index == 2 ).row_details[12].content = "No. " + gatherSheetName.replace(/,/g, "/")
		console.log(gatherSheetContents)
		//contentsからcsvデータを生成。ダブルクオーテーションと改行コードをエスケープする
		let csvData = gatherSheetContents.map( (row) => row.row_details.map( (col) => `"${col.content.replace(/"/g, '""').replace(/\n/g, '\r')}"` ).join(",") ).join("\n")
		// debugger


		//csvデータでシートを更新
		await ZS.updateSheetViaCsv(workbookid, currentWsId, csvData)
		// await clearingRows(workbookid, currentWsId, gatherSheetContents)
		zCurrentTemplateContents = zMultiTemplateContents
		ZS.sheetContents[workbookid][currentWsId] = replacedContentsToSheetContents(gatherSheetContents)
		return gatherSheetContents
	}



	async function clearingRows(workbookId, sheetId){
		// debugger
		let contentsArray = await ZS.getSheetContents(workbookId, sheetId)
		//繰り返し行のキー
		let repeatRowKeys = []
		let repeatStartRow = 0
		let repeatEndRow = 0
		let deleteRowIndexArray = []
		//繰り返し行の余分な行を削除
		let deletStartRow = 1
		let deletEndRow = 0

		//リピートキーのある行番号をキーごとに抽出
		for(let r in contentsArray){
			let key = contentsArray[r].row_details[0].content
			if(key == ""){ continue }
			if( repeatRowKeys.find( k => k.key == key ) ){
				repeatRowKeys.find( k => k.key == key ).rows.push(r*1)
			}else{
				repeatRowKeys.push({key:key, rows:[r*1]})
			}
		}
		//リピートキーごとに削除開始と終了行を決定
		let lastContentRow = 0
		
		for(let keyRows of repeatRowKeys){
			if(keyRows.key.slice(-1) == "/"){ continue }
			let minRowCount = 5
			if(keyRows.key.slice(-1) == "3"){ minRowCount = 3}
			if(keyRows.key.slice(-1) == "1"){ minRowCount = 1}
			let rowItemCount = 0
			for(let idx in keyRows.rows){
				row = keyRows.rows[idx]
				if(idx < minRowCount){
					lastContentRow = row*1
					if(keyRows.rows[idx*1+1]){
						deletStartRow = keyRows.rows[idx*1+1] + 1
					}
					continue
				}
				//もしコンテンツがあれば、最終コンテンツ行を更新
				if(contentsArray[row].row_details.find( 
					c => c.content != "" &&
					c.column_index != 1 &&
					c.content != " " &&
					c.content != "  " &&
					c.content != /[ ].+/ &&
					c.content != "☆" &&
					c.content != "※"
				)){
					lastContentRow = row*1
					if(keyRows.rows[idx*1+1]){
						deletStartRow = keyRows.rows[idx*1+1] + 1
					}
					continue
				}
				//コンテンツがなく、最終コンテンツ行よりも行番号が大きければ、削除開始行を更新
				if(lastContentRow+1 > deletStartRow){
					if(keyRows.rows[idx*1+1]){
						deletStartRow = keyRows.rows[idx*1+1] + 1
					}
				}
			}
			deletEndRow = repeatRowKeys.find(k => k.key == keyRows.key + "/").rows[0]
			if(deletEndRow - deletStartRow == 0 || deletEndRow - deletStartRow == 1){ continue }
			deleteRowIndexArray.push({start_row:deletStartRow, end_row:deletEndRow})
		}


		// for(let r in contentsArray){
			
		// 	//繰り返し行かどうかを判定
		// 	let key = contentsArray[r].row_details[0].content
		// 	if(key == ""){ continue }
			
		// 	//最低限の行数を確保。同じキーの出現が５回以下の場合はスキップする
		// 	if( repeatRowKeys.find( k => k.key == key.left(1) ) ){
		// 		repeatRowKeys.find( k => k.key == key.left(1) ).count++
		// 	}else{
		// 		repeatRowKeys.push({key:key, count:1})
		// 	}

		// 	if(repeatRowKeys.find(k => k.key == key.left(1)).count <= 5){ continue }
		// 	if(key.right(1) == "/"){
		// 		deletEndRow = r*1-1
		// 		continue
		// 	}
		// 	if(contentsArray[r*1].row_details.find( c => c.content != "" && c.column_index != 1)){continue}
		// 	deletStartRow = r*1+1
			
		// 	//繰り返し開始行を保持
		// 	if(repeatStartRow == 0){ repeatStartRow = r*1+3 }

		// 	//次の行が繰り返し行じゃなければ繰り返し終了行を保持
		// 	if(key == ""){
		// 		repeatEndRow = r*1-1
		// 		deleteRowIndexArray.push({start_row:repeatStartRow, end_row:repeatEndRow})
		// 		repeatStartRow = 0
		// 		repeatEndRow = 0
		// 		repeatRowKeys = []
		// 	}
		// }

		// debugger
		if(deleteRowIndexArray.length == 0){ return }
		let result = await ZS.deleteRows(workbookId, sheetId, deleteRowIndexArray)
		console.log(result)
	}


	async function replaceSheetVariables(workbookId, sheetId, tempSheetId, moduleApiName, recordId){
		let contents = await ZS.getSheetContents(workbookId, tempSheetId)
		let replaceContents = []
		for(let r of contents){
			let row = {row_index: r.row_index, row_details:[]}
			for(let c of r.row_details){
				let col = {column_index: c.column_index, content: ''}
				row.row_details.push(col)
			}
			replaceContents.push(row)
		}

		//置換変数を置換
		for(let rowInfo of contents){
			let rowContents = rowInfo.row_details.map( (colInfo) => colInfo.content )
			for(let colInfo of rowInfo.row_details){
				// document.getElementById("progressBar").style.width = progressCurrent() + "%"

				let content = colInfo.content
				if(colInfo.column_index == 1){ continue }
				if( ! content.match(/<Z%[^%]*%Z>/g) ){ continue }
				// debugger
				let values = await replaceZohoFieldVariables( content, moduleApiName, recordId )
				if(values.length == 0){ values = [" "] }
				for(idx in values){
					// if(values[idx] == " "){ continue }
					let colNum = colInfo.column_index
					let rowNum = rowInfo.row_index*1 + idx*1
					let targetRow = replaceContents.find( (row) => row.row_index == rowNum )
					if(targetRow){
						let targetCol = targetRow.row_details.find( (col) => col.column_index == colNum ).content = values[idx]
						//targetCol.content = values[idx]
						// debugger
					}
				}
			}
		}

		//contentsからcsvデータを生成。ダブルクオーテーションとカンマと改行コードをエスケープする
		let csvData = replaceContents.map( (row) => row.row_details.map( (col) => `"${col.content.replace(/"/g, '""').replace(/,/g, '\,').replace(/\n/g, '\r')}"` ).join(",") ).join("\n")
		// debugger


		//csvデータでシートを更新
		ZS.updateSheetViaCsv(workbookId, sheetId, csvData)
		return replaceContents
		// return await ZS.getSheetContents(workbookId, sheetId)
	}




	//Zohoフィールド変数を置き換える関数
	async function replaceZohoFieldVariables(ZVarString, moduleApiName, recordId){
		// debugger
		if(ZVarString.includes("\n")){
			let strs = ZVarString.split("\n")
			ZVarString = '="' + strs.join(`"&CHAR(10)&"`) + '"'
		}
		if(typeof ZVarString === "object" && ZVarString !== null){ return [" "] }

		// debugger
		//ZVarString = "【<Z%商談@->商談名%Z>】<Z%商談@->商品@->商品名%Z>"


		//置き換え文字が有効でなければそのまま返す
		let matches = ZVarString.match(/<Z%[^%]*%Z>/g)
		if(matches === null){ return [" "] }
		if(typeof matches[Symbol.iterator] === "undefined"){ return [" "] }

		//置換え結果の配列変数
		let ReplacedString = []
		
		//置き換え文字があれば置き換える
		for(match of matches){
			// debugger

			// 置き換え文字を -> で分割して、ZVarLabelsに格納
			let ZVarLabels = match.replace("<Z%","").replace("%Z>","").split("->")

			// ZVarLabelsを元に、置き換えを実行
			let ZVarValues = await _getDataFromZVarString(ZVarLabels, moduleApiName, await Z.getRecord(moduleApiName, recordId) )

			//置き換えが失敗してなければ
			if(ZVarValues !== false){
				//置換え結果を元に、ReplaceStringを作成

				for([index, ZVarValue] of Object.entries(ZVarValues)){
					let nextIndex = Number(index) + 1
					if(index == 0 && ReplacedString[index] == undefined){ ReplacedString[index] = ZVarString }
					if(ZVarValues[nextIndex] !== undefined && ReplacedString[nextIndex] == undefined){ ReplacedString[nextIndex] = ReplacedString[index] }
					ReplacedString[index] = ReplacedString[index].replace(match, ZVarValue == null ? " " : ZVarValue )
				}
			}else{
				return [" "]
			}
		}
		//debugger
		return ReplacedString
	}

	async function _getDataFromZVarString(ZVarLabels, tabApiName, record){
		let fieldType

		let f = await Z.getFields(tabApiName)
		let fieldTypeTest = f.find( function(data){
			return (
				data.display_label == ZVarLabels[0] || 
				data.field_label == ZVarLabels[0] ||
				data.name == ZVarLabels[0] || 
				data.api_name == ZVarLabels[0]
			) ? data : false 
		})

		if(!fieldTypeTest){
			if(record.hasOwnProperty("Parent_Id")){
				fieldType = "field"
			}else{
				let rFieldTest = await Z.getRelatedList(tabApiName)
				fieldTypeTest = rFieldTest.find(function(data){
					return (
						data.display_label == ZVarLabels[0] || 
						data.field_label == ZVarLabels[0] ||
						data.name == ZVarLabels[0] || 
						data.api_name == ZVarLabels[0]
					) ? data : false 
				})
				if(fieldTypeTest){
					fieldType = "relatedlist"
				}else{
					fieldType = "field"
				}
			}
		}else{
			if(fieldTypeTest.data_type == "subform"){
				fieldType = "subform"
			}else{
				fieldType = "field"
			}
		}



		//関連リストからのデータ取得であれば
		if(ZVarLabels[0].match(/\@$/g) || fieldType == "relatedlist"){
			//関連リストの情報を取得
			tabAllRelatedListInfo = await getRelatedListInfo(tabApiName)
			//関連リストの情報から、リスト名／API名／表示名が一致するものを取得
			let ZvarRelatedListInfo = tabAllRelatedListInfo.find( function(data){
				return (
					data.display_label == ZVarLabels[0].replace(/@$/,"") || 
					data.field_label == ZVarLabels[0].replace(/@$/,"") || 
					data.name == ZVarLabels[0].replace(/@$/,"") || 
					data.api_name == ZVarLabels[0].replace(/@$/,"")
				) ? data : false 
			})
			//関連リストの情報が取得できれば
			if(ZvarRelatedListInfo){
				// debugger
				//関連リストのレコードを取得
				let relatedRecords = await Z.getRelatedRecords(tabApiName, record.id, ZvarRelatedListInfo.api_name)
				let returnValues = []
				for( relatedRecord of relatedRecords){
					//関連リストのレコードから、再帰的にデータを取得
					let rRecord = await Z.getRecord(ZvarRelatedListInfo.module.api_name, relatedRecord.id)
					let replaceContent = await _getDataFromZVarString(ZVarLabels.slice(1), ZvarRelatedListInfo.module.api_name, rRecord)
					if(replaceContent == " "){ continue }
					returnValues = returnValues.concat( replaceContent )
				}
				console.log(returnValues)
				return returnValues
			}else{
				return [" "]
			}
		}else{

			//タブのフィールド情報を取得
			tabAllFieldsInfo = await Z.getFields(tabApiName)

			//レコード情報を取得
			//let record = await Z.getRecord(tabApiName, recordId)


			//条件付きの場合
			if(ZVarLabels[0].match(/\?\?/g)){
				// debugger
				if( criteriaCheck(record, tabAllFieldsInfo, ZVarLabels[0]) ){
					ZVarLabels[0] = ZVarLabels[0].split("??")[0]
				}else{
					return [" "]
				}
			}



			// debugger
			//実行日時はフォーマットして返す
			if(ZVarLabels[0] === "+実行日時"){
				return [moment().format(ZVarLabels[1])]
			}

			if(ZVarLabels[0] === "+スタンプ"){
				// debugger
				let remindCheckElm = document.getElementById("remindCheck")
				let refineCheckElm = document.getElementById("refineCheck")
				if(remindCheckElm.checked && !refineCheckElm.checked){
					return ["再発行"]
				}
				if(refineCheckElm.checked && !remindCheckElm.checked){
					return ["訂正版"]
				}
				if(refineCheckElm.checked && remindCheckElm.checked){
					return ["再発行訂正版"]
				}
				return [" "]
			}

			if(ZVarLabels[0] === "+自社インボイス登録番号"){
				// debugger
				let companies = await Z.searchRecord("Company_M", "Company_Cd:equals:" + ZVarLabels[1])
				let companyInfo = companies.find(c => c.Company_Cd == ZVarLabels[1])
				if(companyInfo){
					return [companyInfo.Invoice_Num]
				}else{
					return [" "]
				}
			}
			
			//サブフォームの場合
			if(ZVarLabels[0].match(/#$/g) || fieldType == "subform"){
				// debugger
				let ZvarSubformInfo = tabAllFieldsInfo.find( function(data){
					return (
						data.display_label == ZVarLabels[0].replace(/#$/,"") ||
						data.field_label == ZVarLabels[0].replace(/#$/,"") ||
						data.name == ZVarLabels[0].replace(/#$/,"") ||
						data.api_name == ZVarLabels[0].replace(/#$/,"")
					) ? data : false
				})
				if(!ZvarSubformInfo){ return [" "] }
				let results = []
				for(let subformRecord of record[ZvarSubformInfo.api_name]){
					let replaceContent = await _getDataFromZVarString(ZVarLabels.slice(1), ZvarSubformInfo.api_name, subformRecord)
					if(replaceContent == " "){ continue }
					results = results.concat( replaceContent )
				}
				return results
			}

	
			
			let ZvarFieldInfo = tabAllFieldsInfo.find( function(data){
				return (
					data.field_label == ZVarLabels[0] ||
					data.api_name == ZVarLabels[0]
				) ? data : false 
			})
			if(ZvarFieldInfo){
				// debugger
				let fieldData = record[ZvarFieldInfo.api_name]

				if((typeof fieldData === "object" && fieldData !== null ) && ZVarLabels.length > 1){
					let apiName = ZvarFieldInfo.api_name == "Owner" ? "Owner"
								: ZvarFieldInfo.api_name == "Account_Name" ? "Accounts"
								: ZvarFieldInfo.data_type == "lookup" ? ZvarFieldInfo.lookup.module.api_name
								: fieldData.module

					let aaa = await Z.getModule(apiName)
					let rRecord = await Z.getRecord(apiName, fieldData.id)
					return await _getDataFromZVarString(ZVarLabels.slice(1), apiName, rRecord)
				}
		
				if(typeof fieldData === "object" && fieldData !== null){ return [fieldData.name] }

				if(fieldData == null){ return [" "]}
				if(ZvarFieldInfo.data_type == "datetime"){ return [moment(fieldData).format("YYYY/MM/DD HH:mm:ss")] }
				if(ZvarFieldInfo.data_type == "date"){ return [moment(fieldData).format("YYYY/MM/DD")] }
				return [fieldData]	
			}else{
				return false
			}




		}
	}


	function criteriaCheck(record, fields, label){
		let filteredLabel = label.split("??")

		let criterias = []
		if(filteredLabel[1].match(/\|\|/g)){
			let orCreteria = filteredLabel[1].split("||")
			for(let c of orCreteria){
				if(c.match(/\&\&/g)){
					let andCreteria = c.split("&&")
					for(let a of andCreteria){
						criterias.push(a)
					}
				}else{
					criterias.push(c)
				}
			}
		}else{
			criterias.push(filteredLabel[1])
		}
		// debugger

		let fconv = []
		criterias.forEach( c => {
			let ff = c.split("=")
			let field = fields.find(f => f.field_label == ff[0] || f.api_name == ff[0])
			let fieldapiname = field?.api_name
			fconv.push({display: ff[0], apiname:fieldapiname, value:ff[1]})
		})
		fconv.forEach( c => {
			filteredLabel[1] = filteredLabel[1].replaceAll(c.display, `record['${c.apiname}'].toString()`)
		})
		filteredLabel[1] = filteredLabel[1].replaceAll("=", "=='").replaceAll("&&", "' && ").replaceAll("||", "' || ") + "'"

		if(eval(filteredLabel[1])){
			return true
		}else{
			return false
		}


		let criteria = filteredLabel[1].split("=")
		let criteriaField = criteria[0]
		let criteriaValue = criteria[1]

		let criteriaFieldApiName = fields.find(f => f.field_label == criteriaField || f.api_name == criteriaField)?.api_name
		if(record[criteriaFieldApiName].toString() != criteriaValue){ return false }
		return true
	}

	function replacedContentsToSheetContents(replacedContents){
		let sheetContents = replacedContents.map(function(row){
			let cols = row.row_details.map(function(col){
				if(col.content == ""){
					let templateRow = zCurrentTemplateContents.find( r => r.row_index == row.row_index )
					let templateCol = templateRow.row_details.find( c => c.column_index == col.column_index )
					let templateContent = templateCol.content
					// if(templateContent != ""){debugger}
					return {column_index:col.column_index, content:templateContent}
				}else{
					let replacedContent = col.content
					// if(replacedContent != ""){debugger}
					return {column_index:col.column_index, content:replacedContent}
				}
			})
			return {row_index:row.row_index, row_details:cols}
		})
		return sheetContents
	}

	async function getFieldInfo(module_apiName){
		if(typeof FIELDS[module_apiName] === "undefined"){
			FIELDS[module_apiName] = (await ZOHO.CRM.META.getFields({Entity:module_apiName})).fields
			return FIELDS[module_apiName]
		}else{
			return FIELDS[module_apiName]
		}
	}
	async function getRelatedListInfo(module_apiName){
		if(typeof RELATED_LISTS[module_apiName] === "undefined"){
			RELATED_LISTS[module_apiName] = (await ZOHO.CRM.META.getRelatedList({Entity:module_apiName})).related_lists
			return RELATED_LISTS[module_apiName]
		}else{
			return RELATED_LISTS[module_apiName]
		}
	}

	function loadHtml(_html,replace="widget"){
		var xmlhttp = new XMLHttpRequest();
		xmlhttp.open("GET",_html + "?r=" + Date.now(),true);
		xmlhttp.onreadystatechange = function(){
			//とれた場合Idにそって入れ替え
			if(xmlhttp.readyState == 4 && xmlhttp.status==200){
				var data = xmlhttp.responseText;
				var elem = document.getElementById(replace);
				elem.innerHTML = data;
				return data;
			}
		}
		xmlhttp.send(null);
	}
	function waitFor(selector) {
		return new Promise(function (res, rej) {
			waitForElementToDisplay(selector, 200);
			function waitForElementToDisplay(selector, time) {
				if (document.querySelector(selector) != null) {
					res(document.querySelector(selector));
				}
				else {
					setTimeout(function () {
						waitForElementToDisplay(selector, time);
					}, time);
				}
			}
		});
	}

	function initPaymentUI(){
		let genBtn = document.getElementById("pay-generateBtn")
		genBtn.addEventListener("click", function(){
			zSheetTemplate = document.getElementById("payTemplateSelect").value
			createPaymentViaSheets(widgetData)
		})
	}

	function initEstimateUI(){
		let genBtn = document.getElementById("est-generateBtn")
		genBtn.addEventListener("click", function(){
			// debugger
			zSheetTemplate = document.getElementById("estTemplateSelect").value
			createEstimateViaSheets(widgetData)
		})
	}

	function initInvoiceUI(){
		let genBtn = document.getElementById("generateBtn")
		// let genGatherBtn = document.getElementById("generateGatherBtn")
		// genBtn.addEventListener("click", function(){ createInvoice(data) })
		genBtn.addEventListener("click", function(){
			zSheetTemplate = document.getElementById("invTemplateSelect").value
			createInvoiceViaSheets(widgetData)
		})
		// genGatherBtn.addEventListener("click", function(){ gather = true; createInvoiceViaSheets(widgetData) })

		let nostampElm = document.getElementById("nostamp")
		let remindElm = document.getElementById("remindCheck")
		let refineElm = document.getElementById("refineCheck")
		// nostampElm.addEventListener("change", function(){
		// 	if(nostampElm.checked){
		// 		document.getElementById("remindReasonArea").style.display = "none"
		// 	}
		// })
		remindElm.addEventListener("change", function(){
			if(remindElm.checked){
				document.getElementById("remindReasonArea").style.display = "block"
			}
		})
		refineElm.addEventListener("change", function(){
			if(refineElm.checked){
				document.getElementById("remindReasonArea").style.display = "block"
			}
		})
		let remindReason = document.getElementById("remindReason")
	}

	async function createFolderOnWorkDrive(folderName, parentDirectoryId){
		let res = await ZOHO.CRM.FUNCTIONS.execute("f30", {
			arguments:JSON.stringify({
				mode:"createfolder",
				arg1:folderName,
				arg2:parentDirectoryId
			})
		})
		let output = JSON.parse(res.details.output)
		return res
	}

	async function createSheetFromTemplate(workbookName, templateId){
		let res = await ZOHO.CRM.CONNECTION.invoke("zohooauth",{
			"url": "https://sheet.zoho.jp/api/v2/createfromtemplate",
			"method" : "POST",
			"param_type" : 1,
			"parameters" : {
				workbook_name:workbookName,
				resource_id:templateId,
				method:"workbook.createfromtemplate"
			}
		})
		return res
	}

	async function moveFile(id, directoryId){
		let res = await ZOHO.CRM.FUNCTIONS.execute("f30", {
			arguments:JSON.stringify({
				mode:"movefile",
				arg1:id,
				arg2:directoryId
			})
		})
		// debugger
		return res
	}

	async function deleteFile(id){
		let res = await ZOHO.CRM.FUNCTIONS.execute("f30", {
			arguments:JSON.stringify({
				mode:"deletefile",
				arg1:id,
				arg2:""
			})
		})
		// debugger
		return res
	}


	async function shareFile(resourceid){
		let res = await ZOHO.CRM.FUNCTIONS.execute("f30", {
			arguments:JSON.stringify({
				mode:"share",
				arg1:resourceid
			})
		})
		return res
	}

	async function getDirList(directoryId){
		let res = await ZOHO.CRM.FUNCTIONS.execute("f30", {
			arguments:JSON.stringify({
				mode:"list",
				arg1:directoryId
			})
		})
		let dirListOutput = JSON.parse(res.details.output)
		let dirList = dirListOutput.data
		return dirList
	}

}