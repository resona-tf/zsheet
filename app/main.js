window.onload = async function(){
	ZOHO.embeddedApp.on("PageLoad", async function(data){
		let orgInfo = await ZOHO.CRM.CONFIG.getOrgInfo()
		const orgId = orgInfo.org[0].zgid
		ApiDomain = orgId == PRDUCTION_ORGID ? "https://www.zohoapis.jp" : "https://crmsandbox.zoho.jp"

		fileNameAddition = orgId == PRDUCTION_ORGID ? "" : "（テスト）"
		widgetData = data

		await waitFor("#loadCheck")

		let templateUrl

		// debugger
		document.querySelector("#operation-ui").style.display = "block"

		templateUrl = await ZOHO.CRM.API.getOrgVariable(TEMPLATE_CRMVAR)
		if(!templateUrl.Success){
			alert("テンプレートURLが設定されていません。")
			ZOHO.CRM.UI.Popup.close()
			return
		}
		templateSelectoerSetup(templateUrl.Success.Content, document.querySelector("#invTemplateSelect"))


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
		
		ZOHO.CRM.UI.Resize({height:"200", width:"540"})

		ENTITY = data.Entity
		ENTITY_IDS = data.EntityId

		FIELDS[ENTITY] = await Z.getFields(ENTITY)
		
		//ENTITYのモジュール情報を取得
		let modules = await Z.getAllModules()
		for(let idx in modules){
			if(modules[idx].api_name == ENTITY){
				MODULE = modules[idx]
				break
			}
		}
		
		initOperationUI()

	})

	ZOHO.embeddedApp.init();
	// loadHtml("https://air.southernwave.net/ZohoWidgetDevelopment/RecordToExcel/widget.html")



	async function createZohoSheetDocuments(data){
		debugger
		progressAddStep(1)
		progressAddStep(data.EntityId.length)

		// labels = "<Z%日程->日程名%Z>"
		// let rrr = replaceZohoFieldVariables(labels, "Bookings", "10796000013735411")
		// return

		document.getElementById("generateBtnText").style.display = "none"
		document.getElementById("generateBtnInProgress").style.display = "block"
		document.getElementById("generateBtn").setAttribute("disabled", true)


		progressNext()

		//Sheetテンプレートから新規作成
		// debugger
		for(let idx in data.EntityId){
			let recordData = await Z.getRecord(ENTITY, data.EntityId[idx])
			if(IP == "61.200.96.103"){ fileNameAddition += "TEST" }
			let WorkbookName = `${recordData.Name}_${recordData.field5.name}`

			let createBookRes = await createSheetFromTemplate(WorkbookName, zSheetTemplate)
			createdWorkbookId = createBookRes.details.statusMessage.resource_id
			workbookUrl = createBookRes.details.statusMessage.workbook_url

			WORKING_BOOK_ID = createdWorkbookId

			//SheetをWorkDriveに移動
			// if(IP != "61.200.96.103"){
				// レプロ
				// let mvRes = await moveFile(createdWorkbookId, directoryId)
			// }
			
			let worksheets = await ZS.getWorksheetList(createdWorkbookId)
			//debugger
			await generateSheet(WORKING_BOOK_ID, ENTITY, [data.EntityId[idx]])
			
			// contents = await replaceSheetVariables(createdWorkbookId, worksheets[0].worksheet_id, "Gyara_Payment", record.id)
			// await clearingRows(createdWorkbookId, worksheets[0].worksheet_id)
			// progressNext()

			// return workbookUrl
			window.open(workbookUrl, "_blank")
		}
	
		ZOHO.CRM.UI.Popup.close()

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


	function initOperationUI(){
		let genBtn = document.getElementById("generateBtn")
		genBtn.addEventListener("click", function(){
			zSheetTemplate = document.getElementById("invTemplateSelect").value
			createZohoSheetDocuments(widgetData)
		})
	}

}