window.onload = async function(){
	ZOHO.embeddedApp.on("PageLoad", async function(data){
		debugger
		PRDUCTION_ORGID = await _getOrgVariable("Production_Org_Id")
		WidgetKey = `${data.Entity}_${data.ButtonPosition}`
		const orgInfo = await ZOHO.CRM.CONFIG.getOrgInfo()
		const orgId = orgInfo.org[0].zgid
		ApiDomain = orgId == PRDUCTION_ORGID ? "https://www.zohoapis.jp" : "https://crmsandbox.zoho.jp"
		ENVIROMENT = orgId == PRDUCTION_ORGID ? "production" : "sandbox"
		fileNameAddition = orgId == PRDUCTION_ORGID ? "" : "（テスト）"
		widgetData = data

		await loadWidgetSettings(WidgetKey)
		TEMPLATE_CRMVAR = 
		await waitFor("#loadCheck")

		// debugger
		document.querySelector("#operation-ui").style.display = "block"

		templateSelectoerSetup(SETTINGS.SheetTemplateUrl, document.querySelector("#invTemplateSelect"))

		function templateSelectoerSetup(v, elm){
			let templateCheckboxHtml = ""
			for(let entry of v){
				const id = `template-${v.indexOf(entry)}`
				templateCheckboxHtml += `
					<div class="template-checkbox-item">
						<div class="form-check">
							<input class="form-check-input" type="checkbox" 
								value="${entry.url}" 
								data-index="${v.indexOf(entry)}"
								id="${id}">
							<label class="form-check-label" for="${id}">
								${entry.name}
							</label>
						</div>
					</div>`
			}
			elm.innerHTML = templateCheckboxHtml
			WidgetHeight = 60 + elm.offsetTop + elm.offsetHeight

			ZOHO.CRM.UI.Resize({height: WidgetHeight.toString(), width: WidgetWidth.toString()})

			// v1.0.0: チェックボックスの変更イベントを監視
			elm.addEventListener('change', function() {
				const hasChecked = Array.from(elm.querySelectorAll('input[type="checkbox"]'))
					.some(checkbox => checkbox.checked);
				document.getElementById("generateBtn").disabled = !hasChecked;
			});

			// 初期状態でボタンを無効化
			document.getElementById("generateBtn").disabled = true;
		}


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

	ZOHO.embeddedApp.init()

	function addWorkbookLink(workbookName, workbookUrl) {
		const linksContainer = document.getElementById('workbookLinks')
		linksContainer.style.display = 'flex'
		const link = document.createElement('a')
		link.href = workbookUrl
		link.target = '_blank'
		link.className = 'workbook-link'
		link.textContent = workbookName
		linksContainer.appendChild(link)

		// リンクが追加されるたびにUIのサイズを調整
		// debugger
		WidgetHeight = 24 + linksContainer.offsetTop + linksContainer.offsetHeight
		ZOHO.CRM.UI.Resize({height: WidgetHeight.toString(),width: WidgetWidth.toString()})
	}

	function completeProgress() {
		const progressBar = document.getElementById('progressBar')
		// アニメーションクラスを削除
		progressBar.classList.remove('progress-bar-animated')
		progressBar.classList.remove('progress-bar-striped')
		// プログレスバーを完了状態に
		progressBar.style.width = '100%'
		progressBar.setAttribute('aria-valuenow', '100')
		
		// 作成ボタンのスピナーを停止し、無効化状態を維持
		const generateBtnInProgress = document.getElementById('generateBtnInProgress')
		generateBtnInProgress.classList.remove('spinner-border')
		generateBtnInProgress.classList.remove('spinner-border-sm')
		generateBtnInProgress.innerHTML = '完了'
		document.querySelector('.button-progress').style.display = 'none'
	}

	async function createZohoSheetDocuments(data){

		document.getElementById("generateBtnText").style.display = "none"
		document.getElementById("generateBtnInProgress").style.display = "block"
		document.getElementById("generateBtn").setAttribute("disabled", true)

		let proceesOrder = []

		// 選択されたテンプレートの設定を取得
		const templateContainer = document.getElementById("invTemplateSelect")
		const selectedTemplates = Array.from(templateContainer.querySelectorAll('input[type="checkbox"]:checked'))
			.map(checkbox => ({
				index: checkbox.dataset.index,
				url: checkbox.value
			}))

		if (selectedTemplates.length === 0) {
			throw new Error('テンプレートが選択されていません')
		}

		// 処理内容を生成する
		let recordData 
		for (const template of selectedTemplates) {
			let combinedDataSkip = false
			let openDataSkip = false
			for(let idx in data.EntityId){
				recordData = await Z.getRecord(ENTITY, data.EntityId[idx])
				const templateSettings = SETTINGS.SheetTemplateUrl[template.index]
				if(templateSettings?.attachToRecord){
					proceesOrder.push({
						type:"attach",
						format: templateSettings?.attachFormat,
						recordIds: [data.EntityId[idx]],
						templateUrl: template.url,
						name:`${templateSettings?.name}_${recordData.Name}`,
						templateName: templateSettings?.name,
					})
				}
				if(templateSettings?.download){
					if(templateSettings?.downloadFormat.includes("combined")){
						if(!combinedDataSkip){
							proceesOrder.push({
								type:"download",
								format: templateSettings?.downloadFormat,
								recordIds: data.EntityId,
								templateUrl: template?.url,
								name:`${templateSettings?.name}_${moment().format("YYYYMMDD-HHmmss")}`,
								templateName: templateSettings?.name,
							})
							combinedDataSkip = true
						}
					}else{
						proceesOrder.push({
							type:"download",
							format: templateSettings?.downloadFormat,
							recordIds: [data.EntityId[idx]],
							templateUrl: template.url,
							name:`${templateSettings?.name}_${recordData.Name}`,
							templateName: templateSettings?.name,
						})
					}
				}
				if(templateSettings?.open){
					if(!openDataSkip){
						proceesOrder.push({
							type:"open",
							format:"combined",
							recordIds: data.EntityId,
							templateUrl: template.url,
							name:`${templateSettings?.name}_${moment().format("YYYYMMDD-HHmmss")}`,
							templateName: templateSettings?.name,
						})
						openDataSkip = true
					}
				}
			}
		}
		proceesOrder.sort((a, b) => a.type.localeCompare(b.type))
		// プログレスバーの初期化
		initProgress(proceesOrder.length)
		// debugger

		let currentUser = await ZOHO.CRM.CONFIG.getCurrentUser()

		for(let idx in proceesOrder){
			progressNext()

			let WorkbookName,createBookRes
			if(proceesOrder[idx].format.includes("combined")){
				WorkbookName = proceesOrder[idx].templateName
			}else{
				WorkbookName = proceesOrder[idx].name
			}
			
			createBookRes = await createSheetFromTemplate(WorkbookName, proceesOrder[idx].templateUrl)
			WORKING_BOOK_ID = createBookRes.details.statusMessage.resource_id
			await ZS.shareWorkbook(WORKING_BOOK_ID, [{"user_email":currentUser.users[0].email,"access_level":"share"}])
			await generateSheet(WORKING_BOOK_ID, ENTITY, proceesOrder[idx].recordIds)
			proceesOrder[idx].sheetUrl = createBookRes.details.statusMessage.workbook_url

			if(proceesOrder[idx].type == "attach" || proceesOrder[idx].type == "download"){
				let ext = proceesOrder[idx].format
				if(ext.includes("combined")){ext = ext.replace("_combined","")}
				let blob = await ZS.downloadAs(WORKING_BOOK_ID,ext)


				if(proceesOrder[idx].type == "attach"){
					// debugger
					await Z.attachFile(ENTITY, proceesOrder[idx].recordIds[0], WorkbookName+"."+ext, blob)
				}

				if(proceesOrder[idx].type == "download"){
					const url = window.URL.createObjectURL(blob);
					const a = document.createElement("a");
					a.href = url;
					a.download = WorkbookName+"."+ext;
					a.click();
					window.URL.revokeObjectURL(url);
				}
			}
			if(proceesOrder[idx].type == "open"){
				addWorkbookLink(WorkbookName, proceesOrder[idx].sheetUrl)
			}
		}


		// 処理完了時の表示更新
		completeProgress()
		document.getElementById("closeBtnArea").style.display = "flex"
		document.getElementById("closeBtn").addEventListener("click", function() {
			ZOHO.CRM.UI.Popup.close()
		})
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
		genBtn.addEventListener("click", async function(){
			try{
				await createZohoSheetDocuments(widgetData)
			} catch (error) {
				console.log(error)
				alert(JSON.stringify(error))
			}
		})
	}
}
