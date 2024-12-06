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
	},
	attachFile: async function(entity, id, fileName, fileBlob){
		apiCounter("ZOHO.CRM.API.attachFile")
		try {
			let res = await ZOHO.CRM.API.attachFile({Entity:entity, RecordID:id, File:{Name:fileName, Content:fileBlob}})
		} catch (error) {
			console.log(error)
			// alert(JSON.stringify(error))
		}
		return result
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
						}
						if(result.details.statusMessage.error_code == "2866"){
							return
						}

						debugger
						alert("エラーが発生しました。再度実行してください。")
						ZOHO.CRM.UI.Popup.close()
						return
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

		//使用範囲のコンテンツを取得
		await apiCounter("worksheet.content.get")
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
		await apiCounter("worksheet.list")
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
		await apiCounter("worksheet.rows.delete")
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
		await apiCounter("worksheet.csvdata.set")
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
		await apiCounter("worksheet.copy")
		// console.log(`worksheet.copy -> ${wbid}.${origWsid}.${newWsName}`)
		let result = await ZS.zsApi(
			`https://sheet.zoho.jp/api/v2/${wbid}`,"POST",1,
			{
				method:"worksheet.copy",
				worksheet_id:origWsid,
				new_worksheet_name:newWsName
			}
		)
		let newWs = result.worksheet_names.find( (ws) => ws.worksheet_name == result.new_worksheet_name )
		ZS.sheetContents[wbid][newWs.worksheet_id] = ZS.sheetContents[wbid][origWsid]
		ZS.copySheetApiCount++
		return result
	},
	deleteSheet: async function(wbid,wsid){
		ZS.sheetContents[wbid][wsid] = null
		await apiCounter("worksheet.delete")
		// console.log(`worksheet.delete -> ${wbid}.${wsid}`)
		let result = await ZS.zsApi(
			`https://sheet.zoho.jp/api/v2/${wbid}`,"POST",1,
			{
				method:"worksheet.delete",
				worksheet_id:wsid
			}
		)
		return result
	},
	downloadAs: async function(wbid,format){
		await apiCounter("workbook.download")
		let mergedPdf = await PDFLib.PDFDocument.create()
		let pdf = await ZOHO.CRM.CONNECTION.invoke("zohooauth",{
			"url": `https://sheet.zoho.jp/api/v2/download/${wbid}?method=workbook.download`,
			"method" : "POST",
			"param_type" : 1,
			"parameters" :{
				// method:"workbook.download",
				format:format,
				page_settings:{
					print_type:"WORKBOOK",
					scale:3,
					add_gridlines:false,
					add_button:false,
					add_image:true,
					align_vcenter:false,
					align_hcenter:true,
					margin_left:0,
					margin_right:0,
					margin_top:0.25,
					margin_bottom:0.25,
				}
			},
			"RESPONSE_TYPE":"stream"
		})
		let pdfU8Array = await pdf.arrayBuffer()
		let blob = new Blob([pdfU8Array], { type: `application/${format}` });
		return blob
	}
}

async function apiCounter(api) {
    const now = Date.now();
    const limit = apiLimits[api];

    if (!limit) {
        console.warn(`No rate limit found for API: ${api}`);
        return;
    }

    if (!apiUsageLogs[api]) {
        apiUsageLogs[api] = [];
    }

    // 直近60秒以内の呼び出し記録を保持し、古い記録を削除
    apiUsageLogs[api] = apiUsageLogs[api].filter(timestamp => now - timestamp < 60000);

    if (apiUsageLogs[api].length >= limit) {
        const oldestCallTimestamp = apiUsageLogs[api][0];
        const waitTime = 60000 - (now - oldestCallTimestamp);
        console.log(`Waiting ${waitTime}ms for API: ${api} to avoid rate limit`);
        await new Promise(resolve => setTimeout(resolve, waitTime));
        return apiCounter(api); // 再帰的に呼び出すことで再度チェック
    }

    apiUsageLogs[api].push(now);
    console.log(`${api} : ${apiUsageLogs[api].length}/${limit}`);
}

async function createSheetFromTemplate(workbookName, templateUrl){
	// Initialize an empty `templateId`
	let templateId = '';

	// Check if the `templateUrl` contains a '/'
	if (templateUrl.includes('/')) {
		// Extract the part after the last '/' using `split` and `pop`
		let parts = templateUrl.split('/');
		let lastPart = parts.pop();

		// Now, check if the extracted part contains a '?'
		if (lastPart.includes('?')) {
			// Extract the part before the '?' using `split` again
			templateId = lastPart.split('?')[0];
		} else {
			// If there is no '?', use the entire last part
			templateId = lastPart;
		}
	}

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

async function getRelatedListInfo(module_apiName){
	if(typeof RELATED_LISTS[module_apiName] === "undefined"){
		RELATED_LISTS[module_apiName] = (await ZOHO.CRM.META.getRelatedList({Entity:module_apiName})).related_lists
		return RELATED_LISTS[module_apiName]
	}else{
		return RELATED_LISTS[module_apiName]
	}
}