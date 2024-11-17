async function generateSheet(workbookid, moduleApiName, recordId = [], forceGather = false){
	let maxCol, maxRow
	let currentWsName, currentWsId
	let contents

	let gatherBaseSheetId
	let gatherBaseRecordId

	let sheetInfos = []

	let worksheets = await ZS.getWorksheetList(WORKING_BOOK_ID)
	// let singleTemplate = await ZS.copySheet(WORKING_BOOK_ID, worksheets[0].worksheet_id, "single")
	// await clearingRows(WORKING_BOOK_ID, singleTemplate.worksheet_names.find(s => s.worksheet_name == singleTemplate.new_worksheet_name).worksheet_id)

	worksheets = await ZS.getWorksheetList(WORKING_BOOK_ID, true)

	zSingleTemplateSheetId = worksheets[0].worksheet_id
	zSingleTemplateContents = await ZS.getSheetContents(WORKING_BOOK_ID, zSingleTemplateSheetId, true)

	//プログレスバーの設定
	perSheetProgressStep = progress_steps / recordId.length

	//テンプレートシートのIDを取得
	//zTemplateSheetId = worksheets[0].worksheet_id


	for( rId of recordId ){
		let result
		let record = await Z.getRecord(moduleApiName, rId)
		let newSheetName

		newSheetName = `${record.Name}-${record.field5.name}`

		//テンプレートにするシートの選択
		let templateSheetId

		templateSheetId = zSingleTemplateSheetId
		zCurrentTemplateContents = zSingleTemplateContents

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
	

	//雛形シートを削除
	await ZS.deleteSheet(WORKING_BOOK_ID, zSingleTemplateSheetId)
	// await ZS.deleteSheet(WORKING_BOOK_ID, zMultiTemplateSheetId)
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