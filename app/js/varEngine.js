/**
 * シート内の変数を実際の値に置換する関数
 * @param {string} workbookId - 対象のワークブックID
 * @param {string} sheetId - 置換後のデータを書き込む対象シートID
 * @param {string} tempSheetId - 置換前のテンプレートが含まれるシートID
 * @param {string} moduleApiName - Zohoモジュール名
 * @param {string} recordId - 置換に使用するレコードID
 * @returns {Array} - 置換後のシートコンテンツ
 */
async function replaceSheetVariables(workbookId, sheetId, tempSheetId, moduleApiName, recordId){
	// テンプレートシートの内容を取得
	let contents = await ZS.getSheetContents(workbookId, tempSheetId)
	
	// 置換後のデータを格納する配列を初期化
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
			// 1列目はスキップ
			if(colInfo.column_index == 1){ continue }
			// 置換変数(<Z%...%Z>)が含まれていない場合はスキップ
			if( ! content.match(/<Z%[^%]*%Z>/g) ){ continue }
			// debugger
			
			// Zohoフィールド変数を実際の値に置換
			let values = await replaceZohoFieldVariables( content, moduleApiName, recordId )
			if(values.length == 0){ values = [" "] }
			
			// 置換後の値を対応するセルに設定
			for(idx in values){
				// if(values[idx] == " "){ continue }
				let colNum = colInfo.column_index
				let rowNum = rowInfo.row_index
				
				// 複数レコードの場合、同じリピートキーを持つ行を探す
				if(idx > 0){
					// 現在の行の1列目の値（リピートキー）を取得
					let currentRepeatKey = rowContents[0]
					// 同じリピートキーを持つ次の行を探す
					let nextRowWithSameKey = contents.find(r => 
						r.row_index > rowNum && 
						r.row_details[0].content === currentRepeatKey
					)
					if(nextRowWithSameKey){
						rowNum = nextRowWithSameKey.row_index
					}else{
						continue // 同じリピートキーを持つ行がない場合はスキップ
					}
				}
				
				let targetRow = replaceContents.find( (row) => row.row_index == rowNum )
				if(targetRow){
					let targetCol = targetRow.row_details.find( (col) => col.column_index == colNum ).content = values[idx]
				}
			}
		}
	}

	// 置換後のデータをCSV形式に変換
	// ダブルクオーテーション、カンマ、改行コードをエスケープ処理
	let csvData = replaceContents.map( (row) => row.row_details.map( (col) => `"${col.content.replace(/"/g, '""').replace(/,/g, '\,').replace(/\n/g, '\r')}"` ).join(",") ).join("\n")
	// debugger

	// CSVデータで対象シートを更新
	await ZS.updateSheetViaCsv(workbookId, sheetId, csvData)
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
