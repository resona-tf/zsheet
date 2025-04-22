/*
 * Version: 1.3.11
 * Update: テンプレートアイテムの並び替えアニメーションを追加
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

let SETTINGS = {
    "productionOrgId": '',
    "SheetTemplateUrl": [
        {
            "name":'',
            "url":'',
            "attachToRecord": false,
            "attachFormat": 'pdf',
            "download": false,
            "downloadFormat": 'pdf',
        }
    ]
}

// 設定の一時保存用
let TEMP_SETTINGS = JSON.parse(JSON.stringify(SETTINGS))

let gather
let widgetData
let WidgetKey

let WidgetHeight
let WidgetWidth = 900

// 設定UIの初期化と制御
function initializeSettingsUI() {
    const settingsBtn = document.getElementById('settingsBtn');
    const settingsPanel = document.getElementById('settingsPanel');
    const settingsCloseBtn = document.getElementById('settingsCloseBtn');
    const settingsSaveBtn = document.getElementById('settingsSaveBtn');
    const operationUI = document.getElementById('operation-ui');
    const progress = document.getElementById('progress');
    
    const templateList = document.querySelector('.template-list');
    const templateAdd = document.querySelector('.template-add');

    // アイテムの並び替えアニメーション
    async function moveItems(currentIndex, targetIndex) {
        const currentItem = document.getElementById(`template-item-${currentIndex}`);
        const targetItem = document.getElementById(`template-item-${targetIndex}`);
        
        // 移動方向に基づいてクラスを追加
        const isMovingUp = targetIndex < currentIndex;
        currentItem.classList.add(isMovingUp ? 'moving-up' : 'moving-down');
        targetItem.classList.add(isMovingUp ? 'moving-down' : 'moving-up');
        
        // アニメーション完了を待つ
        await new Promise(resolve => setTimeout(resolve, 300));
        
        // データの更新
        const temp = TEMP_SETTINGS.SheetTemplateUrl[currentIndex];
        TEMP_SETTINGS.SheetTemplateUrl[currentIndex] = TEMP_SETTINGS.SheetTemplateUrl[targetIndex];
        TEMP_SETTINGS.SheetTemplateUrl[targetIndex] = temp;
        
        // リストを再描画
        renderTemplateList();
    }

    // テンプレート一覧の生成
    function renderTemplateList() {
        templateList.innerHTML = '';
        TEMP_SETTINGS.SheetTemplateUrl.forEach((template, index) => {
            const item = document.createElement('div');
            item.className = 'template-item';
            item.id = `template-item-${index}`;
            item.innerHTML = `
            <div class="itemsettings">
                <div class="itemsettings-template">
                    <input type="text" placeholder="テンプレート名" value="${template.name || ''}" data-index="${index}" data-field="name">
                    <input type="url" placeholder="URL" value="${template.url || ''}" data-index="${index}" data-field="url">
                </div>
                <div class="itemsettings-postprocesses">
                    <div class="postprocess-label">作成後の処理</div>
                    <div class="postprocess-fields">
                        <div class="setting-field">
                            <div class="form-check">
                                <input type="checkbox" class="form-check-input" id="attachToRecord-${index}" ${template.attachToRecord ? 'checked' : ''}>
                                <label class="form-check-label" for="attachToRecord-${index}">データに添付</label>
                            </div>
                            <select class="form-select form-select-sm" id="attachFormat-${index}" style="display:${template.attachToRecord ? 'block' : 'none'};">
                                <option value="pdf" ${template.attachFormat === 'pdf' ? 'selected' : ''}>PDF</option>
                                <option value="xlsx" ${template.attachFormat === 'xlsx' ? 'selected' : ''}>Excel</option>
                            </select>
                        </div>

                        <div class="setting-field">
                            <div class="form-check">
                                <input type="checkbox" class="form-check-input" id="download-${index}" ${template.download ? 'checked' : ''}>
                                <label class="form-check-label" for="download-${index}">ダウンロード</label>
                            </div>
                            <select class="form-select form-select-sm" id="downloadFormat-${index}" style="display:${template.download ? 'block' : 'none'};">
                                <option value="pdf" ${template.downloadFormat === 'pdf' ? 'selected' : ''}>PDF</option>
                                <option value="xlsx" ${template.downloadFormat === 'xlsx' ? 'selected' : ''}>Excel</option>
                                <option value="xlsx_combined" ${template.downloadFormat === 'xlsx_combined' ? 'selected' : ''}>Excel（1ブック）</option>
                            </select>
                        </div>

                        <div class="setting-field">
                            <div class="form-check">
                                <input type="checkbox" class="form-check-input" id="open-${index}" ${template.open ? 'checked' : ''}>
                                <label class="form-check-label" for="open-${index}">印刷用に開く</label>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="template-sort-buttons">
                    <div class="template-sort-button template-sort-up" data-index="${index}">
                        <span class="material-icons">keyboard_arrow_up</span>
                    </div>
                    <div class="template-sort-button template-sort-down" data-index="${index}">
                        <span class="material-icons">keyboard_arrow_down</span>
                    </div>
                </div>
            </div>
            <span class="material-icons template-remove" data-index="${index}">remove_circle_outline</span>
            `;
            templateList.appendChild(item);

            // 入力値の変更を監視
            item.querySelectorAll('input[type="text"], input[type="url"]').forEach(input => {
                input.addEventListener('change', (e) => {
                    const index = parseInt(e.target.dataset.index);
                    const field = e.target.dataset.field;
                    TEMP_SETTINGS.SheetTemplateUrl[index][field] = e.target.value;
                });
            });

            // チェックボックスの変更を監視
            const attachToRecord = document.getElementById(`attachToRecord-${index}`);
            const attachFormat = document.getElementById(`attachFormat-${index}`);
            const download = document.getElementById(`download-${index}`);
            const downloadFormat = document.getElementById(`downloadFormat-${index}`);
            const openSheet = document.getElementById(`open-${index}`);

            // データに添付の制御
            attachToRecord.addEventListener('change', (e) => {
                TEMP_SETTINGS.SheetTemplateUrl[index].attachToRecord = e.target.checked;
                attachFormat.style.display = e.target.checked ? 'block' : 'none';
            });

            // 添付形式の制御
            attachFormat.addEventListener('change', (e) => {
                TEMP_SETTINGS.SheetTemplateUrl[index].attachFormat = e.target.value;
            });

            // ダウンロードの制御
            download.addEventListener('change', (e) => {
                TEMP_SETTINGS.SheetTemplateUrl[index].download = e.target.checked;
                downloadFormat.style.display = e.target.checked ? 'block' : 'none';
            });

            // ダウンロード形式の制御
            downloadFormat.addEventListener('change', (e) => {
                TEMP_SETTINGS.SheetTemplateUrl[index].downloadFormat = e.target.value;
            });

            // オープンの制御
            openSheet.addEventListener('change', (e) => {
                TEMP_SETTINGS.SheetTemplateUrl[index].open = e.target.checked;
            });

            // 削除ボタンの処理
            item.querySelector('.template-remove').addEventListener('click', (e) => {
                const index = parseInt(e.target.dataset.index);
                TEMP_SETTINGS.SheetTemplateUrl.splice(index, 1);
                if (TEMP_SETTINGS.SheetTemplateUrl.length === 0) {
                    TEMP_SETTINGS.SheetTemplateUrl.push({
                        "name":'',
                        "url":'',
                        "attachToRecord": false,
                        "attachFormat": 'pdf',
                        "download": false,
                        "downloadFormat": 'pdf',
                    });
                }
                renderTemplateList();
            });

            // 並び替えボタンの処理
            item.querySelector('.template-sort-up').addEventListener('click', async (e) => {
                const index = parseInt(e.target.closest('.template-sort-up').dataset.index);
                if (index > 0) {
                    await moveItems(index, index - 1);
                }
            });

            item.querySelector('.template-sort-down').addEventListener('click', async (e) => {
                const index = parseInt(e.target.closest('.template-sort-down').dataset.index);
                if (index < TEMP_SETTINGS.SheetTemplateUrl.length - 1) {
                    await moveItems(index, index + 1);
                }
            });
        });
    }

    // テンプレート追加ボタンの処理
    templateAdd.addEventListener('click', () => {
        TEMP_SETTINGS.SheetTemplateUrl.push({
            "name":'',
            "url":'',
            "attachToRecord": false,
            "attachFormat": 'pdf',
            "download": false,
            "downloadFormat": 'pdf',
            "open": false,
        });
        renderTemplateList();
    });

    // 設定パネルを開く
    function openSettingsPanel() {
        ZOHO.CRM.UI.Resize({height: "700", width: WidgetWidth.toString()})
        // 一時保存用の設定をコピー
        TEMP_SETTINGS = JSON.parse(JSON.stringify(SETTINGS));
        settingsPanel.style.display = 'block';
        operationUI.style.display = 'none';
        settingsBtn.style.display = 'none';
        renderTemplateList();
    }

    // 設定パネルを閉じる
    function closeSettingsPanel() {
        settingsPanel.style.display = 'none';
        operationUI.style.display = 'flex';
        settingsBtn.style.display = 'flex';
        ZOHO.CRM.UI.Resize({height: WidgetHeight.toString(), width: WidgetWidth.toString()})
    }

    // 設定を保存
    async function saveSettings() {
        SETTINGS = JSON.parse(JSON.stringify(TEMP_SETTINGS));
        await saveWidgetSettings(WidgetKey);
        closeSettingsPanel();
    }

    // 設定パネルの開閉制御
    settingsBtn.addEventListener('click', openSettingsPanel);
    settingsCloseBtn.addEventListener('click', closeSettingsPanel);
    settingsSaveBtn.addEventListener('click', saveSettings);
}

async function loadWidgetSettings(key){
    let settingData = await getOrgVariable(`widget_${key}`)
    if(!settingData){
        const orgInfo = await ZOHO.CRM.CONFIG.getOrgInfo()
        SETTINGS.productionOrgId = orgInfo.org[0].zgid
        await saveWidgetSettings(key)
        return
    }else{
        SETTINGS = JSON.parse(settingData)
        if (!SETTINGS.SheetTemplateUrl) {
            SETTINGS.SheetTemplateUrl = [{name:'', url:''}];
        }
        // 古い設定の移行
        SETTINGS.SheetTemplateUrl.forEach(template => {
            if (template.createAsSheets !== undefined) {
                delete template.createAsSheets;
            }
            if (template.attachFormat === 'xlsx') {
                template.attachFormat = 'xlsx';
            }
            if (template.downloadFormat === 'xlsx') {
                template.downloadFormat = 'xlsx';
            }
            if (template.attachFormat === 'url') {
                template.attachFormat = 'zohosheet';
            }
            if (template.downloadFormat === 'url') {
                template.downloadFormat = 'zohosheet';
            }
            // 添付形式のcombinedオプションを通常のフォーマットに変換
            if (template.attachFormat === 'excel_combined') {
                template.attachFormat = 'xlsx_combined';
            }
            if (template.attachFormat === 'zohosheet_combined') {
                template.attachFormat = 'zohosheet';
            }
        });
        // 一時保存用の設定を初期化
        TEMP_SETTINGS = JSON.parse(JSON.stringify(SETTINGS));
        initializeSettingsUI();
    }
}

async function saveWidgetSettings(key){
    let settingData = await getOrgVariable(`widget_${key}`)
    if(!settingData){
        await createOrgVariables(`widget_${key}`)
        settingData = await getOrgVariable(`widget_${key}`)
    }
    await updateOrgVariavbles(`widget_${key}`, SETTINGS)
}

async function getOrgVariable(key){
    // let result = await ZOHO.CRM.CONNECTION.invoke("zohooauth", {
    //     "url":`${ApiDomain}/crm/v7/settings/variables`,
    //     "method" : "GET",
    // })
    // let variables = result.details.statusMessage.variables.find((v) => v.name == key)

    let variables = await _getOrgVariable(key)

    let widgetSettingsData = await ZOHO.CRM.API.searchRecord({Entity:"Widget_Settings",Type:"criteria",Query:`(Name:equals:${key})`})
    if(!widgetSettingsData?.data){
        SETTINGS = {}
        return null
    }else{
        return widgetSettingsData?.data[0].Json
    }
}

async function _getOrgVariable(key){
    let result = await ZOHO.CRM.CONNECTION.invoke("zohooauth", {
        "url":`${ApiDomain}/crm/v7/settings/variables`,
        "method" : "GET",
    })
    let variables = result.details.statusMessage.variables.find((v) => v.name == key)
    if(variables?.value){
        return variables?.value
    }else{
        return null
    }
}

async function createOrgVariables(key){
    let result = await ZOHO.CRM.API.insertRecord({
        Entity:"Widget_Settings",
        APIData:{
            Name:key,
            Json:JSON.stringify(SETTINGS)
        }
    })
}

async function updateOrgVariavbles(key,val){
    debugger
    let widgetSettingsData = await ZOHO.CRM.API.searchRecord({Entity:"Widget_Settings",Type:"criteria",Query:`(Name:equals:${key})`})
    let result = await ZOHO.CRM.API.updateRecord({
        Entity:"Widget_Settings",
        APIData:{
            id:widgetSettingsData.data[0].id,
            Name:key,
            Json:JSON.stringify(SETTINGS)
        }
    })
}

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

let PRDUCTION_ORGID
let TEMPLATE_CRMVAR
let ENVIROMENT = "production"
let ApiDomain = "https://www.zohoapis.jp"
let fileNameAddition = ""

let WORKING_BOOK_ID

let GatherSalesNumbers = ""

const API_COUNT = {}
const LAST_API_CALL = {};
const apiUsageLogs = {};
const apiLimits = {
    "worksheet.content.get": 120,
    "worksheet.list": 60,
    "worksheet.rows.delete": 20,
    "worksheet.csvdata.set": 20,
    "worksheet.copy": 30,
    "worksheet.delete": 20,
    "workbook.download": 30,
    "workbook.share": 20,
};

let IP
