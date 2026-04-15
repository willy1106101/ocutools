let tplArrayBuffer = null;
let fetchtemp_name = "template_1150415";

// --- Excel 匯入功能 (確保資料完整填入) ---
async function importData() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx, .xls';
    input.onchange = e => {
        const file = e.target.files[0];
        const reader = new FileReader();
        reader.onload = event => {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];

                // 1. 定義「中文表頭」轉「英文 Key」的對照表 (必須與匯出時完全一致)
                const chineseToEng = {
                    "鐘點費": "s_hour", "出席費": "s_attend", "專任助理": "s_full",
                    "兼任助理": "s_part", "工讀費": "s_work", "審查費(一般)": "s_rv_g",
                    "主持費": "s_host", "交通費(定價)": "s_rv_gb", "交通費定價金額": "s_rv_es",
                    "薪資其它勾選": "s_50_other", "薪資其它註明": "s_50_text",
                    "演講費": "s_speech", "稿費": "s_write", "審查費(教師)": "s_rv_t",
                    "論文指導費": "s_paper", "執行業務其它勾選": "s_9a_other", "執行業務其它註明": "s_9a_text",
                    "競賽獎金": "s_award", "論文交通費": "s_f_trf", "獎助生": "s_f_stu",
                    "免稅交通費勾選": "s_f_gb", "免稅交通費金額": "s_f_gb_es",
                    "免稅其它勾選": "s_free_other", "免稅其它註明": "s_free_text",
                    "權利金": "s_royalty",
                    "姓名": "name", "日期": "date", "身分證字號": "idCard",
                    "銀行帳號": "acc", "戶籍地址": "addr", "專職機構": "org",
                    "職銜": "job", "月份": "month", "是否兼任": "isPT",
                    "總額(A)": "amt", "所得稅(B)": "tB", "補保(C)": "hC",
                    "勞保(D)": "tD", "健保(E)": "tE", "勞退(F)": "tF",
                    "電話": "phone", "銀行名稱": "bN", "分行名稱": "bB"
                };

                // 2. 將 Excel 原始資料 (中文 Key) 轉換回 英文 Key
                const rawData = XLSX.utils.sheet_to_json(worksheet);
                const importedData = rawData.map(row => {
                    let engRow = {};
                    for (let chKey in row) {
                        const engKey = chineseToEng[chKey] || chKey; // 找不到對應就用原名
                        engRow[engKey] = row[chKey];
                    }
                    return engRow;
                });

                if (importedData.length > 0) {
                    // --- 以下邏輯維持不變，但資料現在已經是英文 Key 了 ---
                    const checkFields = [
                        's_hour', 's_attend', 's_full', 's_part', 's_work', 's_rv_g',
                        's_host', 's_rv_gb', 's_50_other', 's_speech', 's_write',
                        's_rv_t', 's_paper', 's_9a_other', 's_award', 's_f_trf',
                        's_f_stu', 's_f_gb', 's_free_other', 's_royalty'
                    ];
                    const textFields = ['s_rv_es', 's_50_text', 's_9a_text', 's_f_gb_es', 's_free_text'];

                    const firstRow = importedData[0];
                    checkFields.forEach(id => {
                        const el = document.getElementById(id);
                        if (el) el.checked = (firstRow[id] === "■");
                    });
                    textFields.forEach(id => {
                        const el = document.getElementById(id);
                        if (el) el.value = firstRow[id] || "";
                    });

                    const tbody = document.getElementById('dataList');
                    tbody.innerHTML = "";

                    const tableRows = importedData.filter(d => d.name || d.idCard);
                    tableRows.forEach(data => {
                        const tr = document.createElement('tr');
                        tr.innerHTML = `
                            <td><input type="text" class="name" value="${data.name || ''}"></td>
                            <td><input type="date" class="row-date" value="${data.date || ''}"></td>
                            <td><input type="text" class="idCard" value="${data.idCard || ''}" maxlength="10"></td>
                            <td><input type="text" class="acc" value="${data.acc || ''}" maxlength="14"></td>
                            <td><input type="text" class="addr" value="${data.addr || ''}"></td>
                            <td><input type="text" class="org" value="${data.org || ''}"></td>
                            <td><input type="text" class="job" value="${data.job || ''}"></td>
                            <td><input type="text" class="month" value="${data.month || ''}"></td>
                            <td><select class="isPT">
                                <option value="否" ${data.isPT === "否" ? "selected" : ""}>否</option>
                                <option value="是" ${data.isPT === "是" ? "selected" : ""}>是</option>
                            </select></td>
                            <td><input type="number" class="amt" value="${data.amt || ''}"></td>
                            <td><input type="number" class="tB" value="${data.tB || ''}"></td>
                            <td><input type="number" class="hC" value="${data.hC || ''}"></td>
                            <td><input type="number" class="tD" value="${data.tD || ''}"></td>
                            <td><input type="number" class="tE" value="${data.tE || ''}"></td>
                            <td><input type="number" class="tF" value="${data.tF || ''}"></td>
                            <td><input type="text" class="phone" value="${data.phone || ''}"></td>
                            <td><input type="text" class="bN" value="${data.bN || ''}"></td>
                            <td><input type="text" class="bB" value="${data.bB || ''}"></td>
                            <td><button style="background:red; color:white; border:none; padding:5px; cursor:pointer;" onclick="this.parentElement.parentElement.remove()">刪</button></td>
                        `;
                        tbody.appendChild(tr);
                    });
                    alert(`成功匯入 ${tableRows.length} 筆資料`);
                }
            } catch (e) {
                console.error(e);
                alert('匯入失敗，請檢查檔案格式與表頭文字');
            }
        };
        reader.readAsArrayBuffer(file);
    };
    input.click();
}

// --- Excel 匯出功能 (確保欄位完整不缺失) ---
function exportData() {
    const getCheck = (id) => {
        const el = document.getElementById(id);
        return (el && el.checked) ? "■" : "";
    };

    // 1. 定義「英文 Key」對應「中文表頭」的字典
    const columnMapping = {
        // 上方勾選與文字區
        s_hour: "鐘點費", s_attend: "出席費", s_full: "專任助理",
        s_part: "兼任助理", s_work: "工讀費", s_rv_g: "審查費(一般)",
        s_host: "主持費", s_rv_gb: "交通費(定價)", s_rv_es: "交通費定價金額",
        s_50_other: "薪資其它勾選", s_50_text: "薪資其它註明",
        s_speech: "演講費", s_write: "稿費", s_rv_t: "審查費(教師)",
        s_paper: "論文指導費", s_9a_other: "執行業務其它勾選", s_9a_text: "執行業務其它註明",
        s_award: "競賽獎金", s_f_trf: "論文交通費", s_f_stu: "獎助生",
        s_f_gb: "免稅交通費勾選", s_f_gb_es: "免稅交通費金額",
        s_free_other: "免稅其它勾選", s_free_text: "免稅其它註明",
        s_royalty: "權利金",

        // 下方個人清單區
        name: "姓名", date: "日期", idCard: "身分證字號",
        acc: "銀行帳號", addr: "戶籍地址", org: "專職機構",
        job: "職銜", month: "月份", isPT: "是否兼任",
        amt: "總額(A)", tB: "所得稅(B)", hC: "補保(C)",
        tD: "勞保(D)", tE: "健保(E)", tF: "勞退(F)",
        phone: "電話", bN: "銀行名稱", bB: "分行名稱"
    };

    // 2. 收集上方狀態 (英文格式)
    const headerStatus = {
        s_hour: getCheck("s_hour"), s_attend: getCheck("s_attend"), s_full: getCheck("s_full"),
        s_part: getCheck("s_part"), s_work: getCheck("s_work"), s_rv_g: getCheck("s_rv_g"),
        s_host: getCheck("s_host"), s_rv_gb: getCheck("s_rv_gb"),
        s_rv_es: document.getElementById('s_rv_es')?.value || "",
        s_50_other: getCheck("s_50_other"), s_50_text: document.getElementById('s_50_text')?.value || "",
        s_speech: getCheck("s_speech"), s_write: getCheck("s_write"), s_rv_t: getCheck("s_rv_t"),
        s_paper: getCheck("s_paper"), s_9a_other: getCheck("s_9a_other"),
        s_9a_text: document.getElementById('s_9a_text')?.value || "",
        s_award: getCheck("s_award"), s_f_trf: getCheck("s_f_trf"), s_f_stu: getCheck("s_f_stu"),
        s_f_gb: getCheck("s_f_gb"), s_f_gb_es: document.getElementById('s_f_gb_es')?.value || "",
        s_free_other: getCheck("s_free_other"), s_free_text: document.getElementById('s_free_text')?.value || "",
        s_royalty: getCheck("s_royalty")
    };

    // 3. 收集所有行資料並轉換成中文 Key
    const rows = document.querySelectorAll('#dataList tr');
    let allDataChinese = [];

    // 輔助函式：將單一物件的 Key 從英文換成中文
    const convertToChinese = (obj) => {
        let newObj = {};
        for (let key in obj) {
            const chineseKey = columnMapping[key] || key; // 如果找不到對應就用原 Key
            newObj[chineseKey] = obj[key];
        }
        return newObj;
    };

    if (rows.length === 0) {
        allDataChinese.push(convertToChinese(headerStatus));
    } else {
        rows.forEach(row => {
            const rawRow = {
                ...headerStatus,
                name: row.querySelector('.name')?.value || "",
                date: row.querySelector('.row-date')?.value || "",
                idCard: row.querySelector('.idCard')?.value || "",
                acc: row.querySelector('.acc')?.value || "",
                addr: row.querySelector('.addr')?.value || "",
                org: row.querySelector('.org')?.value || "",
                job: row.querySelector('.job')?.value || "",
                month: row.querySelector('.month')?.value || "",
                isPT: row.querySelector('.isPT')?.value || "否",
                amt: row.querySelector('.amt')?.value || "",
                tB: row.querySelector('.tB')?.value || "",
                hC: row.querySelector('.hC')?.value || "",
                tD: row.querySelector('.tD')?.value || "",
                tE: row.querySelector('.tE')?.value || "",
                tF: row.querySelector('.tF')?.value || "",
                phone: row.querySelector('.phone')?.value || "",
                bN: row.querySelector('.bN')?.value || "",
                bB: row.querySelector('.bB')?.value || ""
            };
            allDataChinese.push(convertToChinese(rawRow));
        });
    }

    // 4. 執行 Excel 轉換與下載
    const worksheet = XLSX.utils.json_to_sheet(allDataChinese);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "領據資料");

    XLSX.writeFile(workbook, `領據資料匯出_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

// --- json ---
async function importData_() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    input.onchange = e => {
        const file = e.target.files[0];
        const reader = new FileReader();
        reader.onload = event => {
            try {
                const importedData = JSON.parse(event.target.result);
                if (Array.isArray(importedData)) {

                    // --- 關鍵修正區：這裡的 ID 必須跟你的 HTML <input id="..."> 一模一樣 ---
                    const checkFields = [
                        's_hour', 's_attend', 's_full', 's_part', 's_work', 's_rv_g',
                        's_host', 's_rv_gb', 's_50_other', 's_speech', 's_write',
                        's_rv_t', 's_paper', 's_9a_other', 's_award', 's_f_trf',
                        's_f_stu', 's_f_gb', 's_free_other', 's_royalty'
                    ]; // 注意：這裡已將 s_fr_ot 改回 s_free_other

                    const textFields = [
                        's_rv_es', 's_50_text', 's_9a_text', 's_f_gb_es', 's_free_text'
                    ]; // 注意：這裡已將 s_fr_txt 改回 s_free_text
                    // -------------------------------------------------------

                    // 處理 Checkbox
                    checkFields.forEach(id => {
                        const el = document.getElementById(id);
                        if (el) {
                            // 修正：也要檢查 header 以外的資料，確保勾選狀態能正確還原
                            el.checked = importedData.some(d => d[id] === "■");
                        }
                    });

                    // 處理文字輸入框
                    textFields.forEach(id => {
                        const el = document.getElementById(id);
                        if (el) {
                            const found = importedData.find(d => d[id]);
                            el.value = found ? found[id] : "";
                        }
                    });

                    const tbody = document.getElementById('dataList');
                    tbody.innerHTML = "";

                    // 過濾掉不含姓名的 Header 資料，避免產生空白列
                    const tableRows = importedData.filter(d => d.name || d.idCard);

                    tableRows.forEach(data => {
                        const tr = document.createElement('tr');
                        tr.innerHTML = `
                            <td><input type="text" class="name" value="${data.name || ''}"></td>
                            <td><input type="date" class="row-date" value="${data.date || ''}"></td>
                            <td><input type="text" class="idCard" value="${data.idCard || ''}" maxlength="10"></td>
                            <td><input type="text" class="acc" value="${data.acc || ''}" maxlength="14"></td>
                            <td><input type="text" class="addr" value="${data.addr || ''}"></td>
                            <td><input type="text" class="org" value="${data.org || ''}"></td>
                            <td><input type="text" class="job" value="${data.job || ''}"></td>
                            <td><input type="text" class="month" value="${data.month || ''}"></td>
                            <td><select class="isPT"><option value="否" ${data.isPT === "否" ? "selected" : ""}>否</option><option value="是" ${data.isPT === "是" ? "selected" : ""}>是</option></select></td>
                            <td><input type="number" class="amt" value="${data.amt || ''}"></td>
                            <td><input type="number" class="tB" value="${data.tB || ''}"></td>
                            <td><input type="number" class="hC" value="${data.hC || ''}"></td>
                            <td><input type="number" class="tD" value="${data.tD || ''}"></td>
                            <td><input type="number" class="tE" value="${data.tE || ''}"></td>
                            <td><input type="number" class="tF" value="${data.tF || ''}"></td>
                            <td><input type="text" class="phone" value="${data.phone || ''}"></td>
                            <td><input type="text" class="bN" value="${data.bN || ''}"></td>
                            <td><input type="text" class="bB" value="${data.bB || ''}"></td>
                            <td><button style="background:red; color:white; border:none; padding:5px; cursor:pointer;" onclick="this.parentElement.parentElement.remove()">刪</button></td>
                        `;
                        tbody.appendChild(tr);
                    });
                } else {
                    alert('匯入的資料格式不正確');
                }
            } catch (e) {
                console.error(e);
                alert('匯入資料時發生錯誤');
            }
        };
        reader.readAsText(file);
    };
    input.click();
}

// 資料匯出功能 
function exportData_() {
    // 輔助函式：判斷勾選並回傳特定符號 (符合你 import 的邏輯)
    const getCheck = (id) => {
        const el = document.getElementById(id);
        return (el && el.checked) ? "■" : "";
    };

    // 1. 處理上方勾選與文字 (使用物件收集)
    const header = {
        s_hour: getCheck("s_hour"),
        s_attend: getCheck("s_attend"),
        s_full: getCheck("s_full"),
        s_part: getCheck("s_part"),
        s_work: getCheck("s_work"),
        s_rv_g: getCheck("s_rv_g"),
        s_host: getCheck("s_host"),
        s_rv_gb: getCheck("s_rv_gb"),
        s_rv_es: document.getElementById('s_rv_es')?.value || "",
        s_50_other: getCheck("s_50_other"),
        s_50_text: document.getElementById('s_50_text')?.value || "",
        s_speech: getCheck("s_speech"),
        s_write: getCheck("s_write"),
        s_rv_t: getCheck("s_rv_t"),
        s_paper: getCheck("s_paper"),
        s_9a_other: getCheck("s_9a_other"),
        s_9a_text: document.getElementById('s_9a_text')?.value || "",
        s_award: getCheck("s_award"),
        s_f_trf: getCheck("s_f_trf"),
        s_f_stu: getCheck("s_f_stu"),
        s_f_gb: getCheck("s_f_gb"),
        s_f_gb_es: document.getElementById('s_f_gb_es')?.value || "",
        s_free_other: getCheck("s_free_other"), // 注意 ID 映射
        s_free_text: document.getElementById('s_free_text')?.value || "",
        s_royalty: getCheck("s_royalty")
    };

    // 2. 處理下方清單
    // 定義欄位清單，自動遍歷，避免手寫每一個屬性
    const fields = [
        'name', 'idCard', 'acc', 'addr', 'org', 'job',
        'month', 'isPT', 'amt', 'tB', 'hC', 'tD', 'tE', 'tF',
        'phone', 'bN', 'bB'
    ];

    const rows = document.querySelectorAll('#dataList tr');
    const items = Array.from(rows).map(row => {
        const rowData = {};
        // 處理日期 (因為你的 class 是 row-date 但 key 是 date)
        rowData.date = row.querySelector('.row-date')?.value || "";

        // 處理其他欄位
        fields.forEach(f => {
            rowData[f] = row.querySelector(`.${f}`)?.value || "";
        });
        return rowData;
    });

    // 3. 關鍵修正：將 Header 與 Items 合併到同一個資料結構中
    // 這裡建議將 header 放在陣列的第一筆，或是包成一個大物件
    // 依照你之前的 import 邏輯，放在同一個陣列的第一筆最方便
    const finalData = [header, ...items];

    // 4. 匯出檔案
    const blob = new Blob([JSON.stringify(finalData, null, 2)], { type: 'application/json' });
    const fileName = `data_export_${new Date().toISOString().slice(0, 10)}.json`;
    saveAs(blob, fileName);
}

// 重載入模板(手動)
async function reloadTemplate() {
    tplArrayBuffer = null;
    document.getElementById('tplStatus').innerText = "📡 正在重新載入樣板...";
    const timeoutId = setTimeout(async () => {
        try {
            const resp = await fetch(fetchtemp_name+'.docx');
            if (!resp.ok) throw new Error();
            tplArrayBuffer = await resp.arrayBuffer();
            document.getElementById('tplStatus').innerText = fetchtemp_name+"✅ 伺服器樣板載入成功";
            document.getElementById('tplStatus').className = "status-tag status-ok";
        } catch (e) {
            document.getElementById('tplStatus').innerText = "❌ 找不到 template.docx (請放置於根目錄)";
        }
        clearTimeout(timeoutId);
    }, 500); // 模擬載入時間
}

// 方案 C: 自動抓取 template.docx
async function init() {
    try {
        const resp = await fetch(fetchtemp_name+'.docx');
        if (!resp.ok) throw new Error();
        tplArrayBuffer = await resp.arrayBuffer();
        document.getElementById('tplStatus').innerText = fetchtemp_name+"✅ 伺服器樣板載入成功";
        document.getElementById('tplStatus').className = "status-tag status-ok";
    } catch (e) {
        document.getElementById('tplStatus').innerText = "❌ 找不到 template.docx (請放置於根目錄)";
    }
    // addRow(); // 預設一行
}

function addRow() {
    const today = new Date().toISOString().split('T')[0];
    const tbody = document.getElementById('dataList');
    const tr = document.createElement('tr');
    tr.innerHTML = `
            <td><input type="text" class="name" value=""></td>
            <td><input type="date" class="row-date" value="${today}"></td>
            <td><input type="text" class="idCard" value="" maxlength="10"></td>
            <td><input type="text" class="acc" value="" maxlength="14"></td>
            <td><input type="text" class="addr" value=""></td>
            <td><input type="text" class="org" value=""></td>
            <td><input type="text" class="job" value=""></td>
            <td><input type="text" class="month" value=""></td>
            <td><select class="isPT"><option value="否">否</option><option value="是">是</option></select></td>
            <td><input type="number" class="amt" value=""></td>
            <td><input type="number" class="tB" value=""></td>
            <td><input type="number" class="hC" value=""></td>
            <td><input type="number" class="tD" value=""></td>
            <td><input type="number" class="tE" value=""></td>
            <td><input type="number" class="tF" value=""></td>
            <td><input type="text" class="phone" value=""></td>
            <td><input type="text" class="bN" value=""></td>
            <td><input type="text" class="bB" value=""></td>
            <td><button style="background:red; color:white; border:none; padding:5px; cursor:pointer;" onclick="this.parentElement.parentElement.remove()">刪</button></td>
        `;
    tbody.appendChild(tr);
}

const isChecked = (id) => document.getElementById(id).checked ? "■" : "□";
const anyChecked = (cls) => Array.from(document.getElementsByClassName(cls)).some(el => el.checked) ? "■" : "□";
const splitToTags = (str, prefix, len) => {
    let obj = {};
    for (let i = 0; i < len; i++) obj[`${prefix}${i}`] = (str && str[i]) ? str[i] : " ";
    return obj;
};

function toChineseBig(n) {
    if (n === 0) return "零元整";
    const fraction = ['角', '分'];
    const digit = ['零', '壹', '貳', '參', '肆', '伍', '陸', '柒', '捌', '玖'];
    const unit = [
        ['元', '萬', '億'],
        ['', '拾', '佰', '仟']
    ];
    let s = '';
    for (let i = 0; i < unit[0].length && n > 0; i++) {
        let p = '';
        for (let j = 0; j < unit[1].length && n > 0; j++) {
            p = digit[n % 10] + unit[1][j] + p;
            n = Math.floor(n / 10);
        }
        s = p.replace(/(零.)*零$/, '').replace(/^$/, '零') + unit[0][i] + s;
    }
    return s.replace(/(零.)*零元/, '元')
        .replace(/(零.)+/g, '零')
        .replace(/^整$/, '零元整') + '整';
}

function getDateTags(dateStr) {
    if (!dateStr) return { d0: "  ", d1: "  ", d2: "  " };

    const dateObj = new Date(dateStr);
    const year = dateObj.getFullYear() - 1911; // 轉民國年
    const month = dateObj.getMonth() + 1;
    const day = dateObj.getDate();

    return {
        d0: year.toString(),
        d1: month.toString().padStart(2, '0'),
        d2: day.toString().padStart(2, '0')
    };
}

async function generateBatch() {
    // 1. 檢查樣板是否真的有抓到
    if (!tplArrayBuffer) {
        return alert("❌ 樣板尚未讀取成功！\n1. 請確認伺服器上有 template.docx\n2. 請確認你是透過網址(http)開啟，而非直接雙擊檔案(file://)");
    }

    const zip = new JSZip();
    const rows = document.querySelectorAll('#dataList tr');
    if (rows.length === 0) return alert("請至少新增一筆領款人資料");

    try {
        for (let row of rows) {
            // 讀取 PizZip
            let pizzip;
            try {
                pizzip = new PizZip(tplArrayBuffer);
            } catch (e) {
                throw new Error("樣板格式錯誤，請確保它是標準的 .docx 檔案");
            }

            const doc = new window.docxtemplater(pizzip, {
                paragraphLoop: true,
                linebreaks: true
            });

            // 抓取該列資料
            const name = row.querySelector('.name').value;
            const id = row.querySelector('.idCard').value.toUpperCase();
            const amt = parseInt(row.querySelector('.amt').value) || 0;
            const B = parseInt(row.querySelector('.tB').value) || 0;
            const C = parseInt(row.querySelector('.hC').value) || 0;
            const D = parseInt(row.querySelector('.tD').value) || 0;
            const E = parseInt(row.querySelector('.tE').value) || 0;
            const F = parseInt(row.querySelector('.tF').value) || 0;
            const netG = amt - (B + C + D + E + F);

            const data = {
                name: name,
                address: row.querySelector('.addr').value,
                org: row.querySelector('.org').value,
                job: row.querySelector('.job').value,
                month: row.querySelector('.month').value,
                phone: row.querySelector('.phone').value,
                bank_name: row.querySelector('.bN').value,
                branch_name: row.querySelector('.bB').value,
                is_y: row.querySelector('.isPT').value === "是" ? "■" : "□",
                is_n: row.querySelector('.isPT').value === "否" ? "■" : "□",
                amount_a: amt, tax_b: B || "0", health_c: C || "0",
                tax_d: D || "0", tax_e: E || "0", tax_f: F || "0",
                net_g: netG,
                // 所得連動
                c50: anyChecked("cat50"), c9a: anyChecked("cat9a"),
                c91: anyChecked("cat91"), c_free: anyChecked("catFree"), c53: anyChecked("cat53"),
                // 細項
                s_hour: isChecked("s_hour"), s_attend: isChecked("s_attend"),
                s_full: isChecked("s_full"), s_part: isChecked("s_part"),
                s_work: isChecked("s_work"), s_rv_g: isChecked("s_rv_g"),
                s_host: isChecked("s_host"),
                s_rv_gb: isChecked("s_rv_gb"), s_rv_es: document.getElementById('s_rv_es').value,
                s_50_ot: isChecked("s_50_other"), s_50_txt: document.getElementById('s_50_text').value,
                s_speech: isChecked("s_speech"), s_write: isChecked("s_write"),
                s_rv_t: isChecked("s_rv_t"), s_paper: isChecked("s_paper"),
                s_9a_ot: isChecked("s_9a_other"), s_9a_txt: document.getElementById('s_9a_text').value,
                s_award: isChecked("s_award"), s_f_trf: isChecked("s_f_trf"),
                s_f_stu: isChecked("s_f_stu"), s_f_gb: isChecked("s_f_gb"),
                s_f_gb_es: document.getElementById('s_f_gb_es').value,
                s_fr_ot: isChecked("s_free_other"), s_fr_txt: document.getElementById('s_free_text').value,
                s_royalty: isChecked("s_royalty"),
                // 拆解
                full_chinese_amount: netG > 0 ? toChineseBig(netG) : "零元整",
                ...splitToTags(id, "i", 10),
                ...getDateTags(row.querySelector('.row-date').value, "d", 3),
                ...splitToTags(row.querySelector('.acc').value, "a", 14)
            };

            doc.render(data);
            const out = doc.getZip().generate({ type: "blob" });
            zip.file(`領據_${name}.docx`, out);
        }

        const zipBlob = await zip.generateAsync({ type: "blob" });
        saveAs(zipBlob, "僑光領據_批次產出.zip");

    } catch (error) {
        console.error(error);
        alert("❌ 產生過程中發生錯誤：\n" + error.message + "\n請檢查樣板內的標籤是否寫錯（例如少了括號）。");
    }
}
// 在 <script> 標籤內加入以下邏輯
document.addEventListener('keydown', function (e) {
    if (e.key === 'Enter') {
        // 取得目前獲得焦點的元素
        const activeEl = document.activeElement;

        // 如果目前在 input 或 select 上
        if (activeEl.tagName === 'INPUT' || activeEl.tagName === 'SELECT') {
            // 避免觸發表單預設提交行為
            e.preventDefault();

            // 取得所有可輸入的元素清單
            const formElements = Array.from(document.querySelectorAll('input:not([type="hidden"]), select, button.btn-add, .main-btn'));
            const index = formElements.indexOf(activeEl);

            if (index > -1 && index < formElements.length - 1) {
                // 跳轉到下一個元素
                formElements[index + 1].focus();

                // 如果下一個是文字輸入框，順便自動全選文字，方便直接覆蓋輸入
                if (formElements[index + 1].tagName === 'INPUT') {
                    formElements[index + 1].select();
                }
            }
        }
    }
});

init();

document.addEventListener("DOMContentLoaded", () => {
    const body = document.body;
    const themetoggle = document.querySelector("#theme-btn");
    const lightBtn = document.querySelector('[title="淺色"]');
    const darkBtn = document.querySelector('[title="深色"]');
    const defaultBtn = document.querySelector('[title="預設"]');

    // 1. 初始化檢查
    const userSet = getCookie("theme") || "auto";
    applyTheme(userSet);

    // 2. 綁定按鈕事件
    if (lightBtn && darkBtn && defaultBtn && themetoggle) {
        lightBtn.addEventListener("click", () => {
            setCookie("theme", "light");
            applyTheme("light");
        });

        darkBtn.addEventListener("click", () => {
            setCookie("theme", "dark");
            applyTheme("dark");
        });

        defaultBtn.addEventListener("click", () => {
            setCookie("theme", "auto");
            applyTheme("auto");
        });
    }

    // 3. 監聽系統主題變更 (當使用者設定為 "auto" 時)
    window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", e => {
        if (getCookie("theme") === "auto") {
            applyTheme("auto");
        }
    });

    // 核心切換邏輯
    function applyTheme(theme) {
        let activeTheme = theme;

        // 如果是自動，判斷系統偏好
        if (theme === "auto") {
            activeTheme = window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
        }

        // 執行切換屬性 (對接 CSS)
        body.setAttribute("data-bs-theme", activeTheme);

        // 更新 UI 圖示與文字
        updateButtonUI(theme, activeTheme);
    }

    function updateButtonUI(savedSetting, currentActive) {
        if (!themetoggle) return;

        if (savedSetting === "auto") {
            themetoggle.innerHTML = '<i class="bi bi-circle-half"></i> 預設';
        } else if (currentActive === "dark") {
            themetoggle.innerHTML = '<i class="bi bi-brightness-high-fill"></i> 深色';
        } else {
            themetoggle.innerHTML = '<i class="bi bi-brightness-high"></i> 淺色';
        }
    }

    // Cookie 工具函數
    function setCookie(name, value, days = 365) {
        const expires = new Date(Date.now() + days * 864e5).toUTCString();
        document.cookie = `${name}=${value}; expires=${expires}; path=/; SameSite=Lax`;
    }

    function getCookie(name) {
        const value = `; ${document.cookie}`;
        const parts = value.split(`; ${name}=`);
        if (parts.length === 2) return parts.pop().split(';').shift();
        return null;
    }

    // 在你的 DOMContentLoaded 監聽器裡面加入：
    const dropdown = document.querySelector("#theme-dropdown");
    const trigger = document.querySelector("#theme-btn");

    // 點擊按鈕切換選單顯示/隱藏
    trigger.addEventListener("click", (e) => {
        e.stopPropagation(); // 防止冒泡
        dropdown.classList.toggle("show");
    });

    // 點擊網頁其他地方時關閉選單
    document.addEventListener("click", () => {
        dropdown.classList.remove("show");
    });

    // 當選單內的按鈕被點擊後，也要關閉選單
    document.querySelectorAll(".dropdown-content button").forEach(btn => {
        btn.addEventListener("click", () => {
            dropdown.classList.remove("show");
        });
    });
});