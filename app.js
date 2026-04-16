const GAS_URL = "https://script.google.com/macros/s/AKfycbyoSgJMFnMlzFf0WoX1SSXB9kxiqvMjtiHW_T8MoVRCO9fZJFAYsPp6o1mP2rDI4eJP/exec";
const CONFIG = {
    hiddenColumns: ["तारीख", "मोबाईल क्र.", "उपकेंद्र", "महिना", "वर्ष", "मूळ डेटा (JSON)", "कर्मचाऱ्याचे नाव"]
};

let masterData = { forms: [], villages: [], filledStats: [] };
let user = null;
let currentReports = []; 
let isSaving = false; 

function generateInputHTML(f, id, label, areaId, val="") {
    let html = "";
    if (f.type === 'dropdown') {
        html += `<select id="${id}" data-label="${label}" onchange="calculateAutoSums('${areaId}')"><option value="">-- निवडा --</option>`;
        if(f.options) { f.options.split(',').forEach(opt => { 
            let o = opt.trim(); 
            let sel = (o === val) ? "selected" : "";
            if(o) html += `<option value="${o}" ${sel}>${o}</option>`; 
        }); }
        html += `</select>`;
    } else if(f.type === 'sum') {
        html += `<input type="number" id="${id}" data-label="${label}" data-sum-targets="${f.options}" value="${val}" readonly style="background:#e9ecef; font-weight:bold; color:var(--primary); cursor:not-allowed;" placeholder="Auto Sum">`;
    } else if(f.type === 'number') {
        html += `<input type="number" id="${id}" data-label="${label}" value="${val}" oninput="calculateAutoSums('${areaId}')">`;
    } else if(f.type === 'date') {
        html += `<input type="date" id="${id}" data-label="${label}" value="${val}" onchange="calculateAutoSums('${areaId}')">`;
    } else { 
        html += `<input type="text" id="${id}" data-label="${label}" value="${val}" oninput="calculateAutoSums('${areaId}')">`; 
    }
    return html;
}

async function unlockApp() {
    if(document.getElementById('sitePass').value === "1234") {
        document.getElementById('lockScreen').classList.add('hidden');
        document.getElementById('mainApp').classList.remove('hidden');
        document.getElementById('netStatus').innerText = "डेटा लोड होत आहे...";
        await fetchData();
        document.getElementById('netStatus').innerText = "Online";
        const savedUser = localStorage.getItem("phc_user_session");
        if(savedUser) { user = JSON.parse(savedUser); showAppAfterLogin(); }
    } else { alert("चुकीचा पासवर्ड!"); }
}

async function fetchData() {
    try {
        const r = await fetch(GAS_URL, { method: "POST", body: JSON.stringify({action:"getInitialData"}) });
        const textResponse = await r.text();
        if(textResponse.trim().startsWith("<")) throw new Error("Google Blocked Request");
        const d = JSON.parse(textResponse);
        if(d.success) { 
            masterData = d; 
            updateFormDropdowns(); 
            renderFormsListForEdit(); 
        }
    } catch(e) { console.error("Fetch failed", e); }
}

async function handleLogin() {
    const m = document.getElementById('mob').value.trim();
    const p = document.getElementById('pwd').value.trim();
    if(!m || !p) return;
    document.getElementById('netStatus').innerText = "तपासत आहे...";
    try {
        const r = await fetch(GAS_URL, { method: "POST", body: JSON.stringify({action:"login", mobileNo:m, password:p}) });
        const textResponse = await r.text();
        if(textResponse.trim().startsWith("<")) throw new Error("Google Blocked Request");
        const d = JSON.parse(textResponse);
        if(d.success) {
            user = d.user; user.mobile = m;
            localStorage.setItem("phc_user_session", JSON.stringify(user));
            showAppAfterLogin();
        } else { alert(d.message); }
        document.getElementById('netStatus').innerText = "Online";
    } catch(e) { alert("लॉगिन अयशस्वी. कृपया इंटरनेट कनेक्शन तपासा."); document.getElementById('netStatus').innerText = "Offline"; }
}

function showAppAfterLogin() {
    document.getElementById('loginBox').classList.add('hidden');
    document.getElementById('dashboardWrapper').classList.remove('hidden');
    document.getElementById('uName').innerText = user.name;
    document.getElementById('uSub').innerText = "उपकेंद्र: " + user.subcenter;
    
    if(user.role === "Admin") { 
        document.getElementById('tabAdmin').classList.remove('hidden'); 
        document.getElementById('adminRoleFilterDiv').style.display = "flex";
        renderFormsListForEdit(); 
    }
    updateFormDropdowns();
    updateVillageDropdown();
    updateEditVillageDropdown();
}

function logoutUser() {
    if(confirm("लॉग आऊट करायचे आहे का?")) { localStorage.removeItem("phc_user_session"); location.reload(); }
}

function switchTab(tab) {
    document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
    document.getElementById('entrySection').classList.add('hidden');
    document.getElementById('editSection').classList.add('hidden');
    document.getElementById('reportsSection').classList.add('hidden');
    document.getElementById('adminSection').classList.add('hidden');

    if(tab === 'entry') {
        document.getElementById('tabEntry').classList.add('active');
        document.getElementById('entrySection').classList.remove('hidden');
    } else if(tab === 'edit') {
        document.getElementById('tabEdit').classList.add('active');
        document.getElementById('editSection').classList.remove('hidden');
        updateEditVillageDropdown();
    } else if(tab === 'reports') {
        document.getElementById('tabReports').classList.add('active');
        document.getElementById('reportsSection').classList.remove('hidden');
    } else if(tab === 'admin') {
        document.getElementById('tabAdmin').classList.add('active');
        document.getElementById('adminSection').classList.remove('hidden');
    }
}

function updateFormDropdowns() {
    const selForm = document.getElementById('selForm');
    const editForm = document.getElementById('editFormSelect');
    const repForm = document.getElementById('reportFormSelect');
    
    selForm.innerHTML = '<option value="">-- निवडा --</option>';
    editForm.innerHTML = '<option value="">-- निवडा --</option>';
    repForm.innerHTML = '<option value="">-- निवडा --</option><option value="ALL" style="font-weight:bold; color:var(--primary);">सर्व अहवाल एकत्रित (All Forms)</option>';
    
    masterData.forms.forEach(f => {
        let allowedRoles = f.AllowedRoles ? f.AllowedRoles.split(',').map(r=>r.trim().toUpperCase()) : ["ALL"];
        let userRole = user ? String(user.role).trim().toUpperCase() : "";
        
        if (userRole === "ADMIN" || allowedRoles.includes("ALL") || allowedRoles.includes(userRole)) {
            let opt = `<option value="${f.FormID}">${f.FormName}</option>`;
            selForm.innerHTML += opt;
            editForm.innerHTML += opt;
            repForm.innerHTML += opt;
        }
    });
}

function updateVillageDropdown() {
    const vSel = document.getElementById('selVillage');
    const fId = document.getElementById('selForm').value;
    const month = document.getElementById('selMonth').value;
    const year = document.getElementById('selYear').value;
    vSel.innerHTML = '<option value="">-- गाव निवडा --</option>';
    if(!user || !fId) return;

    const selectedForm = masterData.forms.find(f => f.FormID === fId);
    const isStatsForm = selectedForm && selectedForm.FormType === 'Stats';
    let allowedRoles = selectedForm && selectedForm.AllowedRoles ? selectedForm.AllowedRoles.split(',').map(r=>r.trim().toUpperCase()) : ["ALL"];
    let isAll = allowedRoles.includes("ALL");
    
    const localHistory = JSON.parse(localStorage.getItem("submissionHistory") || "[]");
    const serverHistory = masterData.filledStats || [];

    masterData.villages.filter(v => {
        const belongsToSubCenter = v.SubCenterID.toLowerCase() === user.subcenter.toLowerCase() || v.SubCenterID.toLowerCase() === "all";
        if(!belongsToSubCenter) return false;
        
        if(isStatsForm) {
            const isLocalFilled = localHistory.some(h => 
                h.formID === fId && h.village === v.VillageName && 
                h.month === month && h.year == year && 
                h.mobile === String(user.mobile).trim()
            );
            const isServerFilled = serverHistory.some(h => {
                if(h.formID !== fId || h.village !== v.VillageName || h.month !== month || h.year != year) return false;
                if(h.isAllForm || isAll) return true;
                return h.mobile === String(user.mobile).trim();
            });
            if(isLocalFilled || isServerFilled) return false;
        }
        return true;
    }).forEach(v => { vSel.innerHTML += `<option value="${v.VillageName}">${v.VillageName}</option>`; });
}

function loadDynamicFields() {
    const fId = document.getElementById('selForm').value;
    const area = document.getElementById('dynamicFormArea');
    area.innerHTML = "";
    const f = masterData.forms.find(x => x.FormID === fId);
    if(!f) return;
    
    let html = "";
    JSON.parse(f.StructureJSON).forEach((field, i) => {
        let exactLabel = field.label;
        if (field.type === 'group') {
            html += `<div style="margin-bottom:15px; background:#fffaf0; padding:12px; border-radius:8px; border:1px solid #f5b041;">
                        <h4 style="margin-top:0; color:var(--primary); text-align:left; border-bottom:1px solid #ccc; padding-bottom:5px;">${field.label}</h4>`;
            field.subFields.forEach((sf, j) => {
                if(sf.type === 'group') {
                    html += `<div style="margin-bottom:10px; margin-left:10px; background:#e0f7fa; padding:10px; border-radius:5px; border-left:3px solid #00acc1;">
                             <h5 style="margin:0 0 5px 0; color:#00838f;">${sf.label}</h5>`;
                    sf.subFields.forEach((ssf, k) => {
                        let exactSubSubLabel = `${field.label} - ${sf.label} - ${ssf.label}`;
                        html += `<div style="margin-bottom:8px;"><label style="font-size:13px; color:#555;"><b>${ssf.label}:</b></label>`;
                        html += generateInputHTML(ssf, `inp_${i}_${j}_${k}`, exactSubSubLabel, 'dynamicFormArea', "");
                        html += `</div>`;
                    });
                    html += `</div>`;
                } else {
                    let exactSubLabel = `${field.label} - ${sf.label}`;
                    html += `<div style="margin-bottom:10px;"><label style="font-size:14px; color:#555;"><b>${sf.label}:</b></label>`;
                    html += generateInputHTML(sf, `inp_${i}_${j}`, exactSubLabel, 'dynamicFormArea', "");
                    html += `</div>`;
                }
            });
            html += `</div>`;
        } else {
            let exactLabel = field.label;
            html += `<div style="margin-bottom:15px; background:white; padding:10px; border-radius:8px; border:1px solid #ddd;"><label><b>${field.label}:</b></label>`;
            html += generateInputHTML(field, `inp_${i}`, exactLabel, 'dynamicFormArea', "");
            html += `</div>`;
        }
    });
    area.innerHTML = html;
}

async function saveDataLocal() {
    if(isSaving) return; 
    const saveBtn = document.getElementById('mainSaveBtn');
    if(saveBtn) saveBtn.disabled = true; 
    isSaving = true;

    calculateAutoSums('dynamicFormArea'); 
    
    const fId = document.getElementById('selForm').value;
    const vName = document.getElementById('selVillage').value;
    const month = document.getElementById('selMonth').value;
    const year = document.getElementById('selYear').value;

    if(!fId || !vName) { alert("कृपया फॉर्म आणि गाव निवडा!"); isSaving = false; if(saveBtn) saveBtn.disabled = false; return; }
    
    let formData = {};
    formData["महिना"] = month;
    formData["वर्ष"] = year;

    const f = masterData.forms.find(x => x.FormID === fId);
    JSON.parse(f.StructureJSON).forEach((field, i) => {
        if (field.type === 'group') {
            field.subFields.forEach((sf, j) => { 
                if(sf.type === 'group') {
                    sf.subFields.forEach((ssf, k) => {
                        formData[`${field.label} - ${sf.label} - ${ssf.label}`] = document.getElementById(`inp_${i}_${j}_${k}`).value;
                    });
                } else {
                    formData[`${field.label} - ${sf.label}`] = document.getElementById(`inp_${i}_${j}`).value;
                }
            });
        } else { 
            formData[field.label] = document.getElementById(`inp_${i}`).value; 
        }
    });
    
    const entry = { entryID: Date.now(), mobileNo: user.mobile, subCenter: user.subcenter, village: vName, formID: fId, formData: formData };
    let q = JSON.parse(localStorage.getItem("syncQueue") || "[]");
    q.push(entry);
    localStorage.setItem("syncQueue", JSON.stringify(q));

    let history = JSON.parse(localStorage.getItem("submissionHistory") || "[]");
    history.push({formID: fId, village: vName, month: month, year: year, mobile: String(user.mobile).trim()});
    localStorage.setItem("submissionHistory", JSON.stringify(history));

    await syncDataToServer("syncStatus", "mainSaveBtn");
    updateVillageDropdown(); 
    loadDynamicFields(); 
    isSaving = false; 
}

async function syncDataToServer(statusId, btnId) {
    let q = JSON.parse(localStorage.getItem("syncQueue") || "[]");
    if(q.length === 0) return;
    
    const statusText = document.getElementById(statusId);
    const saveBtn = document.getElementById(btnId);

    statusText.style.color = "orange";
    statusText.innerText = "☁️ डेटा गुगल शीटवर सेव्ह होत आहे... कृपया थांबा.";

    try {
        const r = await fetch(GAS_URL, { method: "POST", body: JSON.stringify({action:"syncData", payload: q}) });
        const textResponse = await r.text();
        if(textResponse.trim().startsWith("<")) throw new Error("Google Blocked Request");

        const d = JSON.parse(textResponse);
        if(d.success) { 
            localStorage.setItem("syncQueue", "[]"); 
            statusText.style.color = "green";
            statusText.innerText = "✅ माहिती यशस्वीरित्या गुगल शीटवर सेव्ह झाली!"; 
            setTimeout(() => { statusText.innerText = ""; }, 4000);
            await fetchData(); 
        } else { throw new Error("Server error"); }
    } catch(e) { 
        statusText.style.color = "red";
        statusText.innerText = "⚠️ इंटरनेट एरर! माहिती मोबाईलमध्ये सुरक्षित आहे, नेटवर्क आल्यावर पुन्हा जतन करा दाबा.";
        setTimeout(() => { statusText.innerText = ""; }, 5000);
    } finally {
        if(saveBtn) saveBtn.disabled = false; 
    }
}

function updateEditVillageDropdown() {
    const vSel = document.getElementById('editVillageSelect');
    const fId = document.getElementById('editFormSelect').value;
    const month = document.getElementById('editMonth').value;
    const year = document.getElementById('editYear').value;
    
    vSel.innerHTML = '<option value="">-- भरलेले गाव / रेकॉर्ड निवडा --</option>';
    if(!user || !fId) return;

    const f = masterData.forms.find(x => x.FormID === fId);
    let allowedRoles = f && f.AllowedRoles ? f.AllowedRoles.split(',').map(r=>r.trim().toUpperCase()) : ["ALL"];
    let isAll = allowedRoles.includes("ALL");

    const serverHistory = masterData.filledStats || [];
    let addedVillages = [];

    serverHistory.forEach(h => {
        if(h.formID === fId && h.month === month && String(h.year) === String(year)) {
            let canEdit = false;
            if(user.role === "Admin") canEdit = true;
            else if(isAll && h.subcenter.toLowerCase() === user.subcenter.toLowerCase()) canEdit = true;
            else if(!isAll && h.mobile === String(user.mobile).trim()) canEdit = true;

            if(canEdit && !addedVillages.includes(h.village)) {
                vSel.innerHTML += `<option value="${h.village}">${h.village}</option>`;
                addedVillages.push(h.village);
            }
        }
    });
    
    document.getElementById('editDynamicFormArea').classList.add('hidden');
    document.getElementById('editSaveBtn').classList.add('hidden');
}

async function fetchRecordForEdit() {
    const formID = document.getElementById('editFormSelect').value;
    const month = document.getElementById('editMonth').value;
    const year = document.getElementById('editYear').value;
    const village = document.getElementById('editVillageSelect').value;

    if(!formID || !village) { alert("कृपया फॉर्म आणि गाव निवडा."); return; }

    document.getElementById('editLoader').style.display = "block";
    document.getElementById('editDynamicFormArea').classList.add('hidden');
    document.getElementById('editSaveBtn').classList.add('hidden');

    try {
        const payload = { formID: formID, village: village, month: month, year: year, subCenter: user.subcenter, mobileNo: user.mobile, role: user.role };
        const r = await fetch(GAS_URL, { method: "POST", body: JSON.stringify({action:"getRecordForEdit", payload}) });
        const textResponse = await r.text();
        
        if(textResponse.trim().startsWith("<")) throw new Error("Google Server Blocked the Request.");
        
        const d = JSON.parse(textResponse);
        document.getElementById('editLoader').style.display = "none";
        
        if(d.success) {
            renderEditForm(formID, d.formData);
        } else {
            alert("रेकॉर्ड सापडला नाही! कदाचित तो डिलीट झाला असेल.");
        }
    } catch(e) {
        document.getElementById('editLoader').style.display = "none";
        alert("डेटा लोड करण्यात एरर आला: " + e.message);
    }
}

function renderEditForm(fId, formData) {
    const area = document.getElementById('editDynamicFormArea');
    area.innerHTML = "";
    const f = masterData.forms.find(x => x.FormID === fId);
    if(!f) return;
    
    let html = "";
    JSON.parse(f.StructureJSON).forEach((field, i) => {
        let exactLabel = field.label;
        if (field.type === 'group') {
            html += `<div style="margin-bottom:15px; background:#e3f2fd; padding:12px; border-radius:8px; border:1px solid #f39c12;">
                        <h4 style="margin-top:0; color:#d35400; text-align:left; border-bottom:1px solid #ccc; padding-bottom:5px;">${field.label}</h4>`;
            field.subFields.forEach((sf, j) => {
                if(sf.type === 'group') {
                    html += `<div style="margin-bottom:10px; margin-left:10px; background:#e0f7fa; padding:10px; border-radius:5px; border-left:3px solid #00acc1;">
                             <h5 style="margin:0 0 5px 0; color:#00838f;">${sf.label}</h5>`;
                    sf.subFields.forEach((ssf, k) => {
                        let exactSubSubLabel = `${field.label} - ${sf.label} - ${ssf.label}`;
                        let val = formData[exactSubSubLabel] || "";
                        html += `<div style="margin-bottom:8px;"><label style="font-size:13px; color:#555;"><b>${ssf.label}:</b></label>`;
                        html += generateInputHTML(ssf, `edit_inp_${i}_${j}_${k}`, exactSubSubLabel, 'editDynamicFormArea', val);
                        html += `</div>`;
                    });
                    html += `</div>`;
                } else {
                    let exactSubLabel = `${field.label} - ${sf.label}`;
                    let val = formData[exactSubLabel] || "";
                    html += `<div style="margin-bottom:10px;"><label style="font-size:14px; color:#555;"><b>${sf.label}:</b></label>`;
                    html += generateInputHTML(sf, `edit_inp_${i}_${j}`, exactSubLabel, 'editDynamicFormArea', val);
                    html += `</div>`;
                }
            });
            html += `</div>`;
        } else {
            let val = formData[exactLabel] || "";
            html += `<div style="margin-bottom:15px; background:white; padding:10px; border-radius:8px; border:1px solid #ddd;"><label><b>${field.label}:</b></label>`;
            html += generateInputHTML(field, `edit_inp_${i}`, exactLabel, 'editDynamicFormArea', val);
            html += `</div>`;
        }
    });
    area.innerHTML = html;
    
    area.classList.remove('hidden');
    document.getElementById('editSaveBtn').classList.remove('hidden');
    calculateAutoSums('editDynamicFormArea'); 
}

async function saveEditedData() {
    if(isSaving) return;
    const saveBtn = document.getElementById('editSaveBtn');
    saveBtn.disabled = true;
    isSaving = true;

    calculateAutoSums('editDynamicFormArea');
    
    const fId = document.getElementById('editFormSelect').value;
    const month = document.getElementById('editMonth').value;
    const year = document.getElementById('editYear').value;
    const vName = document.getElementById('editVillageSelect').value;
    const statusText = document.getElementById('editSyncStatus');

    let updatedFormData = {};
    updatedFormData["महिना"] = month;
    updatedFormData["वर्ष"] = year;

    const f = masterData.forms.find(x => x.FormID === fId);
    JSON.parse(f.StructureJSON).forEach((field, i) => {
        if (field.type === 'group') {
            field.subFields.forEach((sf, j) => { 
                if(sf.type === 'group') {
                    sf.subFields.forEach((ssf, k) => {
                        updatedFormData[`${field.label} - ${sf.label} - ${ssf.label}`] = document.getElementById(`edit_inp_${i}_${j}_${k}`).value;
                    });
                } else {
                    updatedFormData[`${field.label} - ${sf.label}`] = document.getElementById(`edit_inp_${i}_${j}`).value;
                }
            });
        } else { 
            updatedFormData[field.label] = document.getElementById(`edit_inp_${i}`).value; 
        }
    });

    const payload = { formID: fId, village: vName, month: month, year: year, mobileNo: user.mobile, subCenter: user.subcenter, role: user.role, formData: updatedFormData };

    statusText.style.color = "orange";
    statusText.innerText = "☁️ नवीन बदल गुगल शीटवर सेव्ह होत आहेत...";

    try {
        const r = await fetch(GAS_URL, { method: "POST", body: JSON.stringify({action:"updateRecord", payload: payload}) });
        const textResponse = await r.text();
        
        if(textResponse.trim().startsWith("<")) throw new Error("Google Blocked Request");

        const d = JSON.parse(textResponse);
        if(d.success) { 
            statusText.style.color = "green";
            statusText.innerText = "✅ बदल यशस्वीरित्या अपडेट झाले!"; 
            setTimeout(() => { statusText.innerText = ""; }, 4000);
            
            document.getElementById('editDynamicFormArea').classList.add('hidden');
            saveBtn.classList.add('hidden');
            updateEditVillageDropdown();
            await fetchData(); 
        } else { throw new Error(d.message); }
    } catch(e) { 
        statusText.style.color = "red";
        statusText.innerText = "⚠️ बदल सेव्ह करताना एरर आला. कृपया पुन्हा प्रयत्न करा.";
    } finally {
        saveBtn.disabled = false;
        isSaving = false;
    }
}

function calculateAutoSums(containerId) {
    let container = document.getElementById(containerId);
    if(!container) return;
    
    let sumInputs = container.querySelectorAll('input[data-sum-targets]');
    sumInputs.forEach(sumInput => {
        let targets = sumInput.getAttribute('data-sum-targets').split(',');
        let total = 0;
        targets.forEach(targetLabel => {
            let tLabel = targetLabel.trim();
            let targetElement = container.querySelector(`[data-label="${tLabel}"]`);
            
            if(!targetElement) {
                let allInputs = container.querySelectorAll('[data-label]');
                let cleanTarget = tLabel.replace(/\s+/g, '');
                for(let inp of allInputs) {
                    if(inp.getAttribute('data-label').replace(/\s+/g, '') === cleanTarget) {
                        targetElement = inp;
                        break;
                    }
                }
            }

            if(targetElement && targetElement.value) {
                let val = parseFloat(targetElement.value);
                if(!isNaN(val)) total += val;
            }
        });
        sumInput.value = total > 0 ? total : ""; 
    });
}

function clearLocalHistory() {
    if(confirm("जर तुम्ही गुगल शीटमधील जुना डेटा डिलीट केला असेल, तरच हे बटण दाबा. यामुळे लपलेली सर्व गावे पुन्हा दिसू लागतील. तुम्हाला गावे रीसेट करायची आहेत का?")) {
        localStorage.removeItem("submissionHistory"); 
        updateVillageDropdown(); 
        const statusText = document.getElementById('syncStatus');
        statusText.style.color = "#ffc107";
        statusText.innerText = "✅ गावांची यादी यशस्वीरित्या रीसेट झाली आहे!";
        setTimeout(() => { statusText.innerText = ""; statusText.style.color = "green"; }, 3000);
    }
}

// --- Admin Form Builder Logic ---
function toggleRoles(checkbox) {
    document.getElementById('specificRoles').style.display = checkbox.checked ? 'none' : 'block';
    if(checkbox.checked) {
        document.querySelectorAll('.form-role').forEach(cb => cb.checked = false);
    }
}

function getSelectedRoles() {
    if(document.getElementById('roleAll').checked) return "ALL";
    let roles = [];
    document.querySelectorAll('.form-role:checked').forEach(cb => roles.push(cb.value));
    return roles.length > 0 ? roles.join(',') : "ALL";
}

function openNewFormBuilder() {
    document.getElementById('existingFormsArea').classList.add('hidden');
    document.getElementById('formBuilder').classList.remove('hidden');
    document.getElementById('builderTitle').innerText = "नवीन फॉर्म तयार करा";
    document.getElementById('editFormID').value = "";
    document.getElementById('newFormName').value = "";
    
    document.getElementById('roleAll').checked = true;
    toggleRoles(document.getElementById('roleAll'));

    document.getElementById('fieldsList').innerHTML = "";
    document.getElementById('mainActionBtn').innerText = "फॉर्म सेव्ह करा";
    document.getElementById('mainActionBtn').onclick = saveFullForm;
}

function renderFormsListForEdit() {
    const listDiv = document.getElementById('formsEditList');
    listDiv.innerHTML = "";
    masterData.forms.forEach(f => {
        listDiv.innerHTML += `<div class="edit-row" style="background:white; padding:10px; margin-bottom:5px; border-radius:5px; border:1px solid #ddd; display:flex; justify-content:space-between; align-items:center;">
                                <span><b>${f.FormName}</b></span>
                                <button class="btn-edit-tab" style="padding:6px 15px; width:auto; border-radius:4px;" onclick="startEditing('${f.FormID}')">Edit</button>
                              </div>`;
    });
}

function startEditing(fId) {
    const f = masterData.forms.find(x => x.FormID === fId);
    if(!f) return;
    document.getElementById('existingFormsArea').classList.add('hidden');
    document.getElementById('formBuilder').classList.remove('hidden');
    document.getElementById('builderTitle').innerText = "फॉर्म एडिट: " + f.FormName;
    document.getElementById('editFormID').value = f.FormID;
    document.getElementById('newFormName').value = f.FormName;
    document.getElementById('newFormType').value = f.FormType;
    
    let roles = f.AllowedRoles ? f.AllowedRoles.split(',').map(r=>r.trim().toUpperCase()) : ["ALL"];
    if(roles.includes("ALL")) {
        document.getElementById('roleAll').checked = true;
        toggleRoles(document.getElementById('roleAll'));
    } else {
        document.getElementById('roleAll').checked = false;
        toggleRoles(document.getElementById('roleAll'));
        document.querySelectorAll('.form-role').forEach(cb => {
            cb.checked = roles.includes(cb.value.toUpperCase());
        });
    }

    document.getElementById('fieldsList').innerHTML = "";
    document.getElementById('mainActionBtn').innerText = "बदल अपडेट करा (Update)";
    document.getElementById('mainActionBtn').onclick = updateExistingForm;
    JSON.parse(f.StructureJSON).forEach(field => addField(field));
}

function toggleFieldOptions(id, val) {
    document.getElementById('opts_' + id).style.display = (val === 'dropdown' || val === 'sum') ? 'block' : 'none';
    document.getElementById('group_' + id).style.display = (val === 'group') ? 'block' : 'none';
    let labelEl = document.querySelector(`#opts_${id} label`);
    let inputEl = document.querySelector(`#opts_${id} .foptions`);
    if(labelEl && inputEl) {
        if(val === 'sum') {
            labelEl.innerText = "कोणत्या प्रश्नांची बेरीज करायची? (स्वल्पविरामाने अचूक नावे लिहा):";
            inputEl.placeholder = "उदा. 0-5 वर्षे मुले, 0-5 वर्षे मुली";
        } else if(val === 'dropdown') {
            labelEl.innerText = "येथे पर्याय लिहा (स्वल्पविरामाने , वेगळे करा):";
            inputEl.placeholder = "उदा. होय, नाही";
        }
    }
}

function toggleSubFieldOptions(sfId, val) {
    let container = document.getElementById(`sfopts_container_${sfId}`);
    let groupContainer = document.getElementById(`ssf_group_${sfId}`);
    let labelEl = document.getElementById(`sflabel_${sfId}`);
    let inputEl = document.getElementById(`sfopts_${sfId}`);
    
    container.style.display = (val === 'dropdown' || val === 'sum') ? 'block' : 'none';
    if(groupContainer) groupContainer.style.display = (val === 'group') ? 'block' : 'none';

    if(val === 'sum') {
        labelEl.innerText = "कोणत्या प्रश्नांची बेरीज करायची?";
        inputEl.placeholder = "उदा. रुग्ण - पुरुष, रुग्ण - स्त्री";
    } else if(val === 'dropdown') {
        labelEl.innerText = "पर्याय लिहा (स्वल्पविरामाने वेगळे करा):";
        inputEl.placeholder = "उदा. होय, नाही";
    }
}

function toggleSubSubFieldOptions(ssfId, val) {
    let container = document.getElementById(`ssfopts_container_${ssfId}`);
    let inputEl = document.getElementById(`ssfopts_${ssfId}`);
    container.style.display = (val === 'dropdown' || val === 'sum') ? 'block' : 'none';
    if(val === 'sum') inputEl.placeholder = "बेरीज सूत्र (उदा. गट - उपगट - प्रश्न)...";
    else if(val === 'dropdown') inputEl.placeholder = "पर्याय लिहा (उदा. होय, नाही)...";
}

function addField(data = null) {
    const id = Math.floor(Math.random() * 1000000); 
    const html = `
        <div class="field-card" id="f_${id}">
            <button type="button" class="btn-remove" onclick="this.parentElement.remove()">X</button>
            <div style="display:flex; flex-direction:column; gap:5px; margin-top: 5px;">
                <label style="font-size:14px; color:#555; font-weight:bold;">मुख्य प्रश्नाचे नाव:</label>
                <input type="text" placeholder="उदा. ब्लिचिंग पावडर" class="fname" style="width:100%; margin-top:0;" value="${data ? data.label : ''}">
            </div>
            <select class="ftype" onchange="toggleFieldOptions('${id}', this.value)" style="margin-top:10px; background:#f0f8ff;">
                <option value="text" ${data && data.type==='text'?'selected':''}>Text (साधा मजकूर)</option>
                <option value="number" ${data && data.type==='number'?'selected':''}>Number (आकडे)</option>
                <option value="date" ${data && data.type==='date'?'selected':''}>Date (तारीख)</option>
                <option value="dropdown" ${data && data.type==='dropdown'?'selected':''}>Dropdown (पर्याय)</option>
                <option value="group" ${data && data.type==='group'?'selected':''}>Group / गट (उप-प्रश्न)</option>
                <option value="sum" ${data && data.type==='sum'?'selected':''}>Total / बेरीज (Auto Sum)</option>
            </select>
            <div id="opts_${id}" style="display:${data && (data.type==='dropdown' || data.type==='sum') ? 'block' : 'none'}; margin-top:10px; background:#ffe4b5; padding:10px; border-radius:5px;">
                <label style="font-size:12px; font-weight:bold; color:#d35400;">
                    ${data && data.type==='sum' ? 'कोणत्या प्रश्नांची बेरीज करायची?' : 'येथे पर्याय लिहा (स्वल्पविरामाने , वेगळे करा):'}
                </label>
                <input type="text" placeholder="${data && data.type==='sum' ? 'उदा. 0-5 वर्षे मुले, 0-5 वर्षे मुली' : 'उदा. होय, नाही'}" class="foptions" value="${data ? (data.options || '') : ''}">
            </div>
            <div id="group_${id}" style="display:${data && data.type==='group' ? 'block' : 'none'}; margin-top:10px; background:#eef; padding:10px; border-radius:5px; border-left: 3px solid var(--primary);">
                <label style="font-size:13px; font-weight:bold; color:var(--primary);">या गटातील उप-प्रश्न जोडा:</label>
                <div id="subfields_container_${id}"></div>
                <button type="button" onclick="addSubFieldUI('${id}')" style="margin-top:10px; font-size:13px; background:#fff; border:1px solid #ccc; padding:6px; cursor:pointer; width:100%;">+ उप-प्रश्न जोडा</button>
            </div>
        </div>`;
    document.getElementById('fieldsList').insertAdjacentHTML('beforeend', html);
    if(data && data.type === 'group' && data.subFields) { data.subFields.forEach(sf => addSubFieldUI(id, sf)); } 
    else if (data && data.type === 'group') { addSubFieldUI(id); }
}

function addSubFieldUI(parentId, sfData = null) {
    const sfId = Math.floor(Math.random() * 1000000);
    const html = `
        <div class="sub-field-item" id="sf_${sfId}">
            <button type="button" class="btn-remove" onclick="this.parentElement.remove()">X</button>
            <input type="text" placeholder="उप-प्रश्नाचे नाव" class="sfname" style="width: 100%; margin-top: 5px;" value="${sfData ? sfData.label : ''}">
            <select class="sftype" style="width: 100%; margin-top: 5px;" onchange="toggleSubFieldOptions('${sfId}', this.value)">
                <option value="text" ${sfData && sfData.type==='text'?'selected':''}>Text (मजकूर)</option>
                <option value="number" ${sfData && sfData.type==='number'?'selected':''}>Number (आकडे)</option>
                <option value="date" ${sfData && sfData.type==='date'?'selected':''}>Date (तारीख)</option>
                <option value="dropdown" ${sfData && sfData.type==='dropdown'?'selected':''}>Dropdown (पर्याय)</option>
                <option value="sum" ${sfData && sfData.type==='sum'?'selected':''}>Total / बेरीज (Auto Sum)</option>
                <option value="group" ${sfData && sfData.type==='group'?'selected':''}>Group / उप-गट (Sub-Group)</option>
            </select>
            <div id="sfopts_container_${sfId}" style="display:${sfData && (sfData.type==='dropdown' || sfData.type==='sum') ? 'block' : 'none'}; margin-top: 5px; background:#fff3e0; padding:8px; border-radius:4px;">
                <label id="sflabel_${sfId}" style="font-size:11px; font-weight:bold; color:#d35400;">
                    ${sfData && sfData.type==='sum' ? 'कोणत्या प्रश्नांची बेरीज करायची?' : 'पर्याय लिहा (स्वल्पविरामाने वेगळे करा):'}
                </label>
                <input type="text" id="sfopts_${sfId}" placeholder="..." class="sfopts" style="width: 100%; margin-top: 2px;" value="${sfData ? (sfData.options || '') : ''}">
            </div>
            <div id="ssf_group_${sfId}" style="display:${sfData && sfData.type==='group' ? 'block' : 'none'}; margin-top:10px; background:#e0f7fa; padding:10px; border-radius:5px; border-left: 3px solid #00acc1;">
                <label style="font-size:12px; font-weight:bold; color:#00838f;">या उप-गटातील प्रश्न जोडा:</label>
                <div id="subsubfields_container_${sfId}"></div>
                <button type="button" onclick="addSubSubFieldUI('${sfId}')" style="margin-top:10px; font-size:12px; background:#fff; border:1px solid #ccc; padding:4px; cursor:pointer; width:100%;">+ उप-गटातील प्रश्न जोडा</button>
            </div>
        </div>
    `;
    document.getElementById('subfields_container_' + parentId).insertAdjacentHTML('beforeend', html);
    
    if(sfData && sfData.type === 'group' && sfData.subFields) {
        sfData.subFields.forEach(ssf => addSubSubFieldUI(sfId, ssf));
    } else if (sfData && sfData.type === 'group') {
        addSubSubFieldUI(sfId);
    }
}

function addSubSubFieldUI(parentId, ssfData = null) {
    const ssfId = Math.floor(Math.random() * 1000000);
    const html = `
        <div class="sub-sub-field-item" id="ssf_${ssfId}" style="background:white; padding:8px; margin-top:5px; border:1px solid #b2ebf2; border-radius:4px; position:relative; padding-top:25px;">
            <button type="button" class="btn-remove" style="padding:2px 6px !important; font-size:10px !important; top:4px !important; right:4px !important;" onclick="this.parentElement.remove()">X</button>
            <input type="text" placeholder="प्रश्नाचे नाव" class="ssfname" style="width: 100%; font-size:13px; padding:6px;" value="${ssfData ? ssfData.label : ''}">
            <select class="ssftype" style="width: 100%; font-size:13px; padding:6px; margin-top: 5px;" onchange="toggleSubSubFieldOptions('${ssfId}', this.value)">
                <option value="text" ${ssfData && ssfData.type==='text'?'selected':''}>Text</option>
                <option value="number" ${ssfData && ssfData.type==='number'?'selected':''}>Number</option>
                <option value="date" ${ssfData && ssfData.type==='date'?'selected':''}>Date</option>
                <option value="dropdown" ${ssfData && ssfData.type==='dropdown'?'selected':''}>Dropdown</option>
                <option value="sum" ${ssfData && ssfData.type==='sum'?'selected':''}>Total (Sum)</option>
            </select>
            <div id="ssfopts_container_${ssfId}" style="display:${ssfData && (ssfData.type==='dropdown' || ssfData.type==='sum') ? 'block' : 'none'}; margin-top: 5px; background:#f3e5f5; padding:6px; border-radius:4px;">
                <input type="text" id="ssfopts_${ssfId}" placeholder="${ssfData && ssfData.type==='sum' ? 'बेरीज सूत्र...' : 'पर्याय...'}" class="ssfopts" style="width: 100%; font-size:12px; padding:4px;" value="${ssfData ? (ssfData.options || '') : ''}">
            </div>
        </div>
    `;
    document.getElementById('subsubfields_container_' + parentId).insertAdjacentHTML('beforeend', html);
}

function getFieldsData() {
    let fields = [];
    document.querySelectorAll('.field-card').forEach(el => {
        let fieldObj = { label: el.querySelector('.fname').value, type: el.querySelector('.ftype').value, options: el.querySelector('.foptions') ? el.querySelector('.foptions').value : "" };
        if (fieldObj.type === 'group') {
            fieldObj.subFields = [];
            el.querySelectorAll('.sub-field-item').forEach(sfEl => {
                let sfObj = { label: sfEl.querySelector('.sfname').value, type: sfEl.querySelector('.sftype').value, options: sfEl.querySelector('.sfopts') ? sfEl.querySelector('.sfopts').value : "" };
                if(sfObj.type === 'group') {
                    sfObj.subFields = [];
                    sfEl.querySelectorAll('.sub-sub-field-item').forEach(ssfEl => {
                        sfObj.subFields.push({ label: ssfEl.querySelector('.ssfname').value, type: ssfEl.querySelector('.ssftype').value, options: ssfEl.querySelector('.ssfopts') ? ssfEl.querySelector('.ssfopts').value : "" });
                    });
                }
                fieldObj.subFields.push(sfObj);
            });
        }
        fields.push(fieldObj);
    });
    return fields;
}

async function saveFullForm() {
    const name = document.getElementById('newFormName').value;
    const type = document.getElementById('newFormType').value;
    const allowedRoles = getSelectedRoles(); 
    
    if(!name) { alert("फॉर्मचे नाव टाका"); return; }
    if(allowedRoles === "" && !document.getElementById('roleAll').checked) { alert("कृपया किमान एक Role निवडा."); return; }

    const payload = { name, type, allowedRoles, fields: getFieldsData(), adminMobile: user.mobile };
    const r = await fetch(GAS_URL, { method: "POST", body: JSON.stringify({action:"createForm", payload}) });
    const d = await r.json();
    if(d.success) { alert(d.message); location.reload(); }
}

async function updateExistingForm() {
    const formID = document.getElementById('editFormID').value;
    const name = document.getElementById('newFormName').value;
    const type = document.getElementById('newFormType').value;
    const allowedRoles = getSelectedRoles(); 

    if(allowedRoles === "" && !document.getElementById('roleAll').checked) { alert("कृपया किमान एक Role निवडा."); return; }

    const payload = { formID, name, type, allowedRoles, fields: getFieldsData() };
    const r = await fetch(GAS_URL, { method: "POST", body: JSON.stringify({action:"updateForm", payload}) });
    const d = await r.json();
    if(d.success) { alert(d.message); location.reload(); }
}
</script>

<script>
function getTotalsRow(data, headers, showIndices) {
    let totals = Array(headers.length).fill(0);
    let isNumericCol = Array(headers.length).fill(false);
    for(let c of showIndices) {
        let colName = headers[c] || "";
        if(colName.includes("मोबाईल") || colName.includes("क्रमांक") || colName === "तारीख" || colName === "महिना" || colName === "वर्ष" || colName === "उपकेंद्र" || colName === "गाव" || colName === "ग्रामपंचायत" || colName.includes("नाव") || colName.includes("स्तर")) {
            continue; 
        }
        let isNum = false; 
        let colSum = 0;
        for(let r=1; r<data.length; r++) {
            let val = String(data[r][c] || "").trim();
            if(val !== "" && val !== "-") {
                if(!isNaN(val)) {
                    isNum = true;
                    colSum += parseFloat(val);
                } else {
                    isNum = false;
                    break; 
                }
            }
        }
        if(isNum) {
            isNumericCol[c] = true;
            totals[c] = colSum;
        }
    }
    return { totals, isNumericCol };
}

async function fetchReportData() {
    const formID = document.getElementById('reportFormSelect').value;
    const selMonth = document.getElementById('reportMonth').value;
    const selYear = document.getElementById('reportYear').value;

    if(!formID) { alert("कृपया अहवाल निवडा"); return; }
    
    let filterRole = "सर्व";
    if(user && user.role === "Admin" && document.getElementById('reportRoleFilter')) {
        filterRole = document.getElementById('reportRoleFilter').value;
    }

    document.getElementById('reportLoader').style.display = "block";
    document.getElementById('reportContentArea').classList.add('hidden');
    document.getElementById('reportTableContainer').innerHTML = "";
    
    try {
        const payload = { 
            formID: formID, 
            role: user.role, 
            subcenter: user.subcenter, 
            mobileNo: user.mobile, 
            filterRole: filterRole 
        };
        
        const r = await fetch(GAS_URL, { method: "POST", body: JSON.stringify({action:"getReportData", payload}) });
        const responseText = await r.text();
        
        if(responseText.trim().startsWith("<")) throw new Error("गुगल सर्व्हर ब्लॉक करत आहे.");
        
        const d = JSON.parse(responseText);
        document.getElementById('reportLoader').style.display = "none";
        
        if(d.success) {
            if(d.reports && d.reports.length > 0) {
                let finalReports = [];
                d.reports.forEach(rep => {
                    let headers = rep.data[0];
                    if(!headers) return; 
                    
                    let monthIdx = headers.indexOf("महिना");
                    let yearIdx = headers.indexOf("वर्ष");

                    let fData = [headers];
                    let dataRows = [];
                    
                    for(let i=1; i<rep.data.length; i++) {
                        let row = rep.data[i];
                        let matchMonth = (selMonth === "सर्व" || String(row[monthIdx]).trim() === String(selMonth).trim());
                        let matchYear = (selYear === "सर्व" || String(row[yearIdx]).trim() === String(selYear).trim());
                        if(matchMonth && matchYear) dataRows.push(row);
                    }
                    
                    fData = fData.concat(dataRows);
                    finalReports.push({ formName: rep.formName, data: fData });
                });

                if(finalReports.length > 0) {
                    currentReports = finalReports;
                    renderMultipleTables(finalReports, selMonth, selYear);
                    document.getElementById('reportContentArea').classList.remove('hidden');
                } else { alert("निवडलेल्या महिना आणि वर्षासाठी कोणताही डेटा उपलब्ध नाही."); }
            } else { alert("अद्याप कोणतीही माहिती उपलब्ध नाही."); }
        } else { alert("सर्व्हरकडून माहिती मिळाली नाही."); }
    } catch(e) {
        console.error("Report Fetch Error:", e);
        document.getElementById('reportLoader').style.display = "none";
        alert("एरर: " + e.message);
    }
}

function renderMultipleTables(reports, month, year) {
    let container = document.getElementById('reportTableContainer');
    let html = "";
    let periodText = (month === 'सर्व' && year === 'सर्व') ? 'सर्व महिने' : `${month} ${year !== 'सर्व' ? year : ''}`;

    reports.forEach(rep => {
        let data = rep.data;
        let headers = data[0];
        let showIndices = [];
        headers.forEach((h, i) => { if(!CONFIG.hiddenColumns.includes(h)) showIndices.push(i); });
        let colCount = showIndices.length + 1; 
        
        const formObj = masterData.forms.find(x => x.FormName === rep.formName);
        const isStats = formObj && formObj.FormType === 'Stats';

        let subCenterIdx = headers.indexOf("उपकेंद्र");
        let nameIdx = headers.indexOf("कर्मचाऱ्याचे नाव");
        let mobIdx = headers.indexOf("मोबाईल क्र.");
        let villageIdx = headers.indexOf("गाव");

        let h1Arr = []; let h2Arr = [];
        showIndices.forEach(idx => {
            let parts = headers[idx].split(" - ");
            h1Arr.push(parts[0]);
            h2Arr.push(parts.length > 1 ? parts.slice(1).join(" - ") : "");
        });
        let hasSubQ = h2Arr.some(sub => sub !== "");

        let dataRows = data.slice(1);
        let groups = {};
        if(dataRows.length === 0) {
            groups["All"] = [];
        } else {
            dataRows.forEach(row => {
                let sc = subCenterIdx > -1 ? String(row[subCenterIdx] || "").trim() : "Unknown";
                let mob = mobIdx > -1 ? String(row[mobIdx] || "").trim() : "Unknown";
                let ename = nameIdx > -1 ? String(row[nameIdx] || "").trim() : mob; 
                if(ename === "undefined" || ename === "") ename = mob;
                let key = sc + "###" + ename; 
                if(!groups[key]) groups[key] = [];
                groups[key].push(row);
            });
        }

        let groupKeys = Object.keys(groups).sort();
        
        html += `<div style="background:white; padding:15px; border-radius:8px; box-shadow:0 2px 5px rgba(0,0,0,0.1); margin-bottom:20px;">`;
        html += `<h3 style="color:var(--primary); border-bottom:2px solid var(--primary); padding-bottom:10px;">${rep.formName} अहवाल</h3>`;

        groupKeys.forEach((gKey) => {
            let gRows = groups[gKey];
            let [sc, ename] = gKey.split("###");
            
            if(villageIdx > -1) {
                gRows.sort((a, b) => {
                    let vilA = String(a[villageIdx] || "").trim();
                    let vilB = String(b[villageIdx] || "").trim();
                    return vilA.localeCompare(vilB);
                });
            }
            
            if(dataRows.length > 0) {
                html += `<div style="background:#e8f4f8; padding:10px; border-left:5px solid var(--secondary); margin-top:20px; font-weight:bold; color:#0056b3; border-radius:4px;">
                            उपकेंद्र: <span style="color:#333;">${sc}</span> &nbsp;&nbsp;|&nbsp;&nbsp; कर्मचारी: <span style="color:#333;">${ename}</span> &nbsp;&nbsp;|&nbsp;&nbsp; अहवाल महिना: <span style="color:#333;">${periodText}</span>
                         </div>`;
            }

            html += `<table class="report-table" style="margin-top:0; border-top:none;"><thead>`;
            
            if(hasSubQ) {
                let row1HTML = `<tr><th rowspan="2" style="background:#f4b400; color:#000; border:1px solid #ddd; vertical-align:middle;">अ.क्र.</th>`;
                let row2HTML = `<tr>`;
                let c = 0;
                while(c < h1Arr.length) {
                    let end = c;
                    while(end + 1 < h1Arr.length && h1Arr[end+1] === h1Arr[c] && h2Arr[end+1] !== "") { end++; }
                    let colspan = end - c + 1;
                    if(colspan > 1) {
                        row1HTML += `<th colspan="${colspan}" style="background:#f4b400; color:#000; border:1px solid #ddd;">${h1Arr[c]}</th>`;
                        for(let k=c; k<=end; k++) { row2HTML += `<th style="background:#ffe082; color:#000; border:1px solid #ddd; font-size: 13px;">${h2Arr[k]}</th>`; }
                    } else {
                        if(h2Arr[c] === "") { row1HTML += `<th rowspan="2" style="background:#f4b400; color:#000; border:1px solid #ddd; vertical-align:middle;">${h1Arr[c]}</th>`; } 
                        else {
                            row1HTML += `<th style="background:#f4b400; color:#000; border:1px solid #ddd;">${h1Arr[c]}</th>`;
                            row2HTML += `<th style="background:#ffe082; color:#000; border:1px solid #ddd; font-size: 13px;">${h2Arr[c]}</th>`;
                        }
                    }
                    c = end + 1;
                }
                html += row1HTML + `</tr>` + row2HTML + `</tr>`;
            } else {
                html += `<tr><th style="background:#f4b400; color:#000; border:1px solid #ddd;">अ.क्र.</th>`;
                h1Arr.forEach(h => { html += `<th style="background:#f4b400; color:#000; border:1px solid #ddd;">${h}</th>`; });
                html += `</tr>`;
            }
            html += `</thead><tbody>`;

            if(gRows.length === 0) { 
                html += `<tr><td colspan="${colCount}" style="color:red; font-weight:bold; font-size:16px;">निरंक (Nil)</td></tr>`; 
            } else {
                gRows.forEach((row, i) => {
                    html += `<tr><td>${i+1}</td>`;
                    showIndices.forEach(idx => { html += `<td>${row[idx] || "-"}</td>`; });
                    html += `</tr>`;
                });

                if(isStats && gRows.length > 0) { 
                    let pseudoData = [headers].concat(gRows);
                    let { totals, isNumericCol } = getTotalsRow(pseudoData, headers, showIndices);
                    html += `<tr style="background:#d4edda; font-weight:bold; color:#155724;"><td>एकूण</td>`;
                    showIndices.forEach(idx => {
                        if(isNumericCol[idx]) html += `<td>${totals[idx]}</td>`;
                        else html += `<td>-</td>`;
                    });
                    html += `</tr>`;
                }
            }
            html += `</tbody></table>`;
        });
        html += `</div>`;
    });
    container.innerHTML = html;
}

function downloadConsolidatedExcel() {
    if(currentReports.length === 0) return;
    let month = document.getElementById('reportMonth').value;
    let year = document.getElementById('reportYear').value;
    let filterRole = document.getElementById('reportRoleFilter') ? document.getElementById('reportRoleFilter').value : "सर्व";
    let periodText = (month === 'सर्व' && year === 'सर्व') ? 'सर्व महिने' : `${month} ${year}`;
    if(user.role === 'Admin' && filterRole !== 'सर्व') periodText += ` (${filterRole})`;
    
    let wb = XLSX.utils.book_new();

    currentReports.forEach((rep, index) => {
        let data = rep.data;
        let headers = data[0];
        let showIndices = [];
        headers.forEach((h, i) => { if(!CONFIG.hiddenColumns.includes(h)) showIndices.push(i); });
        let colCount = showIndices.length + 1; 
        
        const formObj = masterData.forms.find(x => x.FormName === rep.formName);
        const isStats = formObj && formObj.FormType === 'Stats';

        let subCenterIdx = headers.indexOf("उपकेंद्र");
        let nameIdx = headers.indexOf("कर्मचाऱ्याचे नाव");
        let mobIdx = headers.indexOf("मोबाईल क्र.");
        let villageIdx = headers.indexOf("गाव");

        let h1Arr = []; let h2Arr = [];
        showIndices.forEach(idx => {
            let parts = headers[idx].split(" - ");
            h1Arr.push(parts[0]); h2Arr.push(parts.length > 1 ? parts.slice(1).join(" - ") : "");
        });
        let hasSubQ = h2Arr.some(sub => sub !== "");

        let sheetData = [];
        let merges = [];
        
        sheetData.push([`${rep.formName} अहवाल`]);
        sheetData.push(Array(colCount).fill("")); 
        merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: colCount - 1 } });
        
        let dataRows = data.slice(1);
        let groups = {};
        if(dataRows.length === 0) {
            groups["All"] = [];
        } else {
            dataRows.forEach(row => {
                let sc = subCenterIdx > -1 ? String(row[subCenterIdx] || "").trim() : "Unknown";
                let mob = mobIdx > -1 ? String(row[mobIdx] || "").trim() : "Unknown";
                let ename = nameIdx > -1 ? String(row[nameIdx] || "").trim() : mob; 
                if(ename === "undefined" || ename === "") ename = mob;
                let key = sc + "###" + ename;
                if(!groups[key]) groups[key] = [];
                groups[key].push(row);
            });
        }

        let groupKeys = Object.keys(groups).sort();
        let currentR = 2; 
        
        let groupHeaderRows = [];
        let headerRowIndices = [];
        let subHeaderRowIndices = [];
        let totalRowIndices = [];

        groupKeys.forEach(gKey => {
            let gRows = groups[gKey];
            let [sc, ename] = gKey.split("###");
            
            if(villageIdx > -1) {
                gRows.sort((a, b) => {
                    let vilA = String(a[villageIdx] || "").trim();
                    let vilB = String(b[villageIdx] || "").trim();
                    return vilA.localeCompare(vilB);
                });
            }

            if(dataRows.length > 0) {
                let gHeaderArr = [`उपकेंद्र: ${sc}   |   कर्मचारी: ${ename}   |   महिना: ${periodText}`];
                for(let x=1; x<colCount; x++) gHeaderArr.push("");
                sheetData.push(gHeaderArr);
                merges.push({ s: { r: currentR, c: 0 }, e: { r: currentR, c: colCount - 1 } });
                groupHeaderRows.push(currentR);
                currentR++;
            }

            let headerRow1 = ["अ.क्र."].concat(h1Arr);
            sheetData.push(headerRow1);
            headerRowIndices.push(currentR);
            currentR++;

            let dataStartRow = currentR;
            if(hasSubQ) {
                let headerRow2 = [""].concat(h2Arr);
                sheetData.push(headerRow2);
                subHeaderRowIndices.push(currentR);
                dataStartRow = currentR + 1;
                
                merges.push({ s: { r: currentR - 1, c: 0 }, e: { r: currentR, c: 0 } }); 
                let c = 0;
                while(c < h1Arr.length) {
                    let end = c;
                    while(end + 1 < h1Arr.length && h1Arr[end+1] === h1Arr[c] && h2Arr[end+1] !== "") { end++; }
                    if(end > c) { merges.push({ s: { r: currentR - 1, c: c + 1 }, e: { r: currentR - 1, c: end + 1 } }); } 
                    else if(h2Arr[c] === "") { merges.push({ s: { r: currentR - 1, c: c + 1 }, e: { r: currentR, c: c + 1 } }); }
                    c = end + 1;
                }
                currentR++;
            }

            if(gRows.length === 0) {
                let nilRow = Array(colCount).fill("");
                nilRow[0] = "निरंक (Nil)";
                sheetData.push(nilRow);
                merges.push({ s: { r: currentR, c: 0 }, e: { r: currentR, c: colCount - 1 } });
                currentR++;
            } else {
                gRows.forEach((row, i) => {
                    let rowData = [i+1];
                    showIndices.forEach(idx => rowData.push(row[idx] || "-"));
                    sheetData.push(rowData);
                    currentR++;
                });

                if(isStats) {
                    let pseudoData = [headers].concat(gRows);
                    let { totals, isNumericCol } = getTotalsRow(pseudoData, headers, showIndices);
                    let totalRow = ["एकूण"];
                    showIndices.forEach(idx => {
                        if(isNumericCol[idx]) totalRow.push(totals[idx]);
                        else totalRow.push("-");
                    });
                    sheetData.push(totalRow);
                    totalRowIndices.push(currentR);
                    currentR++;
                }
            }
            
            sheetData.push(Array(colCount).fill("")); 
            currentR++;
        });

        let ws = XLSX.utils.aoa_to_sheet(sheetData);
        ws["!merges"] = merges;

        for(let R=0; R<sheetData.length; R++) {
            for(let C=0; C<colCount; C++) {
                let cellRef = XLSX.utils.encode_cell({r: R, c: C});
                if(!ws[cellRef]) continue; 
                let cellStyle = { font: { name: "Arial", sz: 11, color: { rgb: "000000" } }, alignment: { vertical: "center", horizontal: "center", wrapText: true } };

                if(R === 0) {
                    cellStyle.fill = { fgColor: { rgb: "00705A" } };
                    cellStyle.font = { name: "Arial", sz: 16, bold: true, color: { rgb: "FFFFFF" } };
                } else if(groupHeaderRows.includes(R)) { 
                    cellStyle.fill = { fgColor: { rgb: "CDE4EC" } };
                    cellStyle.font = { name: "Arial", sz: 12, bold: true, color: { rgb: "0056B3" } };
                    cellStyle.alignment = { vertical: "center", horizontal: "left", wrapText: true };
                    cellStyle.border = { top: { style: "medium", color: { rgb: "0056B3" } }, bottom: { style: "medium", color: { rgb: "0056B3" } } };
                } else if(headerRowIndices.includes(R) || subHeaderRowIndices.includes(R)) {
                    cellStyle.fill = { fgColor: { rgb: headerRowIndices.includes(R) ? "F4B400" : "FFE082" } }; 
                    cellStyle.font = { name: "Arial", sz: 11, bold: true, color: { rgb: "000000" } };
                    cellStyle.border = { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" } };
                } else if(totalRowIndices.includes(R)) { 
                    cellStyle.fill = { fgColor: { rgb: "D4EDDA" } };
                    cellStyle.font = { name: "Arial", sz: 11, bold: true, color: { rgb: "155724" } };
                    cellStyle.border = { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" } };
                } else if(sheetData[R][0] !== "") { 
                    if(sheetData[R][0] === "निरंक (Nil)") {
                        cellStyle.font = { name: "Arial", sz: 12, bold: true, color: { rgb: "FF0000" } };
                    } else {
                        cellStyle.border = { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" } };
                    }
                }
                ws[cellRef].s = cellStyle; 
            }
        }
        
        let wscols = [{ wch: 8 }]; 
        for(let c=1; c<colCount; c++) wscols.push({ wch: 20 }); 
        ws["!cols"] = wscols;

        let safeSheetName = rep.formName.replace(/[\\\/\?\*\[\]\:]/g, "").substring(0, 31);
        if(!safeSheetName) safeSheetName = "Sheet" + (index + 1);
        XLSX.utils.book_append_sheet(wb, ws, safeSheetName);
    });

    let fileName = `मासिक_अहवाल_${user.subcenter}_${periodText}.xlsx`;
    XLSX.writeFile(wb, fileName);
}
</script>
</body>
</html>
