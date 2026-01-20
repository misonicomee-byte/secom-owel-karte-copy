let karteData = null, allEntries = [], tabId = null;
const elements = {};

document.addEventListener('DOMContentLoaded', async () => {
    ['loading','loadingText','loadingProgress','error','patientInfo','loadSection','rangeSection','startEntry','endEntry','selectedCount','dateRange','copyBtn','successMessage','includePatientInfo','includeVitals','totalEntries','currentPageInfo','totalPages','loadAllBtn'].forEach(id => {
          elements[id] = document.getElementById(id.replace(/([A-Z])/g, '-$1').toLowerCase());
    });
    elements.loading = document.getElementById('loading');
    elements.loadingText = document.getElementById('loading-text');
    elements.loadingProgress = document.getElementById('loading-progress');
    elements.error = document.getElementById('error-message');
    elements.patientInfo = document.getElementById('patient-info');
    elements.loadSection = document.getElementById('load-section');
    elements.rangeSection = document.getElementById('range-section');
    elements.startEntry = document.getElementById('start-entry');
    elements.endEntry = document.getElementById('end-entry');
    elements.selectedCount = document.getElementById('selected-count');
    elements.dateRange = document.getElementById('date-range');
    elements.copyBtn = document.getElementById('copy-btn');
    elements.successMessage = document.getElementById('success-message');
    elements.includePatientInfo = document.getElementById('include-patient-info');
    elements.includeVitals = document.getElementById('include-vitals');
    elements.totalEntries = document.getElementById('total-entries');
    elements.currentPageInfo = document.getElementById('current-page-info');
    elements.totalPages = document.getElementById('total-pages');
    elements.loadAllBtn = document.getElementById('load-all-btn');

                            showLoading('初期化中...');
    try {
          const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
          tabId = tab.id;
          if (!tab.url.includes('secom-owel.jp/cl/kanja/karte')) { showError('SECOM OWELのカルテページを開いてください'); return; }
          const results = await chrome.scripting.executeScript({ target: { tabId }, func: getCurrentPageInfo });
          karteData = results[0].result;
          if (!karteData) { showError('ページ情報を取得できませんでした'); return; }
          showPatientInfo(karteData.patientInfo);
          elements.currentPageInfo.textContent = 'ページ' + karteData.currentPage;
          elements.totalPages.textContent = karteData.totalPages;
          allEntries = karteData.entries;
          elements.loadSection.classList.remove('hidden');
          elements.loadAllBtn.addEventListener('click', loadAllPages);
          hideLoading();
    } catch (e) { showError('エラー: ' + e.message); }
});

function getCurrentPageInfo() {
    const result = { patientInfo: {}, entries: [], currentPage: 1, totalPages: 1 };
    const textNodes = Array.from(document.querySelectorAll('*')).filter(el => el.children.length === 0);
    const kanjiNameEl = textNodes.find(el => (el.textContent?.trim() || '').match(/^[一-龯ぁ-んァ-ン]+[\s　]+[一-龯ぁ-んァ-ン]+$/) && el.textContent.length < 20);
    if (kanjiNameEl) result.patientInfo.name = kanjiNameEl.textContent.trim();
    const patientNumEl = textNodes.find(el => el.textContent?.match(/^患者番号[：:]\s*\d+$/));
    if (patientNumEl) { const m = patientNumEl.textContent.match(/(\d+)/); if (m) result.patientInfo.patientNumber = m[1]; }
    const ageEl = textNodes.find(el => el.textContent?.match(/^\d+歳$/));
    if (ageEl) result.patientInfo.age = ageEl.textContent.replace('歳', '');
    const careLevelEl = textNodes.find(el => el.textContent?.match(/^要介護[１-５1-5]$/));
    if (careLevelEl) result.patientInfo.careLevel = careLevelEl.textContent.trim();
    if (typeof karte_get_order_page_no === 'function') result.currentPage = karte_get_order_page_no();
    const pageLinks = jQuery('.pagination-margin-left li a').filter(function() { return !isNaN(parseInt(jQuery(this).text().trim())); }).map(function() { return parseInt(jQuery(this).text().trim()); }).get();
    result.totalPages = Math.max(...pageLinks, 1);
    result.entries = extractCurrentPageEntries();
    return result;
}

function extractCurrentPageEntries() {
    const entries = [];
    document.querySelectorAll('.karte-order-panel-header-no').forEach((header, index) => {
          const entry = { index: index + 1, headId: header.id, date: header.querySelector('[id*="date-hide"]')?.textContent || '', time: header.querySelector('[id*="time-hide"]')?.textContent || '', orderId: header.querySelector('[id*="id-hide"]')?.textContent || '', contents: [], vitals: null };
          header.querySelectorAll('.karte-order-panel-table-size').forEach(table => {
                  let text = table.textContent?.trim().replace(/\s+/g, ' ') || '';
                  if (text.length > 0) {
                            const vm = text.match(/血圧[\s　]*(\d+)\s*\/\s*(\d+)\s*mmHg.*?心拍数[\s　]*(\d+).*?体温[\s　]*([\d.]+)/);
                            if (vm) entry.vitals = { bp: vm[1]+'/'+vm[2], hr: vm[3], temp: vm[4] };
                            entry.contents.push(text);
                  }
          });
          if (entry.date) entries.push(entry);
    });
    return entries;
}

async function loadAllPages() {
    const totalPages = parseInt(elements.totalPages.textContent);
    showLoading('全' + totalPages + 'ページを読み込み中...');
    elements.loadSection.classList.add('hidden');
    allEntries = [];
    try {
          for (let page = 1; page <= totalPages; page++) {
                  elements.loadingProgress.textContent = 'ページ ' + page + ' / ' + totalPages;
                  await chrome.scripting.executeScript({ target: { tabId }, func: navigateToPage, args: [page] });
                  await new Promise(r => setTimeout(r, 1500));
                  const results = await chrome.scripting.executeScript({ target: { tabId }, func: extractCurrentPageEntries });
                  results[0].result.forEach((entry, idx) => { entry.globalIndex = allEntries.length + idx + 1; allEntries.push(entry); });
          }
          await chrome.scripting.executeScript({ target: { tabId }, func: navigateToPage, args: [1] });
          hideLoading();
          showRangeSelector();
    } catch (e) { showError('読み込みエラー: ' + e.message); }
}

function navigateToPage(pageNum) {
    const pageLink = jQuery('.pagination-margin-left li a').filter(function() { return jQuery(this).text().trim() === String(pageNum); });
    if (pageLink.length > 0) pageLink[0].click();
    else if (pageNum === 1) jQuery('.pagination-margin-left li a').filter(function() { return jQuery(this).text().trim() === '最初'; })[0]?.click();
}

function showPatientInfo(info) {
    elements.patientInfo.querySelector('.patient-name').textContent = info.name || '患者名不明';
    const details = [];
    if (info.patientNumber) details.push('ID: ' + info.patientNumber);
    if (info.age) details.push(info.age + '歳');
    if (info.careLevel) details.push(info.careLevel);
    elements.patientInfo.querySelector('.patient-details').textContent = details.join(' / ');
    elements.patientInfo.classList.remove('hidden');
}

function showRangeSelector() {
    elements.totalEntries.textContent = allEntries.length;
    elements.startEntry.innerHTML = '';
    elements.endEntry.innerHTML = '';
    allEntries.forEach((entry, index) => {
          const optionText = entry.globalIndex + '. ' + entry.date + ' ' + entry.time;
          elements.startEntry.innerHTML += '<option value="' + index + '">' + optionText + '</option>';
          elements.endEntry.innerHTML += '<option value="' + index + '">' + optionText + '</option>';
    });
    elements.startEntry.value = 0;
    elements.endEntry.value = allEntries.length - 1;
    elements.startEntry.addEventListener('change', updatePreview);
    elements.endEntry.addEventListener('change', updatePreview);
    document.getElementById('select-all').addEventListener('click', () => { elements.startEntry.value = 0; elements.endEntry.value = allEntries.length - 1; updatePreview(); });
    document.getElementById('select-latest').addEventListener('click', () => { elements.startEntry.value = 0; elements.endEntry.value = 0; updatePreview(); });
    document.getElementById('select-latest5').addEventListener('click', () => { elements.startEntry.value = 0; elements.endEntry.value = Math.min(4, allEntries.length - 1); updatePreview(); });
    document.getElementById('select-latest10').addEventListener('click', () => { elements.startEntry.value = 0; elements.endEntry.value = Math.min(9, allEntries.length - 1); updatePreview(); });
    elements.copyBtn.addEventListener('click', copySelectedRange);
    document.getElementById('download-btn').addEventListener('click', downloadSelectedRange);
    elements.rangeSection.classList.remove('hidden');
    updatePreview();
}

function updatePreview() {
    const startIdx = parseInt(elements.startEntry.value), endIdx = parseInt(elements.endEntry.value);
    elements.selectedCount.textContent = Math.abs(endIdx - startIdx) + 1;
    const minIdx = Math.min(startIdx, endIdx), maxIdx = Math.max(startIdx, endIdx);
    const startDate = allEntries[minIdx]?.date || '', endDate = allEntries[maxIdx]?.date || '';
    elements.dateRange.textContent = startDate === endDate ? '(' + startDate + ')' : '(' + startDate + ' 〜 ' + endDate + ')';
}

async function copySelectedRange() {
    const text = getOutputText();
    try { await navigator.clipboard.writeText(text); showSuccessMsg('クリップボードにコピーしました！'); } catch (e) { showError('コピーに失敗しました: ' + e.message); }
}

function downloadSelectedRange() {
    const text = getOutputText();
    const format = document.querySelector('input[name="output-format"]:checked').value;
    const patientName = (karteData?.patientInfo?.name || '患者').replace(/[\s　]/g, '');
    const dateStr = new Date().toISOString().slice(0, 10);
    const fileName = 'カルテ_' + patientName + '_' + dateStr + (format === 'json' ? '.json' : '.txt');
    const blob = new Blob([text], { type: (format === 'json' ? 'application/json' : 'text/plain') + ';charset=utf-8' });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = fileName; document.body.appendChild(a); a.click(); document.body.removeChild(a);
    showSuccessMsg('ダウンロードしました！');
}

function getOutputText() {
    const startIdx = parseInt(elements.startEntry.value), endIdx = parseInt(elements.endEntry.value);
    const selectedEntries = allEntries.slice(Math.min(startIdx, endIdx), Math.max(startIdx, endIdx) + 1);
    const includePatient = elements.includePatientInfo.checked, includeVitals = elements.includeVitals.checked;
    const format = document.querySelector('input[name="output-format"]:checked').value;
    if (format === 'json') {
          return JSON.stringify({ exportDate: new Date().toISOString(), patientInfo: includePatient ? karteData?.patientInfo : undefined, entries: selectedEntries.map(e => ({ date: e.date, time: e.time, orderId: e.orderId, vitals: includeVitals ? e.vitals : undefined, contents: e.contents })) }, null, 2);
    }
    let text = '';
    if (includePatient && karteData?.patientInfo) { text += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n【患者情報】\n氏名: ' + (karteData.patientInfo.name || '不明') + '\n'; if (karteData.patientInfo.patientNumber) text += '患者番号: ' + karteData.patientInfo.patientNumber + '\n'; if (karteData.patientInfo.age) text += '年齢: ' + karteData.patientInfo.age + '歳\n'; if (karteData.patientInfo.careLevel) text += '介護度: ' + karteData.patientInfo.careLevel + '\n'; text += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n'; }
    selectedEntries.forEach(entry => { text += '【' + entry.date + ' ' + entry.time + '】\n────────────────────────\n'; if (includeVitals && entry.vitals) text += '[バイタル] BP: ' + entry.vitals.bp + ' mmHg / HR: ' + entry.vitals.hr + ' 回/分 / BT: ' + entry.vitals.temp + ' ℃\n'; entry.contents.forEach(c => { text += c + '\n'; }); text += '\n'; });
    return text;
}

function showLoading(msg) { elements.loadingText.textContent = msg; elements.loadingProgress.textContent = ''; elements.loading.classList.remove('hidden'); }
function hideLoading() { elements.loading.classList.add('hidden'); }
function showError(msg) { elements.error.textContent = msg; elements.error.classList.remove('hidden'); hideLoading(); }
function showSuccessMsg(msg) { elements.successMessage.textContent = '✅ ' + msg; elements.successMessage.classList.remove('hidden'); setTimeout(() => { elements.successMessage.classList.add('hidden'); }, 2000); }
