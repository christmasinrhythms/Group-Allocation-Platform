const $ = (id) => document.getElementById(id);
const views = { tables: 'table', groups: 'group', roster: null };
let previousSeating = {};
let state, workbook, binding = 0, dragged, undo, selected, loading = false, writing = false, failed = false;
let history = new Map();
let active = 'tables';
let drawerOpen = false;

const copy = {
  connectTitle: ['連接 Excel', 'Connect Excel'],
  connectIntro: ['使用公司 Microsoft 365 帳戶連接有編輯權限嘅工作簿。', 'Use your work Microsoft 365 account to connect a workbook you can edit.'],
  workbookLink: ['Excel 工作簿連結', 'Excel workbook link'],
  setup: ['Microsoft 365 設定', 'Microsoft 365 settings'],
  tenant: ['目錄（租戶）ID', 'Directory (tenant) ID'],
  client: ['應用程式（用戶端）ID', 'Application (client) ID'],
  redirect: ['在 Microsoft Entra 將以下網址登記為 SPA 重新導向 URI。', 'Register this URL as the SPA redirect URI in Microsoft Entra.'],
  connect: ['登入並連接 Excel', 'Sign in and connect Excel'],
  disconnect: ['離開工作簿', 'Disconnect workbook'],
  demo: ['試用示範名單', 'Try demo roster'],
  demoMode: ['示範模式', 'Demo mode'],
  demoHelp: ['示範只使用虛構名單。試做嘅改動會在離開示範時清除。', 'The demo uses fictional names. Trial changes clear when you leave the demo.'],
  demoSaved: ['已更新示範名單。', 'Demo roster updated.'],
  updateDemo: ['更新示範名單', 'Update demo roster'],
  cloudMode: ['已連接 Excel 網頁版', 'Connected to Excel for the web'],
  cloudHelp: ['可同時開啟 Excel。Excel 自動儲存後，工作台會讀取更新。', 'Keep Excel open alongside the board. The board checks for updates after Excel autosaves.'],
  connecting: ['正在連接 Microsoft 365…', 'Connecting to Microsoft 365…'],
  language: ['語言', 'Language'], theme: ['顯示主題', 'Appearance'],
  light: ['淺色', 'Light'], dark: ['深色', 'Dark'],
  saveBook: ['儲存工作簿', 'Save workbook'], openExcel: ['開啟 Excel ↗', 'Open Excel ↗'],
  view: ['檢視', 'View'], tables: ['餐桌', 'Tables'], groups: ['遊戲組', 'Game groups'], roster: ['總名單', 'Roster'],
  tableField: ['餐桌', 'Table'], groupField: ['遊戲組', 'Game group'],
  search: ['搜尋姓名或部門', 'Name or department'], splitFilter: ['只看與舊同桌分開', 'Only separated tablemates'],
  undo: ['復原', 'Undo'], assignments: ['分組結果', 'Assignments'], loadingRoster: ['正在載入名單…', 'Loading roster…'],
  unassigned: ['未分配', 'Unassigned'], unassignedList: ['未分配名單', 'Unassigned list'],
  openDrawer: ['展開', 'Expand'], closeDrawer: ['收起未分配名單', 'Collapse unassigned list'],
  drawerHelp: ['拖動姓名分配。拖返呢個面板可移回名單。', 'Drag names out to assign them, or back here to unassign.'],
  dropTargets: ['拖放目的地', 'Drop destinations'], close: ['關閉', 'Close'],
  formerTablemates: ['去年同桌', 'Last year’s tablemates'],
  assignmentField: ['調整項目', 'Assignment type'], moveTo: ['移至', 'Move to'], swapWith: ['交換對象（選填）', 'Swap with (optional)'],
  updateExcel: ['更新 Excel', 'Update Excel'],
  tableNumber: ['第 {n} 桌', 'Table {n}'], groupNumber: ['第 {n} 組', 'Group {n}'],
  people: ['{n} 人', '{n} people'], onePerson: ['1 人', '1 person'], hostTable: ['主家席', 'Head table'],
  disconnected: ['Excel 同步中斷', 'Excel sync interrupted'], writingExcel: ['正在更新 Excel…', 'Updating Excel…'],
  readingExcel: ['正在讀取 Excel…', 'Reading Excel…'], paused: ['Excel 暫停同步', 'Excel sync paused'],
  live: ['Excel 即時同步', 'Live Excel sync'], unsaved: ['有未儲存更改', 'Unsaved changes'],
  syncedFile: ['已同步 Excel 檔案', 'Excel file synced'],
  syncLiveHelp: ['在 Excel 按 Enter 完成輸入後，工作台會自動更新。按 ⌘S 或「儲存工作簿」保留更改。', 'Press Enter after editing a cell to sync. Press ⌘S or select Save workbook to keep changes.'],
  syncFileHelp: ['網頁移位會更新連結的 Excel。', 'Moves update the linked Excel file.'],
  splitTables: ['分坐不同枱', 'Different tables'], splitGroups: ['不同遊戲組', 'Different groups'],
  adjustPerson: ['調整 {name}', 'Edit {name}'], overBadge: ['超額', 'Over capacity'],
  noFiltered: ['沒有符合篩選嘅同事。', 'No matches for the current filters.'],
  dropName: ['可拖姓名到呢度。', 'Drop a name here.'],
  noSearch: ['沒有符合搜尋嘅同事。', 'No search results.'],
  noSplits: ['呢個檢視未有與舊同桌分開嘅同事。', 'No separated tablemates in this view.'],
  foundInDrawer: ['符合搜尋嘅同事喺右側「未分配名單」。', 'Matches are in the Unassigned list on the right.'],
  name: ['姓名', 'Name'], unit: ['品牌／業務單位', 'Brand / business unit'], dept: ['部門', 'Department'],
  business: ['事業群', 'Business group'], lastTable: ['去年餐桌', 'Last year’s table'], noOldMatch: ['未能核對', 'No verified match'],
  rosterId: ['名單 ID', 'Roster ID'],
  otherDepartments: [' 等 {n} 個部門', ' ({n} departments)'],
  tableCount: ['{n} 桌', '{n} tables'], groupCount: ['{n} 組', '{n} groups'],
  openSeats: ['總剩餘 {n} 個座位', '{n} seats open'], overSeats: ['總超出 {n} 個座位', '{n} seats over capacity'],
  overCount: ['{n} 桌超額', 'Tables over capacity: {n}'], matches: ['{n} 人符合搜尋', 'Matches: {n}'],
  tableHelp: ['拖動姓名到另一張枱。點姓名可調位或交換。', 'Drag a name to another table, or click it to move or swap.'],
  groupHelp: ['拖動姓名到另一組。點姓名可調位或交換。', 'Drag a name to another group, or click it to move or swap.'],
  rosterHelp: ['點姓名可調整遊戲組或餐桌。', 'Click a name to update its game group or table.'],
  historySummary: ['去年座位已核對 {matched}/{total} 人 · 與舊同桌分坐 {tables} 人 · 不同遊戲組 {groups} 人',
    'Last year’s seats matched: {matched}/{total} · Separated at tables: {tables} · In different game groups: {groups}'],
  lastTableMissing: ['未能從去年座位表核對此人的桌號。', 'No verified table for this person in last year’s seating plan.'],
  tablemateCount: ['今年名單有 {n} 位舊同桌。點姓名可查看或調整。', 'Last year’s tablemates on this roster: {n}. Click a name to review or move.'],
  noTablemate: ['今年名單內沒有其他已核對的舊同桌。', 'No other verified tablemates on this year’s roster.'],
  moveDirect: ['直接移動', 'Move without swapping'],
  readOnly: ['Excel 目前以唯讀模式開啟。', 'The Excel workbook is open as read-only.'],
  fullTable: ['呢張枱已滿。可選一位同事交換。', 'This table is full. Select someone to swap with.'],
  excelUpdated: ['已讀取 Excel 嘅最新內容。', 'Latest Excel changes loaded.'],
  retrySelection: ['已讀取 Excel 最新內容，請重新選擇要移動嘅人。', 'Excel has changed. Select the person again to move them.'],
  couldNotRead: ['未能讀取 Excel。', 'Could not read Excel.'],
  couldNotSave: ['未能儲存。', 'Could not save.'],
  changedLive: ['已更新開啟中嘅 Excel。', 'Updated in the open workbook.'],
  changedFile: ['已儲存到 Excel。', 'Saved to Excel.'],
  undone: ['已復原。', 'Move undone.'], swapped: ['已交換。', 'Assignments swapped.'],
  overNotice: ['第 {table} 桌有 {n} 人，超出容量。', 'Table {table} has {n} people and is over capacity.'],
  editExcel: ['可在 Excel 嘅 Assignments 頁更改黃色欄位。', 'Edit the yellow cells on the Assignments sheet in Excel.'],
  saved: ['Excel 工作簿已儲存。', 'Excel workbook saved.'],
  couldNotComplete: ['未能完成操作。', 'Could not complete the request.'],
};

function preference(key, options, fallback) {
  try {
    const saved = localStorage.getItem(key);
    if (options.includes(saved)) return saved;
  } catch {}
  return fallback;
}

function savePreference(key, value) {
  try { localStorage.setItem(key, value); } catch {}
}

let language = preference('grouping-language', ['zh','en'], 'zh');
let theme = preference('grouping-theme', ['light','dark'], 'light');
let notice = null;

function t(key, tokens = {}) {
  const value = copy[key][language === 'en' ? 1 : 0];
  return value.replace(/\{(\w+)\}/g, (_, token) => tokens[token]);
}

function peopleLabel(n) {
  return t(n === 1 ? 'onePerson' : 'people', {n});
}

function applyPreferences() {
  document.documentElement.lang = language === 'en' ? 'en' : 'zh-Hant-HK';
  document.documentElement.dataset.theme = theme;
  document.title = "Group Allocation Platform";
  document.querySelectorAll('[data-i18n]').forEach(node => {node.textContent = t(node.dataset.i18n);});
  document.querySelectorAll('[data-i18n-label]').forEach(node => {node.setAttribute('aria-label', t(node.dataset.i18nLabel));});
  document.querySelectorAll('[data-i18n-placeholder]').forEach(node => {node.placeholder = t(node.dataset.i18nPlaceholder);});
  for (const key of ['zh','en']) $(`language-${key}`).setAttribute('aria-pressed', String(language === key));
  for (const key of ['light','dark']) $(`theme-${key}`).setAttribute('aria-pressed', String(theme === key));
  renderNotice();
}

function translateDiagnostic(message) {
  if (!message || language === 'zh') return message;
  if (/Automation|自動化/.test(message)) return 'In System Settings → Privacy & Security → Automation, allow the app running this desk to control Microsoft Excel.';
  const missingId = message.match(/Excel 第 (\d+) 列缺少有效的名單 ID 或姓名/);
  if (missingId) return `Row ${missingId[1]} is missing a valid roster ID or name.`;
  const duplicateId = message.match(/Excel 有重複的名單 ID (\w+)/);
  if (duplicateId) return `Duplicate roster ID: ${duplicateId[1]}.`;
  const invalidAssignment = message.match(/^(\w+)（第 (\d+) 列）的(遊戲組|餐桌)請填 1–(\d+) 或留空/);
  if (invalidAssignment) {
    const [,id,row,field,limit]=invalidAssignment;
    return `${id} (row ${row}): choose a ${field === '遊戲組' ? 'group' : 'table'} from 1–${limit}, or leave it blank.`;
  }
  if (/暫時未能連線/.test(message)) return 'Excel is temporarily unavailable. Press Enter to finish editing the cell, then try again.';
  if (/暫未開啟/.test(message)) return 'Open the linked workbook in Excel.';
  if (/唯讀/.test(message)) return t('readOnly');
  if (/結構|欄位|遊戲組編號/.test(message) && /Overview|Assignments/.test(message)) return 'Check the layout of the Assignments and Overview sheets.';
  if (/儲存這份工作簿/.test(message)) return 'Could not save this workbook. Press ⌘S directly in Excel.';
  if (/重新選擇|剛更新|交換對象已移動/.test(message)) return t('retrySelection');
  if (/重新整理/.test(message)) return 'Refresh the workbench and try again.';
  if (/暫時讀不到/.test(message)) return 'Could not read the workbook. Check the file location.';
  if (/目的地/.test(message)) return 'Invalid destination.';
  if (/找不到(這位|要交換)/.test(message)) return 'Could not find the selected person.';
  return /[\u3400-\u9fff]/.test(message) ? t('couldNotComplete') : message;
}

function elem(tag, css, text) {
  const node = document.createElement(tag);
  if (css) node.className = css;
  if (text !== undefined) node.textContent = text;
  return node;
}

function label(field, value) {
  if (value == null) return t('unassigned');
  return t(field === 'group' ? 'groupNumber' : 'tableNumber', {n:value});
}

function groupSubtitle(group) {
  if (group == null) return '';
  const name = state.groups[group - 1];
  return name === `遊戲組 ${group}` ? '' : name;
}

function note(message, error = false) {
  notice = {message,error};
  renderNotice();
}

function renderNotice() {
  const message = typeof notice?.message === 'function' ? notice.message() : notice?.message;
  $('message').textContent = message || '';
  $('message').classList.toggle('error', Boolean(notice?.error));
  $('message').hidden = !message;
}

function handleStatus() {
  $('sync').textContent = failed ? t('disconnected') : writing ? t('writingExcel')
    : !state ? t('readingExcel') : workbook?.demo ? t('demoMode') : t('cloudMode');
  $('sync-help').textContent = t(workbook?.demo ? 'demoHelp' : 'cloudHelp');
  $('undo').disabled = !undo || writing || failed || !state?.writable;
  $('disconnect').disabled = writing;
  $('open').hidden = !workbook?.url;
  if (workbook?.url) $('open').href = workbook.url;
  const submit = $('move-form').querySelector('[type=submit]');
  submit.disabled = writing || failed || !state?.writable;
  submit.textContent = t(workbook?.demo ? 'updateDemo' : 'updateExcel');
}

function peopleFor(field, value) {
  return state.people.filter((person) => person[field] === value);
}

function byName(a, b) {
  return a.name.localeCompare(b.name, ['en', 'zh-Hant-HK']) || a.id.localeCompare(b.id);
}

function match(person, term) {
  return [person.name, person.department, person.business, person.unit, person.id]
    .some((text) => text.toLocaleLowerCase().includes(term));
}

function buildHistory(people) {
  const oldTables = new Map();
  for (const person of people) {
    const oldTable = previousSeating[person.id];
    if (!oldTable) continue;
    if (!oldTables.has(oldTable)) oldTables.set(oldTable, []);
    oldTables.get(oldTable).push(person);
  }
  return new Map(people.map((person) => {
    const oldTable = previousSeating[person.id];
    const peers = (oldTables.get(oldTable) || []).filter((peer) => peer.id !== person.id);
    const split = (field) => person[field] != null
      && peers.some(peer => peer[field] != null && peer[field] !== person[field]);
    return [person.id, {oldTable, peers, table:split('table'), group:split('group')}];
  }));
}

function visiblePerson(person, field, term) {
  const info = history.get(person.id);
  const isSplit = field ? info[field] : info.table || info.group;
  return match(person, term) && (!$('split-only').checked || isSplit);
}

function personButton(person, field, duplicate, showOther = true) {
  const button = elem('button', 'person');
  const name = elem('span', 'person-name');
  name.append(elem('span','',person.name));
  const info = history.get(person.id);
  const splitLabels = [
    field !== 'group' && info.table ? t('splitTables') : '',
    field !== 'table' && info.group ? t('splitGroups') : '',
  ].filter(Boolean);
  for (const splitLabel of splitLabels) name.append(elem('span','history-badge',splitLabel));
  const sub = elem('span', 'person-dept', [...new Set([person.unit, person.department])].filter(Boolean).join(' · '));
  const text = elem('span');
  text.append(name);
  if (field) text.append(sub);
  if (duplicate) text.append(elem('span','person-dept',person.id));
  button.append(text);
  if (showOther) {
    const other = field === 'table' ? 'group' : 'table';
    button.append(elem('span','chip', label(other, person[other])));
  }
  button.setAttribute('aria-label', [t('adjustPerson',{name:person.name}), person.department, ...splitLabels].filter(Boolean).join(' '));
  button.type = 'button';
  button.onclick = () => openMove(person, field || 'table');
  button.draggable = !!field && state.writable && !failed;
  button.ondragstart = (event) => {
    if (!button.draggable || writing) return event.preventDefault();
    dragged = { id: person.id, revision: state.revision };
    event.dataTransfer.setData('text/plain', person.id);
    event.dataTransfer.effectAllowed = 'move';
    showDropDestinations(field);
  };
  button.ondragend = () => {
    dragged = null;
    $('drop-destinations').hidden = true;
    document.querySelectorAll('.drop').forEach((target) => target.classList.remove('drop'));
    render();
  };
  return button;
}

function makeDropTarget(node, field, value) {
  node.ondragover = (event) => {
    if (!dragged) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = 'move';
    node.classList.add('drop');
  };
  node.ondragleave = (event) => {
    if (!(event.relatedTarget instanceof Node) || !node.contains(event.relatedTarget)) node.classList.remove('drop');
  };
  node.ondrop = (event) => {
    if (!dragged) return;
    event.preventDefault();
    node.classList.remove('drop');
    const { id, revision } = dragged;
    dragged = null;
    $('drop-destinations').hidden = true;
    const current = state.people.find((person) => person.id === id);
    if (current && current[field] !== value) void move({person:id, field, value, revision});
    else render();
  };
}

function showDropDestinations(field) {
  const rail = $('drop-destinations');
  const limit = field === 'table' ? state.tableCount : state.groups.length;
  const destinations = [...Array.from({length:limit}, (_,i) => i+1), null];
  rail.replaceChildren(...destinations.map((value) => {
    const target = elem('span', 'drop-shortcut', value == null ? t('unassignedList') : label(field,value));
    makeDropTarget(target, field, value);
    return target;
  }));
  rail.hidden = false;
}

function card(field, value, term, duplicates) {
  const members = peopleFor(field, value);
  const visible = members.filter((person) => visiblePerson(person, field, term)).sort(byName);
  const count = members.length;
  const max = state.capacity;
  const host = field === 'table' && value === state.hostTable;
  const bucket = elem('article', `bucket${host ? ' host' : ''}${field === 'table' && count > max ? ' over' : ''}`);
  const title = elem('h3');
  title.append(elem('span', '', label(field, value) + (host ? ' · ' + t('hostTable') : '')), elem('span', 'bucket-count',
    field === 'table' ? `${count}/${max}${count > max ? ' ' + t('overBadge') : ''}` : peopleLabel(count)));
  bucket.append(title);
  const groupName = field === 'group' ? groupSubtitle(value) : '';
  if (groupName) bucket.append(elem('p', 'bucket-name', groupName));
  const departments = [...new Set(members.map(person => person.department).filter(Boolean))];
  if (departments.length) {
    const byDepartment = new Map(departments.map(department => [
      department, members.filter(person => person.department === department).length,
    ]));
    const sorted = departments.sort((a,b) => byDepartment.get(b) - byDepartment.get(a) || a.localeCompare(b));
    const separator = language === 'en' ? ', ' : '、';
    const summary = sorted.slice(0,3).join(separator) + (sorted.length > 3 ? t('otherDepartments',{n:sorted.length}) : '');
    const description = elem('p', 'bucket-name', summary);
    description.title = sorted.join(separator);
    bucket.append(description);
  }
  if (!visible.length) bucket.append(elem('p','empty', t(count ? 'noFiltered' : 'dropName')));
  else {
    const list = elem('ul');
    for (const person of visible) {
      const li = elem('li');
      li.append(personButton(person, field, duplicates.has(person.name)));
      list.append(li);
    }
    bucket.append(list);
  }
  makeDropTarget(bucket, field, value);
  return bucket;
}

function renderUnassigned(field, term, duplicates) {
  const toggle = $('unassigned-toggle');
  const drawer = $('unassigned-drawer');
  const showDrawer = !!field && drawerOpen;
  $('desk').classList.toggle('has-drawer', showDrawer);
  $('desk').classList.toggle('has-dock', !!field);
  toggle.hidden = !field;
  drawer.hidden = !showDrawer;
  toggle.setAttribute('aria-expanded', String(showDrawer));
  if (!field) return;
  const unassigned = peopleFor(field,null);
  $('unassigned-count').textContent = unassigned.length;
  toggle.setAttribute('aria-label', `${t('unassignedList')} · ${peopleLabel(unassigned.length)} · ${t(showDrawer ? 'closeDrawer' : 'openDrawer')}`);
  makeDropTarget(toggle, field, null);
  if (!showDrawer) return;

  makeDropTarget(drawer, field, null);
  $('unassigned-context').textContent = `${t(active)} · ${peopleLabel(unassigned.length)}`;
  const list = $('unassigned-people');
  const scrollPosition = list.scrollTop;
  const visible = unassigned.filter((person) => match(person,term)).sort(byName);
  list.replaceChildren(...visible.map((person) => {
    const item = elem('li');
    item.append(personButton(person, field, duplicates.has(person.name)));
    return item;
  }));
  list.scrollTop = scrollPosition;
  $('unassigned-empty').textContent = t(unassigned.length ? 'noSearch' : 'dropName');
  $('unassigned-empty').hidden = !!visible.length;
}

function roster(visible, duplicates) {
  const table = elem('table','roster');
  const thead = elem('thead');
  const heading = elem('tr');
  ['name','unit','dept','groupField','tableField'].forEach((name, i) => heading.append(elem('th',i === 1 || i === 2 ? 'dept' : '',t(name))));
  thead.append(heading);
  const body = elem('tbody');
  for (const person of visible.sort(byName)) {
    const row = elem('tr');
    const name = elem('td');
    name.append(personButton(person, null, duplicates.has(person.name), false));
    row.append(name, elem('td','dept',person.unit || person.business), elem('td','dept',person.department), elem('td','',label('group',person.group)), elem('td','',label('table',person.table)));
    body.append(row);
  }
  table.append(thead, body);
  return table;
}

function render() {
  handleStatus();
  if (!state || dragged) return;
  history = buildHistory(state.people);
  const field = views[active];
  $('desk').dataset.view = active;
  const term = $('search').value.trim().toLocaleLowerCase();
  const visible = state.people.filter((person) => visiblePerson(person, field, term));
  const labelCounts = new Map();
  for (const person of state.people) labelCounts.set(person.name, (labelCounts.get(person.name) || 0) + 1);
  const duplicates = new Set([...labelCounts].filter(([, count]) => count > 1).map(([name]) => name));
  for (const tab of Object.keys(views)) $(tab).setAttribute('aria-pressed', String(tab === active));
  $('view-title').textContent = t(active);
  const over = Array.from({length: state.tableCount}, (_, i) => peopleFor('table', i + 1).length)
    .filter((count) => count > state.capacity).length;
  const open = state.tableCount * state.capacity - state.people.filter((p) => p.table !== null).length;
  $('counts').textContent = [
    peopleLabel(state.people.length),
    field === 'table' ? t('tableCount',{n:state.tableCount}) : field === 'group' ? t('groupCount',{n:state.groups.length}) : '',
    field === 'table' ? t(open >= 0 ? 'openSeats' : 'overSeats',{n:Math.abs(open)}) : '',
    field === 'table' && over ? t('overCount',{n:over}) : '',
    term ? t('matches',{n:visible.length}) : '',
  ].filter(Boolean).join(' · ');
  renderUnassigned(field, term, duplicates);
  $('help').textContent = t(field === 'table' ? 'tableHelp' : field === 'group' ? 'groupHelp' : 'rosterHelp');
  const records = [...history.values()];
  const splitTables = records.filter((info) => info.table).length;
  const splitGroups = records.filter((info) => info.group).length;
  $('history-summary').textContent = t('historySummary',{
    matched:records.filter(info => info.oldTable).length, total:records.length, tables:splitTables, groups:splitGroups,
  });
  $('history-summary').hidden = !Object.keys(previousSeating).length;
  $('split-only').closest('label').hidden = $('history-summary').hidden;
  $('issues').replaceChildren(...state.issues.map((issue) => elem('p','issue',translateDiagnostic(issue))));
  $('issues').hidden = state.issues.length === 0;
  const board = $('board');
  board.setAttribute('aria-busy','false');
  if (!field) return board.replaceChildren(visible.length ? roster(visible,duplicates) : elem('p','empty',t('noFiltered')));
  const grid = elem('div','grid');
  const count = field === 'table' ? state.tableCount : state.groups.length;
  const allCards = Array.from({length:count}, (_,index) => index+1);
  const visibleCards = term || $('split-only').checked
    ? allCards.filter((number) => visible.some(person => person[field] === number))
    : allCards;
  if (!visibleCards.length) {
    const unassignedMatches = term && peopleFor(field,null).some(person => match(person,term));
    const message = t(unassignedMatches ? 'foundInDrawer' : term ? 'noSearch' : 'noSplits');
    return board.replaceChildren(elem('p','empty',message));
  }
  for (const number of visibleCards) grid.append(card(field, number, term, duplicates));
  board.replaceChildren(grid);
}

async function refresh() {
  if (!workbook || loading || writing) return;
  loading = true;
  const currentBinding = binding;
  try {
    const body = {state:await workbook.read()};
    if (currentBinding !== binding || writing) return;
    previousSeating = body.state.previousSeating;
    const changed = state && body.state.revision !== state.revision;
    const statusChanged = state && ['mode','excelSaved','writable','syncError'].some((key) => state[key] !== body.state[key]);
    if (failed || !state || changed || statusChanged) {
      if (failed || changed) {
        undo = null;
        note(state ? () => t('excelUpdated') : '');
      }
      if (selected && changed) {
        $('move-dialog').close();
        note(() => t('retrySelection'));
      }
      failed = false;
      state = body.state;
      render();
    }
  } catch (error) {
    if (currentBinding !== binding || writing) return;
    failed = true;
    note(() => translateDiagnostic(error.message) || t('couldNotRead'),true);
    render();
  } finally {
    loading = false;
  }
}

async function move(change, isUndo = false) {
  if (!state || !workbook || writing || failed) return;
  writing = true;
  handleStatus();
  try {
    const data = await workbook.move({...change, revision:change.revision || state.revision});
    previousSeating = data.state.previousSeating;
    state = data.state;
    undo = isUndo ? null : {...data.undo, revision:state.revision};
    const person = state.people.find((item) => item.id === change.person);
    const count = peopleFor(change.field, change.value).length;
    const overflow = change.field === 'table' && change.value && count > state.capacity;
    const ending = workbook.demo ? 'demoSaved' : 'changedFile';
    note(() => [
      isUndo ? t('undone') : change.swap ? t('swapped')
        : `${person.name} → ${label(change.field,change.value)}${language === 'en' ? '.' : '。'}`,
      t(ending),
      overflow && t('overNotice',{table:change.value,n:count}),
    ].filter(Boolean).join(' '), overflow);
    if ($('move-dialog').open) $('move-dialog').close();
  } catch (error) {
    await readFailedMove(error);
  } finally {
    writing = false;
    render();
  }
}

async function readFailedMove(error) {
  writing = false;
  await refresh();
  note(() => translateDiagnostic(error.message) || t('couldNotSave'),true);
}

function openMove(person, field) {
  selected = person;
  $('person-name').textContent = person.name;
  const oldTable = history.get(person.id).oldTable;
  const details = [
    ['unit',person.unit], ['dept',person.department], ['business',person.business],
    ['groupField',[label('group',person.group),groupSubtitle(person.group)].filter(Boolean).join(' · ')],
    ['tableField',label('table',person.table) + (person.table === state.hostTable ? ' · ' + t('hostTable') : '')],
    ['lastTable',oldTable ? label('table',oldTable) : t('noOldMatch')],
    ['rosterId',person.id],
  ];
  $('person-detail').replaceChildren(...details.flatMap(([title,value]) => [elem('dt','',t(title)),elem('dd','',value || '—')]));
  showFormerTablemates(person, field);
  $('field').value = field;
  destinations();
  if (!$('move-dialog').open) $('move-dialog').showModal();
  handleStatus();
}

function showFormerTablemates(person, field) {
  const info = history.get(person.id);
  const list = $('former-tablemates');
  const isSplit = (peer, key) => person[key] != null && peer[key] != null && person[key] !== peer[key];
  const peers = [...info.peers].sort((a,b) =>
    Number(isSplit(b,field)) - Number(isSplit(a,field)) || byName(a,b));
  $('history-note').textContent = t(
    !info.oldTable ? 'lastTableMissing' : peers.length ? 'tablemateCount' : 'noTablemate', {n:peers.length}
  );
  list.replaceChildren(...peers.map((peer) => {
    const row = elem('li');
    const peerButton = elem('button','history-person',peer.name);
    peerButton.type = 'button';
    peerButton.onclick = () => {
      openMove(peer,$('field').value);
      $('person-name').focus();
    };
    row.append(peerButton, elem('span','history-location',`${label('table',peer.table)} · ${label('group',peer.group)}`));
    if (isSplit(peer,'table')) row.append(elem('span','history-badge',t('splitTables')));
    if (isSplit(peer,'group')) row.append(elem('span','history-badge',t('splitGroups')));
    return row;
  }));
}

function destinations() {
  const field = $('field').value;
  const limit = field === 'group' ? state.groups.length : state.tableCount;
  const option = (value) => {
    const n = peopleFor(field,value).length;
    const name = field === 'group' ? groupSubtitle(value) : '';
    const suffix = name ? ` · ${name}`
      : field === 'table' && value === state.hostTable ? ' · ' + t('hostTable') : '';
    return new Option(`${label(field,value)}${suffix} (${peopleLabel(n)})`,value ?? '');
  };
  $('destination').replaceChildren(...Array.from({length:limit}, (_,i) => option(i+1)), option(null));
  $('destination').value = selected[field] ?? '';
  swapOptions();
}

function swapOptions() {
  const field = $('field').value;
  const target = Number($('destination').value) || null;
  const swap = $('swap');
  const candidates = target !== selected[field] ? peopleFor(field,target).sort(byName) : [];
  swap.replaceChildren(new Option(t('moveDirect'),''), ...candidates.map((person) =>
    new Option(`${person.name} · ${person.department} [${person.id}]`,person.id)));
  $('swap-label').hidden = !candidates.length;
  $('move-note').textContent = translateDiagnostic(state.syncError) || (!state.writable ? t('readOnly')
    : field === 'table' && target && candidates.length >= state.capacity ? t('fullTable') : '');
}

for (const tab of Object.keys(views)) $(tab).onclick = () => { active = tab; render(); };
for (const key of ['zh','en']) $(`language-${key}`).onclick = () => {
  language = key;
  savePreference('grouping-language',key);
  applyPreferences();
  render();
};
for (const key of ['light','dark']) $(`theme-${key}`).onclick = () => {
  theme = key;
  savePreference('grouping-theme',key);
  applyPreferences();
};
$('search').oninput = render;
$('split-only').onchange = render;
$('unassigned-toggle').onclick = () => { drawerOpen = !drawerOpen; render(); };
$('unassigned-close').onclick = () => {
  drawerOpen = false;
  render();
  $('unassigned-toggle').focus();
};
$('undo').onclick = () => { if (undo) void move(undo,true); };
$('cancel').onclick = () => $('move-dialog').close();
$('move-dialog').onclose = () => { selected = null; };
$('field').onchange = destinations;
$('destination').onchange = swapOptions;
$('move-form').onsubmit = (event) => {
  event.preventDefault();
  const field = $('field').value;
  const value = Number($('destination').value) || null;
  if (selected[field] === value) return $('move-dialog').close();
  void move({person:selected.id,field,value,swap:$('swap').value || null});
};
function showConnectionStatus(message) {
  const status = $('connection-status');
  status.textContent = message;
  status.hidden = !message;
}

async function enterWorkbook(connection) {
  const firstState = await connection.read();
  binding += 1;
  workbook = connection;
  state = firstState;
  previousSeating = firstState.previousSeating;
  failed = false;
  $('workbench').hidden = false;
  $('excel-controls').hidden = false;
  $('connection').hidden = true;
  render();
}

$('redirect-url').textContent = new URL('./auth.html', location.href).href;
$('connection-form').addEventListener('invalid', event => {
  if (event.target.closest('details')) $('connection-settings').open = true;
}, true);
$('connection-form').onsubmit = async event => {
  event.preventDefault();
  $('connect').disabled = $('demo').disabled = true;
  showConnectionStatus(t('connecting'));
  try {
    const connection = await CloudWorkbook.connect({
      link: $('workbook-link').value.trim(),
      clientId: $('client-id').value.trim(),
      tenantId: $('tenant-id').value.trim(),
    });
    await enterWorkbook(connection);
    $('workbook-link').value = '';
    showConnectionStatus('');
  } catch (error) {
    showConnectionStatus(error.message);
  } finally {
    $('connect').disabled = $('demo').disabled = false;
  }
};
$('demo').onclick = async () => enterWorkbook(CloudWorkbook.demo());
$('disconnect').onclick = async () => {
  const previous = workbook;
  binding += 1;
  workbook = state = undo = dragged = selected = null;
  previousSeating = {};
  drawerOpen = false;
  $('workbench').hidden = $('excel-controls').hidden = $('unassigned-toggle').hidden = $('unassigned-drawer').hidden = true;
  $('connection').hidden = false;
  $('open').removeAttribute('href');
  $('board').replaceChildren();
  if ($('move-dialog').open) $('move-dialog').close();
  $('person-detail').replaceChildren();
  $('former-tablemates').replaceChildren();
  $('unassigned-people').replaceChildren();
  note('');
  try { await previous.close?.(); }
  catch (error) { showConnectionStatus(error.message); }
};
window.addEventListener('focus', refresh);
document.addEventListener('visibilitychange', () => { if (!document.hidden) void refresh(); });
applyPreferences();
setInterval(() => { if (!document.hidden && !workbook?.demo) void refresh(); }, 12000);
