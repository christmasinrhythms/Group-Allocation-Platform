const $ = (id) => document.getElementById(id);
const views = { tables: 'table', groups: 'group', roster: null };
let previousSeating = {};
let state, workbook, binding = 0, dragged, undo, selected, loading = false, writing = false, exporting = false, failed = false;
let history = new Map();
let active = 'tables';
let drawerOpen = false;

const copy = {
  connectTitle: ['匯入分組 Excel', 'Import grouping Excel'],
  connectIntro: ['只需匯入一次。之後喺工作台改安排，會自動儲存在呢個瀏覽器。', 'Import once. The workbench saves future changes in this browser.'],
  importExcel: ['選擇 Excel 檔案', 'Choose Excel file'],
  importing: ['正在匯入分組名單…', 'Importing roster…'],
  saveBoard: ['儲存安排', 'Save arrangements'],
  exportExcel: ['匯出 Excel', 'Export Excel'],
  exporting: ['正在產生 Excel…', 'Preparing Excel…'],
  exported: ['已產生 Excel 檔案。', 'Excel file generated.'],
  exportError: ['未能匯出 Excel。', 'Could not export Excel.'],
  exitDemo: ['離開示範', 'Exit demo'],
  saveMove: ['儲存更改', 'Save change'],
  demo: ['試用示範名單', 'Try demo roster'],
  demoMode: ['示範模式', 'Demo mode'],
  demoHelp: ['示範只使用虛構名單。試做嘅改動會在離開示範時清除。', 'The demo uses fictional names. Trial changes clear when you leave the demo.'],
  demoSaved: ['已更新示範名單。', 'Demo roster updated.'],
  updateDemo: ['更新示範名單', 'Update demo roster'],
  stored: ['已儲存在此瀏覽器', 'Saved in this browser'],
  saving: ['正在儲存…', 'Saving…'],
  storeUnavailable: ['瀏覽器暫時未能讀取已存安排。', 'The browser could not read the saved arrangements.'],
  storageHelp: ['安排會自動儲存在呢個瀏覽器。匯出 Excel 可另存一份。', 'Changes save automatically in this browser. Export Excel to keep a copy.'],
  language: ['語言', 'Language'], theme: ['顯示主題', 'Appearance'],
  light: ['淺色', 'Light'], dark: ['深色', 'Dark'],
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
  tableNumber: ['第 {n} 桌', 'Table {n}'], groupNumber: ['第 {n} 組', 'Group {n}'],
  people: ['{n} 人', '{n} people'], onePerson: ['1 人', '1 person'], hostTable: ['主家席', 'Head table'],
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
  readOnly: ['工作台目前只可查看。', 'The board is read-only.'],
  fullTable: ['呢張枱已滿。可選一位同事交換。', 'This table is full. Select someone to swap with.'],
  excelUpdated: ['已載入最新安排。', 'Latest arrangements loaded.'],
  retrySelection: ['安排已變更，請重新選擇。', 'The board changed. Select the person again.'],
  couldNotRead: ['未能讀取安排。', 'Could not read the arrangements.'],
  couldNotSave: ['未能儲存。', 'Could not save.'],
  changedFile: ['已儲存到工作台。', 'Saved to the workbench.'],
  undone: ['已復原。', 'Move undone.'], swapped: ['已交換。', 'Assignments swapped.'],
  overNotice: ['第 {table} 桌有 {n} 人，超出容量。', 'Table {table} has {n} people and is over capacity.'],
  saved: ['分組安排已儲存。', 'Arrangements saved.'],
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
  const missingId = message.match(/Excel 第 (\d+) 列缺少有效的名單 ID 或姓名/);
  if (missingId) return `Row ${missingId[1]} is missing a valid roster ID or name.`;
  const duplicateId = message.match(/Excel 有重複的名單 ID (\w+)/);
  if (duplicateId) return `Duplicate roster ID: ${duplicateId[1]}.`;
  const invalidAssignment = message.match(/^(\w+)（第 (\d+) 列）的(遊戲組|餐桌)請填 1–(\d+) 或留空/);
  if (invalidAssignment) {
    const [,id,row,field,limit]=invalidAssignment;
    return `${id} (row ${row}): choose a ${field === '遊戲組' ? 'group' : 'table'} from 1–${limit}, or leave it blank.`;
  }
  if (/目的地/.test(message)) return 'Invalid destination.';
  if (/找不到(這位|要交換)/.test(message)) return 'Could not find the selected person.';
  return message;
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
  $('save-status').textContent = failed ? t('storeUnavailable') : writing ? t('saving')
    : workbook?.demo ? t('demoMode') : t('stored');
  $('storage-help').textContent = t(workbook?.demo ? 'demoHelp' : 'storageHelp');
  $('undo').disabled = !undo || writing || exporting || failed || !state?.writable;
  $('save').hidden = $('export').hidden = Boolean(workbook?.demo);
  $('exit-demo').hidden = !workbook?.demo;
  $('save').disabled = $('export').disabled = writing || exporting || failed;
  const submit = $('move-form').querySelector('[type=submit]');
  submit.disabled = writing || exporting || failed || !state?.writable;
  submit.textContent = t(workbook?.demo ? 'updateDemo' : 'saveMove');
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
  if (!workbook || loading || writing || exporting) return;
  loading = true;
  const currentBinding = binding;
  try {
    const body = {state:await workbook.read()};
    if (currentBinding !== binding || writing) return;
    previousSeating = body.state.previousSeating;
    const changed = state && body.state.revision !== state.revision;
    if (failed || !state || changed) {
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
  if (!state || !workbook || writing || exporting || failed) return;
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
  $('move-note').textContent = (!state.writable ? t('readOnly')
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

function enableImport(enabled) {
  $('import').disabled = $('demo').disabled = !enabled;
}

$('import').onclick = () => $('workbook-file').click();
$('workbook-file').onchange = async () => {
  const file = $('workbook-file').files[0];
  if (!file) return;
  enableImport(false);
  showConnectionStatus(t('importing'));
  try {
    await enterWorkbook(await AllocationWorkbook.importExcel(file));
    showConnectionStatus('');
  } catch (error) {
    showConnectionStatus(error.message);
  } finally {
    $('workbook-file').value = '';
    enableImport(true);
  }
};
$('demo').onclick = async () => enterWorkbook(AllocationWorkbook.demo());
$('exit-demo').onclick = () => {
  binding += 1;
  workbook = state = undo = dragged = selected = null;
  previousSeating = {};
  drawerOpen = false;
  $('workbench').hidden = $('excel-controls').hidden = $('unassigned-toggle').hidden = $('unassigned-drawer').hidden = true;
  $('connection').hidden = false;
  $('board').replaceChildren();
  if ($('move-dialog').open) $('move-dialog').close();
  note('');
};
$('save').onclick = async () => {
  writing = true;
  handleStatus();
  try {
    const saved = await workbook.save();
    if (saved.revision !== state.revision) undo = null;
    state = saved;
    previousSeating = saved.previousSeating;
    note(() => t('saved'));
  } catch (error) {
    note(() => error.message || t('couldNotSave'),true);
  } finally {
    writing = false;
    render();
  }
};
$('export').onclick = async () => {
  exporting = true;
  handleStatus();
  note(() => t('exporting'));
  try {
    const {data, filename} = await workbook.export();
    const url = URL.createObjectURL(data);
    const link = elem('a');
    link.href = url;
    link.download = filename;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    note(() => t('exported'));
  } catch (error) {
    note(() => (error.message ? t('exportError') + ' ' + error.message : t('exportError')),true);
  } finally {
    exporting = false;
    render();
  }
};
window.addEventListener('focus', refresh);
document.addEventListener('visibilitychange', () => { if (!document.hidden) void refresh(); });
applyPreferences();
(async () => {
  enableImport(false);
  try {
    const saved = await AllocationWorkbook.open();
    if (saved) await enterWorkbook(saved);
  } catch (error) {
    showConnectionStatus(t('storeUnavailable') + ' ' + error.message);
  } finally {
    enableImport(true);
  }
})();
