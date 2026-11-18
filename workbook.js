'use strict';

const AllocationWorkbook = (() => {
  const columns = {group: 'F', table: 'G'};
  const assignmentRow = 7;
  const tableCount = 13;
  const capacity = 10;
  const hostTable = 1;

  function bilingual(zh, en) {
    return new Error(document.documentElement.lang === 'en' ? en : zh);
  }

  function assignment(value, limit) {
    const text = String(value ?? '').trim();
    if (!text) return null;
    if (/^[1-9]\d*$/.test(text) && Number(text) <= limit) return Number(text);
    throw new Error('Invalid assignment');
  }

  function parseWorkbook(grid, labels, previous, revision) {
    const groups = [];
    for (const [number, name] of labels) {
      if (Number(number) !== groups.length + 1) break;
      groups.push(String(name || '遊戲組 ' + number));
    }
    if (!groups.length) throw bilingual('Overview 搵唔到遊戲組。', 'No game groups found on Overview.');

    const people = [];
    const rows = new Map();
    const issues = [];
    grid.forEach(([id, name, department, business, unit, group, table], index) => {
      if ([id, name, department, business, unit, group, table].every(value => value == null || value === '')) return;
      const row = assignmentRow + index;
      if (!/^R\d{3,}$/.test(id ?? '') || !name || rows.has(id)) {
        throw bilingual('Assignments 第 ' + row + ' 列請核對名單 ID 同姓名。', 'Check the ID and name in Assignments row ' + row + '.');
      }
      const person = {id, name:String(name), department:String(department || ''), business:String(business || ''), unit:String(unit || '')};
      for (const [field, value, limit] of [['group', group, groups.length], ['table', table, tableCount]]) {
        try { person[field] = assignment(value, limit); }
        catch {
          person[field] = null;
          issues.push(id + '（第 ' + row + ' 列）的' + (field === 'group' ? '遊戲組' : '餐桌') + '請填 1–' + limit + ' 或留空。');
        }
      }
      people.push(person);
      rows.set(id, row);
    });
    if (!people.length) throw bilingual('Excel 入面搵唔到分組名單。', 'No assignments found in the workbook.');
    const previousSeating = Object.fromEntries(previous
      .filter(([id, table]) => rows.has(id) && /^[1-9]\d*$/.test(String(table)))
      .map(([id, table]) => [id, Number(table)]));
    return {rows, state:{
      people, groups, issues, previousSeating, tableCount, capacity, hostTable, writable:true,
      revision,
    }};
  }

  function demo() {
    const grid = [
      ['R001','Alex Demo','Management','Operations','Demo Brand A',1,1],
      ['R002','Jamie Demo','Management','Operations','Demo Brand A',1,1],
      ['R003','Casey Demo','Design','Creative','Demo Brand B',2,2],
      ['R004','Jordan Demo','Design','Creative','Demo Brand B',2,3],
      ['R005','Morgan Demo','Retail','Sales','Demo Brand C',3,4],
      ['R006','Taylor Demo','Retail','Sales','Demo Brand C','',''],
      ['R007','Sam Demo','Events','Operations','Demo Brand A',4,''],
      ['R008','Riley Demo','Events','Operations','Demo Brand A','',5],
    ];
    const labels = Array.from({length:12}, (_,i) => [i + 1, '']);
    const previous = [['R001',1],['R002',1],['R003',2],['R004',2],['R005',4],['R006',4]];
    let record = {grid, labels, previous, revision:0};
    return {
      demo:true,
      read:async () => readState(record),
      async move(change) {
        const result = applyMove(record, change);
        record = result.record;
        return {state:readState(record), undo:result.undo};
      },
    };
  }

  const spreadsheetNS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const relationNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  function readState({grid, labels, previous, revision}) {
    return parseWorkbook(grid, labels, previous, revision).state;
  }

  function applyMove(record, change) {
    const {state, rows} = parseWorkbook(record.grid, record.labels, record.previous, record.revision);
    const {person: id, swap, field, value, revision} = change;
    if (revision !== record.revision) {
      throw bilingual('另一個分頁已更新安排，請重新選擇。', 'Another tab changed the board. Select the person again.');
    }
    const limit = field === 'table' ? tableCount : state.groups.length;
    const person = state.people.find(item => item.id === id);
    const partner = swap ? state.people.find(item => item.id === swap) : null;
    if (!Object.hasOwn(columns, field) || !person || (value !== null && (!Number.isInteger(value) || value < 1 || value > limit)) ||
        (swap && (!partner || swap === id || partner[field] !== value))) {
      throw bilingual('請核對目的地或交換對象。', 'Check the destination or swap partner.');
    }
    const grid = record.grid.map(row => row.slice());
    const index = field === 'table' ? 6 : 5;
    grid[rows.get(id) - assignmentRow][index] = value ?? '';
    if (partner) grid[rows.get(swap) - assignmentRow][index] = person[field] ?? '';
    return {
      record:{...record, grid, revision:record.revision + 1},
      undo:{person:id, field, value:person[field], swap:swap || null},
    };
  }

  function database() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open('group-allocation-platform', 1);
      request.onupgradeneeded = () => request.result.createObjectStore('workbench');
      request.onsuccess = () => {
        request.result.onversionchange = () => request.result.close();
        resolve(request.result);
      };
      request.onerror = () => reject(request.error);
      request.onblocked = () => reject(bilingual('請關閉其他工作台分頁再試。', 'Close other workbench tabs and try again.'));
    });
  }

  async function access(update) {
    const db = await database();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction('workbench', update ? 'readwrite' : 'readonly');
      const store = transaction.objectStore('workbench');
      const request = store.get('current');
      let result, failure;
      request.onsuccess = () => {
        try {
          result = update ? update(request.result) : request.result;
          if (update) store.put(result.record, 'current');
        } catch (error) {
          failure = error;
          transaction.abort();
        }
      };
      transaction.oncomplete = () => { db.close(); resolve(result); };
      transaction.onabort = () => { db.close(); reject(failure ?? transaction.error); };
    });
  }

  function nameOf(path, relationships) {
    if (!path) throw new Error('Missing worksheet');
    const target = relationships.get(path);
    if (!target) throw new Error('Missing worksheet relationship');
    const url = new URL(target, 'https://workbook.invalid/xl/workbook.xml');
    if (url.hostname !== 'workbook.invalid' || !url.pathname.startsWith('/xl/')) throw new Error('Invalid worksheet path');
    return url.pathname.slice(1);
  }

  const cells = (doc, tag) => [...doc.getElementsByTagNameNS(spreadsheetNS, tag)];
  const child = (node, tag) => [...node.children].find(item => item.localName === tag);

  async function xml(zip, path) {
    const entry = zip.file(path);
    if (!entry) throw new Error('Missing workbook XML');
    const doc = new DOMParser().parseFromString(await entry.async('string'), 'application/xml');
    if (doc.querySelector('parsererror')) throw new Error('Invalid workbook XML');
    return doc;
  }

  async function unpack(buffer) {
    const zip = await JSZip.loadAsync(buffer);
    const workbook = await xml(zip, 'xl/workbook.xml');
    const relDoc = await xml(zip, 'xl/_rels/workbook.xml.rels');
    const relations = new Map([...relDoc.getElementsByTagNameNS('*','Relationship')]
      .map(item => [item.getAttribute('Id'),item.getAttribute('Target')]));
    const sheets = new Map(cells(workbook,'sheet').map(sheet => [sheet.getAttribute('name'), sheet.getAttributeNS(relationNS,'id')]));
    const paths = {
      assignments:nameOf(sheets.get('Assignments'),relations),
      overview:nameOf(sheets.get('Overview'),relations),
      previous:sheets.has('Previous Seating') ? nameOf(sheets.get('Previous Seating'),relations) : null,
    };
    const assignments = await xml(zip,paths.assignments);
    const overview = await xml(zip,paths.overview);
    const history = paths.previous ? await xml(zip,paths.previous) : null;
    const strings = zip.file('xl/sharedStrings.xml')
      ? cells(await xml(zip,'xl/sharedStrings.xml'),'si').map(item => cells(item,'t').map(t=>t.textContent).join('')) : [];
    return {zip, workbook, paths, assignments, overview, history, strings};
  }

  function spreadsheetValues(doc, strings) {
    return new Map(cells(doc,'c').map(cell => {
      const raw = child(cell,'v')?.textContent ?? '';
      const type = cell.getAttribute('t');
      const value = type === 's' ? strings[Number(raw)] ?? ''
        : type === 'inlineStr' ? cells(cell,'t').map(t=>t.textContent).join('')
          : raw;
      return [cell.getAttribute('r'),value];
    }));
  }

  function gridFrom(values, cols, first, last) {
    return Array.from({length:last - first + 1}, (_,i) => cols.map(col=>values.get(col + (first+i)) ?? ''));
  }

  async function importExcel(file) {
    if (!/\.xlsx$/i.test(file.name) || file.size > 8 * 1024 * 1024) {
      throw bilingual('請選擇 8 MB 以下嘅 .xlsx 分組工作簿。', 'Choose a grouping .xlsx workbook under 8 MB.');
    }
    const source = await file.arrayBuffer();
    let book;
    try { book = await unpack(source); }
    catch {
      throw bilingual('未能開啟這份 Excel。請使用有 Assignments 同 Overview 工作表嘅分組檔。',
        'Could not open this Excel file. Use the grouping workbook with Assignments and Overview sheets.');
    }
    const values = spreadsheetValues(book.assignments,book.strings);
    const end = Math.max(assignmentRow,...[...values.keys()]
      .filter(cell => /^A\d+$/.test(cell) && /^R\d{3,}$/.test(values.get(cell)))
      .map(cell => Number(cell.slice(1))));
    if (end > 1500) throw bilingual('分組名單超出可匯入範圍。', 'The workbook exceeds the supported roster range.');
    const grid = gridFrom(values,['A','B','C','D','E','F','G'],assignmentRow,end);
    const labels = gridFrom(spreadsheetValues(book.overview,book.strings),['A','B'],7,18);
    const previous = book.history ? gridFrom(spreadsheetValues(book.history,book.strings),['A','B'],2,end) : [];
    const record = {source, grid, labels, previous, revision:0, filename:file.name};
    readState(record);
    await access(existing => {
      if (existing) throw bilingual('此瀏覽器已有分組主檔。請重新開啟工作台。', 'This browser already has a saved board. Reopen the workbench.');
      return {record};
    });
    return persisted();
  }

  function numberCell(doc, indexedCells, address, value) {
    let cell = indexedCells.get(address);
    if (!cell) {
      if (value == null) return;
      const rowNumber = address.match(/\d+$/)[0];
      const row = cells(doc,'row').find(item => item.getAttribute('r') === rowNumber);
      if (!row) throw new Error('Missing assignment row');
      const later = [...row.children].find(item => (item.getAttribute('r') ?? '').replace(/\d+$/,'') > address.replace(/\d+$/,''));
      cell = doc.createElementNS(spreadsheetNS,'c');
      cell.setAttribute('r',address);
      row.insertBefore(cell,later ?? null);
      indexedCells.set(address,cell);
    }
    cell.removeAttribute('t');
    child(cell,'is')?.remove();
    let number = child(cell,'v');
    if (value == null) { number?.remove(); return; }
    if (!number) {
      number = doc.createElementNS(spreadsheetNS,'v');
      cell.append(number);
    }
    number.textContent = String(value);
  }

  function cellIndex(doc) {
    return new Map(cells(doc,'c').map(cell=>[cell.getAttribute('r'),cell]));
  }

  function setNote(doc, index, address, text) {
    const cell = index.get(address);
    if (!cell) return;
    cell.replaceChildren();
    cell.setAttribute('t','inlineStr');
    const value = doc.createElementNS(spreadsheetNS,'is');
    const run = doc.createElementNS(spreadsheetNS,'t');
    run.textContent = text;
    value.append(run);
    cell.append(value);
  }

  async function exportExcel(record) {
    const {zip, workbook, assignments, overview, paths} = await unpack(record.source);
    const {state, rows} = parseWorkbook(record.grid,record.labels,record.previous,record.revision);
    const assignmentCells = cellIndex(assignments);
    setNote(assignments,assignmentCells,'A2','此檔為工作台匯出。請在工作台更新安排。 / Workbench export. Update assignments in the workbench.');
    setNote(assignments,assignmentCells,'A3','第 1 桌為主家席。留空代表未分配。 / Table 1 is the head table. Blank means unassigned.');

    for (const person of state.people) {
      const row = rows.get(person.id);
      for (const field of ['group','table']) {
        const address = columns[field]+row;
        child(assignmentCells.get(address) ?? {children:[]},'f')?.remove();
        numberCell(assignments,assignmentCells,address,person[field]);
      }
    }
    const groupCounts = new Map(), tableCounts = new Map();
    for (const {group,table} of state.people) {
      groupCounts.set(group,(groupCounts.get(group) ?? 0)+1);
      tableCounts.set(table,(tableCounts.get(table) ?? 0)+1);
    }
    const overviewCells = cellIndex(overview);
    const cache = (address, value) => {
      const cell = overviewCells.get(address);
      if (cell && child(cell,'f')) numberCell(overview,overviewCells,address,value);
    };
    state.groups.forEach((_,i)=>cache('C'+(7+i),groupCounts.get(i+1) ?? 0));
    cache('C19',groupCounts.get(null) ?? 0);
    for (let i=1;i<=tableCount;i++) {
      const count = tableCounts.get(i) ?? 0;
      cache('F'+(6+i),count);
      cache('G'+(6+i),capacity-count);
    }
    cache('F20',tableCounts.get(null) ?? 0);
    cells(workbook,'calcPr')[0]?.setAttribute('fullCalcOnLoad','1');
    const serializer = new XMLSerializer();
    zip.file(paths.assignments,serializer.serializeToString(assignments));
    zip.file(paths.overview,serializer.serializeToString(overview));
    zip.file('xl/workbook.xml',serializer.serializeToString(workbook));
    const data = await zip.generateAsync({type:'blob',compression:'DEFLATE',mimeType:mime});
    return {data, filename:record.filename.replace(/(?:-Export)?\.xlsx$/i,'-Export.xlsx'), state};
  }

  function persisted() {
    return {
      read:async () => readState(await access()),
      async move(change) {
        const result = await access(record=>applyMove(record,change));
        return {state:readState(result.record), undo:result.undo};
      },
      async save() {
        const saved = await access(record => ({record}));
        return readState(saved.record);
      },
      export:async () => exportExcel(await access()),
    };
  }

  return {open:async () => (await access()) ? persisted() : null, importExcel, demo};
})();
