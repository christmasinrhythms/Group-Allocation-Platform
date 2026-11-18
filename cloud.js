'use strict';

const CloudWorkbook = (() => {
  const scopes = ['Files.ReadWrite'];
  const graph = 'https://graph.microsoft.com/v1.0';
  const guid = /^[a-f\d]{8}(?:-[a-f\d]{4}){3}-[a-f\d]{12}$/i;
  const columns = {group: 'F', table: 'G'};
  const assignmentRow = 7;
  const assignmentEnd = 134;
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

  function parseWorkbook(grid, labels, previous) {
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
    const previousSeating = Object.fromEntries(previous
      .filter(([id, table]) => rows.has(id) && /^[1-9]\d*$/.test(String(table)))
      .map(([id, table]) => [id, Number(table)]));
    return {rows, state:{
      people, groups, issues, previousSeating, tableCount, capacity, hostTable, writable:true,
      revision:JSON.stringify([grid, labels, previous]),
    }};
  }

  function planMove({state, rows}, change) {
    if (change.revision !== state.revision) {
      throw bilingual('Excel 有新更改，請重新選擇。', 'Excel has changed. Select the person again to move them.');
    }
    const {person: id, swap, field, value} = change;
    const column = columns[field];
    const limit = field === 'table' ? tableCount : state.groups.length;
    const person = state.people.find(item => item.id === id);
    const partner = swap ? state.people.find(item => item.id === swap) : null;
    if (!Object.hasOwn(columns, field) || !person || (value !== null && (!Number.isInteger(value) || value < 1 || value > limit)) ||
        (swap && (!partner || swap === id || partner[field] !== value))) {
      throw bilingual('請核對目的地或交換對象。', 'Check the destination or swap partner.');
    }

    const row = rows.get(id);
    const partnerRow = partner ? rows.get(swap) : row;
    const first = Math.min(row, partnerRow);
    const last = Math.max(row, partnerRow);
    // Graph interprets null as “leave this cell alone”. An empty string clears an assignment.
    const values = Array.from({length:last - first + 1}, () => [null]);
    values[row - first][0] = value ?? '';
    if (partner) values[partnerRow - first][0] = person[field] ?? '';
    return {field, first, values, address:column + first + ':' + column + last,
      undo:{person:id, field, value:person[field], swap:swap || null}};
  }

  function readLink(value) {
    let url;
    try { url = new URL(value); } catch {}
    if (!url || url.protocol !== 'https:' || !url.hostname.endsWith('.sharepoint.com') || url.username || url.password || url.port) {
      throw bilingual('請貼上公司 SharePoint 或 OneDrive 嘅 HTTPS Excel 連結。', 'Paste an HTTPS workbook link from your work SharePoint or OneDrive.');
    }
    return url;
  }

  function lookupPath(value) {
    let url;
    try { url = new URL(value); } catch {}
    if (url?.origin === 'https://graph.microsoft.com' && !url.username && !url.password && !url.search && !url.hash) {
      const item = url.pathname.match(/^\/v1\.0(\/drives\/[A-Za-z0-9!_-]+\/items\/[A-Za-z0-9!_-]+)$/);
      if (item) return {path:item[1], shareUrl:null};
    }
    const shareUrl = readLink(value);
    const bytes = new TextEncoder().encode(shareUrl.href);
    const encoded = 'u!' + btoa(String.fromCharCode(...bytes)).replaceAll('/', '_').replaceAll('+', '-').replace(/=+$/, '');
    return {path:'/shares/' + encoded + '/driveItem', shareUrl};
  }

  async function connect({link, clientId, tenantId}) {
    const lookup = lookupPath(link);
    if (!guid.test(clientId) || !guid.test(tenantId)) {
      throw bilingual('請填上有效嘅 Microsoft Entra 租戶 ID 同應用程式 ID。', 'Enter your Microsoft Entra tenant and application IDs.');
    }
    const auth = new msal.PublicClientApplication({
      auth:{clientId, authority:'https://login.microsoftonline.com/' + tenantId,
        redirectUri:new URL('./auth.html', location.href).href, supportedBroker:false},
      cache:{cacheLocation:'memoryStorage'},
    });
    let account;
    try {
      await auth.initialize();
      account = (await auth.loginPopup({scopes, prompt:'select_account'})).account;
    } catch {
      throw bilingual('Microsoft 登入未完成。請核對 Entra 設定、彈出視窗同公司核准。', 'Microsoft sign-in did not complete. Check Entra setup, pop-ups, and your organization’s approval.');
    }
    let retryAfter = 0;
    async function request(path, patch) {
      if (Date.now() < retryAfter) {
        throw bilingual('Microsoft 正限制請求，稍後會再同步。', 'Microsoft is throttling requests. Sync will resume shortly.');
      }
      let accessToken;
      try { ({accessToken} = await auth.acquireTokenSilent({account, scopes})); }
      catch { throw bilingual('登入已失效。請離開工作簿再登入。', 'Your sign-in expired. Disconnect and sign in again.'); }
      const headers = {Authorization:'Bearer ' + accessToken};
      if (patch) headers['Content-Type'] = 'application/json';
      const response = await fetch(graph + path, {
        headers, method:patch ? 'PATCH' : 'GET', body:patch && JSON.stringify(patch), cache:'no-store',
        referrerPolicy:'no-referrer', redirect:'error', signal:AbortSignal.timeout(15000),
      });
      if (response.status === 429 || response.status === 503) {
        const seconds = Number(response.headers.get('Retry-After')) || 15;
        retryAfter = Date.now() + seconds * 1000;
        throw bilingual('Microsoft 暫時未能同步，稍後再試。', 'Microsoft sync is temporarily unavailable. Try again shortly.');
      }
      if (response.status === 403 || response.status === 401) {
        throw bilingual('Microsoft 拒絕存取。請確認你有工作簿編輯權限同應用程式獲准。', 'Microsoft denied access. Check workbook edit access and app approval.');
      }
      if (!response.ok) {
        throw bilingual('無法連接指定 Excel 工作簿（' + response.status + '）。', 'Could not access the Excel workbook (' + response.status + ').');
      }
      return response.json();
    }

    const item = await request(lookup.path + '?$select=id,name,webUrl,parentReference,file');
    if (!item.file || !/\.xlsx$/i.test(item.name) || !item.parentReference?.driveId) {
      throw bilingual('請使用包含 Assignments 同 Overview 嘅 .xlsx 工作簿。', 'Link to the .xlsx workbook with the Assignments and Overview sheets.');
    }
    const path = '/drives/' + encodeURIComponent(item.parentReference.driveId) + '/items/' + encodeURIComponent(item.id) + '/workbook';
    const sheets = await request(path + '/worksheets?$select=name');
    const names = new Set(sheets.value.map(sheet => sheet.name));
    if (!names.has('Assignments') || !names.has('Overview')) {
      throw bilingual('Excel 缺少 Assignments 或 Overview 工作表。', 'The workbook is missing Assignments or Overview.');
    }
    function range(sheet, address, patch) {
      return request(path + '/worksheets/' + sheet + "/range(address='" + address + "')" +
        (patch ? '' : '?$select=values'), patch);
    }
    let labels, previous;
    async function fetchState() {
      if (!labels) ({values:labels} = await range('Overview', 'A7:B18'));
      if (!previous) previous = names.has('Previous Seating') ? (await range('Previous%20Seating', 'A2:B129')).values : [];
      const grid = await range('Assignments', 'A' + assignmentRow + ':G' + assignmentEnd);
      return parseWorkbook(grid.values, labels, previous);
    }
    return {
      url:readLink(item.webUrl || lookup.shareUrl?.href).href,
      async read() {
        labels = null;
        previous = null;
        return (await fetchState()).state;
      },
      async move(change) {
        const plan = planMove(await fetchState(), change);
        await range('Assignments', plan.address, {values:plan.values});
        return {state:(await fetchState()).state, undo:plan.undo};
      },
      close:() => auth.clearCache({account}),
    };
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
    const read = () => parseWorkbook(grid, labels, previous);
    return {
      demo:true,
      read:async () => read().state,
      async move(change) {
        const plan = planMove(read(), change);
        const column = plan.field === 'group' ? 5 : 6;
        plan.values.forEach(([value], offset) => {
          if (value !== null) grid[plan.first - assignmentRow + offset][column] = value;
        });
        return {state:read().state, undo:plan.undo};
      },
    };
  }
  return {connect, demo};
})();
