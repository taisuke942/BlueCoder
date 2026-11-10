/**
 * 共有ドライブ検索 v2（2026-11-10）
 * B1: 名前の部分一致検索 / B2: 検索フォルダ / 3行目: 見出し / 4行目以降: 結果
 *
 * 設定: DRIVE_IDを変更し、「サービス +」でGoogle Drive API v3を追加。
 * setupDriveSearchを一度実行後、シートを再読み込みする。
 * 「共有ドライブ検索」メニューの「フォルダを選択」「新しく検索」を使用。
 * 約4分で区切り、「続きから検索」で再開する。自動トリガーは使用しない。
 * 検索・再開待ちの間はB1/B2、結果行、作業シートを編集・並べ替えない。
 *
 * B2は /（共有ドライブ全体）、/営業/見積、フォルダURL、id:フォルダID。
 * バックスラッシュ区切りも可。G:\\のようなPCのドライブパスは使用しない。
 * 全階層のファイル名・フォルダ名が対象。本文検索・誤字補正は行わない。
 * 全角半角と英字大小を正規化。空白区切りはAND検索にせず、そのまま検索。
 * ゴミ箱を除外し、フォルダへのショートカットは検索先として辿らない。
 * Drive上は読み取りのみ。シートの結果・入力表示・非表示作業表は更新する。
 * APIの一覧は時点固定のスナップショットではない。トークン失効時は新規検索。
 */

const DRIVE_SEARCH = {
  DRIVE_ID: 'ここに共有ドライブIDを入力',
  SHEET_NAME: '検索',
  PAGE_SIZE: 1000,
  RUN_LIMIT_MS: 4 * 60 * 1000,
  RESULT_FIRST_ROW: 4,
  STATE_KEY: 'SHARED_DRIVE_NAME_SEARCH_V2',
  SELECTION_KEY: 'SHARED_DRIVE_FOLDER_SELECTION_V2',
  QUEUE_SHEET_NAME: '_DriveSearchQueueV2',
  QUEUE_SIGNATURE: 'SHARED_DRIVE_SEARCH_V2_QUEUE',
  PARENT_BATCH_SIZE: 20
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('共有ドライブ検索')
    .addItem('フォルダを選択', 'showFolderPicker')
    .addSeparator()
    .addItem('新しく検索', 'startDriveSearch')
    .addItem('続きから検索', 'continueDriveSearch')
    .addToUi();
}

// 既存B1とv2のB2・結果は保持。旧版の見出しが完全一致した場合だけ移行する。
function setupDriveSearch() {
  const lock = acquireDriveSearchLock_();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(DRIVE_SEARCH.SHEET_NAME);
    const created = !sheet;
    if (!sheet) {
      sheet = ss.insertSheet(DRIVE_SEARCH.SHEET_NAME);
      if (sheet.getMaxColumns() > 8) sheet.deleteColumns(9, sheet.getMaxColumns() - 8);
    }
    ensureDriveSearchSize_(sheet, 4, 8);
    const headers = driveSearchHeaders_();
    const oldHeaders = sheet.getRange(2, 1, 1, 6).getDisplayValues()[0];
    const legacy = oldHeaders.every((value, i) => value === headers[i]);
    const currentHeaders = sheet.getRange(3, 1, 1, 6).getDisplayValues()[0];
    if (!created && !legacy && currentHeaders.some(String) &&
        !currentHeaders.every((value, i) => value === headers[i])) {
      throw new Error('「検索」シートの3行目に別のデータがあります。移動してから初期設定してください。');
    }
    if (legacy) {
      sheet.getRange(2, 1, Math.max(sheet.getLastRow(), 3) - 1, 6).clearContent();
      const props = PropertiesService.getDocumentProperties();
      props.deleteProperty(DRIVE_SEARCH.STATE_KEY);
      props.deleteProperty(DRIVE_SEARCH.SELECTION_KEY);
    }
    sheet.getRange('A1').setValue('検索語');
    sheet.getRange('B1').setBackground('#fff2cc');
    sheet.getRange('C1').setValue('状態');
    sheet.getRange('A2').setValue('検索場所');
    const pathRange = sheet.getRange('B2:F2');
    if (!pathRange.isPartOfMerge() &&
        sheet.getRange('C2:F2').getDisplayValues()[0].every(value => value === '')) {
      pathRange.merge();
    }
    sheet.getRange('B2').setNumberFormat('@').setBackground('#e8f0fe');
    if (created || legacy) {
      sheet.getRange('B2').setValue('/');
      sheet.getRange('D1').setValue('未検索');
    }
    sheet.getRange(3, 1, 1, 6).setValues([headers])
      .setFontWeight('bold').setBackground('#d9eaf7');
    sheet.setFrozenRows(3);
    sheet.setColumnWidth(1, 70);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 350);
    sheet.setColumnWidth(4, 180);
    sheet.setColumnWidths(5, 2, 300);
    sheet.setColumnWidths(7, 2, 140);
    ss.setActiveSheet(sheet);
    SpreadsheetApp.flush();
    onOpen();
    ss.toast(legacy ? '旧版を移行しました。B2が検索場所、結果は4行目からです。'
      : 'B1に検索語、B2に検索場所を指定してください。', '共有ドライブ検索', 8);
  } finally {
    lock.releaseLock();
  }
}

function startDriveSearch() { runDriveSearch_(true); }
function continueDriveSearch() { runDriveSearch_(false); }

function runDriveSearch_(restart) {
  const started = Date.now();
  const lock = acquireDriveSearchLock_();
  let sheet;
  let processing = false;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = getDriveSearchSheet_(ss);
    const props = PropertiesService.getDocumentProperties();
    const keyword = sheet.getRange('B1').getDisplayValue().trim();
    if (!keyword) throw new Error('B1に検索する文字を入力してください。');
    const key = normalizeDriveName_(keyword);
    const owner = Session.getTemporaryActiveUserKey();
    const root = getDriveSearchRoot_();
    let state;
    let queue;

    if (restart) {
      const scope = resolveDriveSearchScope_(sheet, root);
      if (scope.id !== root.id) queue = getDriveSearchQueue_(ss, true);
      // URLや手入力パスを、検証済みの現在のパスへそろえる。
      writeDriveSearchSelection_(sheet, scope);
      state = {
        version: 2, driveId: root.id, sheetId: sheet.getSheetId(), owner: owner,
        keywordHash: driveSearchHash_(keyword), scopeInputHash: driveSearchHash_(scope.path),
        scopeId: scope.id, mode: scope.id === root.id ? 'drive' : 'subtree',
        token: '', nextRow: DRIVE_SEARCH.RESULT_FIRST_ROW, checked: 0, hits: 0, done: false,
        queueSheetId: queue ? queue.getSheetId() : null,
        queueReadIndex: 0, queueLength: 0, activeBatch: []
      };
      // 新規検索を先に記録。初期化途中で止まっても前回のキューを再利用しない。
      saveDriveSearchState_(props, state);
    } else {
      const saved = props.getProperty(DRIVE_SEARCH.STATE_KEY);
      if (!saved) throw new Error('再開する検索がありません。「新しく検索」を実行してください。');
      state = JSON.parse(saved);
      if (state.owner !== owner) throw new Error('検索を開始したGoogleアカウントで再開してください。');
      const input = sheet.getRange('B2').getDisplayValue();
      if (state.version !== 2 || state.driveId !== root.id || state.sheetId !== sheet.getSheetId() ||
          state.keywordHash !== driveSearchHash_(keyword) || state.scopeInputHash !== driveSearchHash_(input)) {
        throw new Error('B1・B2または設定が変わっています。「新しく検索」を実行してください。');
      }
      const scope = resolveDriveSearchScope_(sheet, root);
      if (state.scopeId !== scope.id) {
        throw new Error('検索フォルダが変わっています。「新しく検索」を実行してください。');
      }
      if (state.done) { showDriveSearchStatus_(sheet, state, '完了'); return; }
      if (state.mode === 'subtree') queue = getDriveSearchQueue_(ss, false, state.queueSheetId);
    }

    processing = true;
    clearDriveSearchTail_(sheet, state.nextRow, 6);
    if (queue) {
      clearDriveSearchTail_(queue, state.queueLength + 2, 1);
      if (!state.queueLength) {
        ensureDriveSearchSize_(queue, 2, 1);
        queue.getRange(2, 1).setValue(state.scopeId);
        SpreadsheetApp.flush();
        state.queueLength = 1;
        saveDriveSearchState_(props, state);
      }
    }
    showDriveSearchStatus_(sheet, state, '検索中');
    SpreadsheetApp.flush();

    while (Date.now() - started < DRIVE_SEARCH.RUN_LIMIT_MS) {
      if (queue && !state.activeBatch.length) {
        const length = Math.min(DRIVE_SEARCH.PARENT_BATCH_SIZE, state.queueLength - state.queueReadIndex);
        if (length <= 0) { state.done = true; saveDriveSearchState_(props, state); break; }
        state.activeBatch = queue.getRange(state.queueReadIndex + 2, 1, length, 1)
          .getDisplayValues().map(row => row[0]);
        if (state.activeBatch.some(id => !isDriveSearchId_(id))) {
          throw new Error('検索用作業シートが変更されています。「新しく検索」でやり直してください。');
        }
        state.queueReadIndex += length;
        state.token = '';
      }
      const options = driveSearchListOptions_(root.id);
      options.fields = 'nextPageToken,incompleteSearch,files(id,name,mimeType,modifiedTime,webViewLink,parents)';
      if (queue) {
        options.q += ' and (' + state.activeBatch
          .map(id => "'" + escapeDriveQuery_(id) + "' in parents").join(' or ') + ')';
      }
      if (state.token) options.pageToken = state.token;
      const page = Drive.Files.list(options);
      if (page.incompleteSearch) throw new Error('Drive APIが検索未完了を返しました。「新しく検索」でやり直してください。');
      const files = page.files || [];
      const rows = files.filter(file => normalizeDriveName_(file.name).includes(key))
        .map((file, i) => driveSearchResultRow_(file, state.hits + i + 1));
      if (rows.length) {
        ensureDriveSearchSize_(sheet, state.nextRow + rows.length - 1, 6);
        sheet.getRange(state.nextRow, 3, rows.length, 1).setNumberFormat('@');
        sheet.getRange(state.nextRow, 1, rows.length, 6).setValues(rows);
        sheet.getRange(state.nextRow, 4, rows.length, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
      }
      const children = queue ? files.filter(file => file.mimeType === 'application/vnd.google-apps.folder')
        .map(file => [file.id]) : [];
      if (children.length) {
        ensureDriveSearchSize_(queue, state.queueLength + children.length + 1, 1);
        queue.getRange(state.queueLength + 2, 1, children.length, 1).setValues(children);
      }
      state.checked += files.length;
      state.hits += rows.length;
      state.nextRow += rows.length;
      state.queueLength += children.length;
      state.token = page.nextPageToken || '';
      if (queue && !state.token) state.activeBatch = [];
      state.done = queue ? !state.token && state.queueReadIndex >= state.queueLength : !state.token;
      showDriveSearchStatus_(sheet, state, '検索中');
      // 結果と作業キューの両方を確定後にcheckpoint。保存失敗時は再開時に末尾を戻す。
      SpreadsheetApp.flush();
      saveDriveSearchState_(props, state);
      if (state.done) break;
    }
    showDriveSearchStatus_(sheet, state, state.done ? '完了' : '一時停止：「続きから検索」を実行');
    SpreadsheetApp.flush();
    ss.toast(state.done ? '検索が完了しました。' : '続きから検索できます。', '共有ドライブ検索', 8);
  } catch (error) {
    if (sheet && processing) sheet.getRange('D1').setValue('中断（未完了）：' + error.message);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function showFolderPicker() {
  getDriveSearchSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const html = HtmlService.createHtmlOutputFromFile('FolderPicker').setWidth(760).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '検索するフォルダを選択');
}

// HTMLから呼ぶ公開関数。Dateを返さず、現在フォルダの直下だけ取得する。
function getFolderPickerPage(folderId, pageToken) {
  const sheet = getDriveSearchSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const root = getDriveSearchRoot_();
  let info;
  let warning = '';
  if (!folderId) {
    try { info = resolveDriveSearchScope_(sheet, root); }
    catch (error) { info = root; warning = 'B2の場所を開けないためルートを表示します。' + error.message; }
  } else {
    info = getDriveSearchFolderInfo_(String(folderId), root);
  }
  const options = driveSearchListOptions_(root.id);
  options.q += " and '" + escapeDriveQuery_(info.id) + "' in parents"
    + " and mimeType = 'application/vnd.google-apps.folder'";
  options.pageSize = 100;
  options.orderBy = 'name_natural';
  options.fields = 'nextPageToken,incompleteSearch,files(id,name)';
  if (pageToken) options.pageToken = String(pageToken);
  const page = Drive.Files.list(options);
  if (page.incompleteSearch) throw new Error('フォルダ一覧を取得しきれませんでした。もう一度開いてください。');
  return {
    rootId: root.id, rootName: root.name, folderId: info.id, folderName: info.name,
    path: info.path, breadcrumbs: info.breadcrumbs,
    folders: (page.files || []).map(file => ({ id: file.id, name: file.name })),
    nextPageToken: page.nextPageToken || '', warning: warning
  };
}

function selectSearchFolder(folderId) {
  const lock = acquireDriveSearchLock_();
  try {
    const sheet = getDriveSearchSheet_(SpreadsheetApp.getActiveSpreadsheet());
    const root = getDriveSearchRoot_();
    const info = getDriveSearchFolderInfo_(String(folderId || ''), root);
    writeDriveSearchSelection_(sheet, info);
    return { id: info.id, path: info.path };
  } finally {
    lock.releaseLock();
  }
}

function getDriveSearchRoot_() {
  if (typeof Drive === 'undefined') throw new Error('「サービス +」でGoogle Drive API（v3）を追加してください。');
  const id = String(DRIVE_SEARCH.DRIVE_ID).trim();
  if (!isDriveSearchId_(id)) throw new Error('DRIVE_IDに共有ドライブ自体のIDを設定してください。');
  const drive = Drive.Drives.get(id, { fields: 'id,name' });
  return { id: id, name: drive.name, driveId: id, path: '/', breadcrumbs: [{ id: id, name: drive.name }] };
}

// 選択済みの表示文字列に対応するIDを優先。同名フォルダもIDで区別する。
function resolveDriveSearchScope_(sheet, root) {
  const rawInput = sheet.getRange('B2').getDisplayValue();
  const saved = PropertiesService.getDocumentProperties().getProperty(DRIVE_SEARCH.SELECTION_KEY);
  if (saved) {
    const selection = JSON.parse(saved);
    if (selection.driveId === root.id && selection.sheetId === sheet.getSheetId() &&
        (selection.pathHash ? selection.pathHash === driveSearchHash_(rawInput) : selection.path === rawInput)) {
      return getDriveSearchFolderInfo_(selection.id, root);
    }
  }
  // 選択画面で選んだ名前の末尾空白は保持。手入力を解決するときだけ外側の空白を除く。
  const input = rawInput.trim();
  if (!input || /^[\\/]+$/.test(input)) return root;
  if (/^[A-Za-z]:[\\/]/.test(input)) {
    throw new Error('G:\\などPCのパスは使えません。共有ドライブを起点に「/営業/見積」と入力するか、フォルダを選択してください。');
  }
  let id = '';
  if (/^id:/i.test(input)) {
    id = input.slice(3).trim();
    if (!isDriveSearchId_(id)) throw new Error('id:の後ろにフォルダIDを入力してください。');
  } else if (/^https?:\/\//i.test(input)) {
    const host = input.match(/^https?:\/\/([^/]+)/i);
    if (!host || host[1].toLowerCase() !== 'drive.google.com') throw new Error('Google DriveのフォルダURLを入力してください。');
    const match = input.match(/\/folders\/([A-Za-z0-9_-]+)/) || input.match(/[?&]id=([A-Za-z0-9_-]+)/);
    if (!match) throw new Error('Google DriveのフォルダURLを入力してください。');
    id = match[1];
  } else if (isDriveSearchId_(input) && input.length >= 10) {
    // 裸の長い英数字はIDとして扱う。同じ形のフォルダ名は / を先頭に付ける。
    id = input;
  }
  if (id) return getDriveSearchFolderInfo_(id, root);

  const segments = input.replace(/\\/g, '/').split('/').filter(part => part !== '');
  let currentId = root.id;
  for (const segment of segments) {
    const options = driveSearchListOptions_(root.id);
    options.q += " and '" + escapeDriveQuery_(currentId) + "' in parents"
      + " and mimeType = 'application/vnd.google-apps.folder' and name = '" + escapeDriveQuery_(segment) + "'";
    options.fields = 'nextPageToken,incompleteSearch,files(id,name)';
    options.pageSize = 100;
    const matches = new Map();
    do {
      const page = Drive.Files.list(options);
      if (page.incompleteSearch) throw new Error('検索場所を確認しきれませんでした。もう一度実行してください。');
      (page.files || []).forEach(file => matches.set(file.id, file));
      if (matches.size > 1) throw new Error('「' + segment + '」という同名フォルダが複数あります。選択画面またはフォルダURLで指定してください。');
      options.pageToken = page.nextPageToken || '';
    } while (options.pageToken);
    if (!matches.size) throw new Error('検索場所に「' + segment + '」が見つかりません。');
    currentId = matches.values().next().value.id;
  }
  return getDriveSearchFolderInfo_(currentId, root);
}

// IDの実体と所属共有ドライブを確認し、親を辿ってルートまでの経路を作る。
function getDriveSearchFolderInfo_(folderId, root) {
  if (!isDriveSearchId_(folderId)) throw new Error('フォルダIDが正しくありません。');
  if (folderId === root.id) return root;
  const chain = [];
  const seen = new Set();
  let id = folderId;
  while (id !== root.id) {
    if (seen.has(id) || chain.length >= 500) throw new Error('フォルダ階層を確認できません。フォルダURLを確認してください。');
    seen.add(id);
    const file = Drive.Files.get(id, {
      supportsAllDrives: true, fields: 'id,name,mimeType,trashed,parents,driveId'
    });
    if (file.mimeType !== 'application/vnd.google-apps.folder' || file.trashed) {
      throw new Error('ゴミ箱以外の通常のフォルダを指定してください。');
    }
    if (file.driveId !== root.id) throw new Error('設定した共有ドライブの外のフォルダは指定できません。');
    chain.push({ id: file.id, name: file.name });
    if (!file.parents || !file.parents.length) throw new Error('共有ドライブまでの親フォルダを確認できません。');
    id = file.parents[0];
  }
  chain.reverse();
  return {
    id: folderId, name: chain[chain.length - 1].name, driveId: root.id,
    path: '/' + chain.map(folder => folder.name).join('/'),
    breadcrumbs: [{ id: root.id, name: root.name }].concat(chain)
  };
}

function writeDriveSearchSelection_(sheet, info) {
  const selection = {
    id: info.id, driveId: info.driveId, sheetId: sheet.getSheetId(), pathHash: driveSearchHash_(info.path)
  };
  if (info.path.length <= 1500) selection.path = info.path;
  sheet.getRange('B2').setNumberFormat('@').setValue(safeDriveSearchText_(info.path));
  SpreadsheetApp.flush();
  setDriveSearchProperty_(PropertiesService.getDocumentProperties(), DRIVE_SEARCH.SELECTION_KEY, selection);
}

function driveSearchListOptions_(driveId) {
  return {
    corpora: 'drive', driveId: driveId, includeItemsFromAllDrives: true,
    supportsAllDrives: true, spaces: 'drive', q: 'trashed = false', pageSize: DRIVE_SEARCH.PAGE_SIZE
  };
}

function getDriveSearchQueue_(ss, create, sheetId) {
  let sheet = sheetId != null ? ss.getSheets().find(item => item.getSheetId() === sheetId)
    : ss.getSheetByName(DRIVE_SEARCH.QUEUE_SHEET_NAME);
  if (!sheet) {
    if (!create) throw new Error('検索用作業シートがありません。「新しく検索」でやり直してください。');
    sheet = ss.insertSheet(DRIVE_SEARCH.QUEUE_SHEET_NAME);
    if (sheet.getMaxColumns() > 1) sheet.deleteColumns(2, sheet.getMaxColumns() - 1);
    sheet.getRange('A1').setValue(DRIVE_SEARCH.QUEUE_SIGNATURE);
  } else if (sheet.getRange('A1').getDisplayValue() !== DRIVE_SEARCH.QUEUE_SIGNATURE) {
    throw new Error('「' + DRIVE_SEARCH.QUEUE_SHEET_NAME + '」は検索用の作業シートではありません。既存データは変更しません。別の名前へ変更してください。');
  }
  if (!sheet.isSheetHidden()) sheet.hideSheet();
  return sheet;
}

function getDriveSearchSheet_(ss) {
  const sheet = ss.getSheetByName(DRIVE_SEARCH.SHEET_NAME);
  if (!sheet) throw new Error('先にsetupDriveSearchを実行してください。');
  const headers = driveSearchHeaders_();
  if (sheet.getMaxRows() < 4 || sheet.getMaxColumns() < 6 ||
      !sheet.getRange(3, 1, 1, 6).getDisplayValues()[0].every((value, i) => value === headers[i])) {
    throw new Error('先にsetupDriveSearchを実行し、v2のシート配置へ更新してください。');
  }
  return sheet;
}

function acquireDriveSearchLock_() {
  const lock = LockService.getDocumentLock();
  if (!lock || !lock.tryLock(1000)) throw new Error('別の処理が実行中です。終了後に実行してください。');
  return lock;
}

function saveDriveSearchState_(props, state) {
  setDriveSearchProperty_(props, DRIVE_SEARCH.STATE_KEY, state);
}

function setDriveSearchProperty_(props, name, value) {
  const json = JSON.stringify(value);
  if (Utilities.newBlob(json).getBytes().length > 8500) {
    throw new Error('検索の進捗情報が保存上限に近づきました。対象フォルダを狭めて新しく検索してください。');
  }
  props.setProperty(name, json);
}

function driveSearchHash_(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value), Utilities.Charset.UTF_8)
    .map(byte => ('0' + (byte & 255).toString(16)).slice(-2)).join('');
}

function driveSearchResultRow_(file, number) {
  const folder = file.mimeType === 'application/vnd.google-apps.folder';
  const shortcut = file.mimeType === 'application/vnd.google-apps.shortcut';
  return [
    number, folder ? 'フォルダ' : (shortcut ? 'ショートカット' : 'ファイル'),
    safeDriveSearchText_(file.name || ''), file.modifiedTime ? new Date(file.modifiedTime) : '',
    file.webViewLink || (folder ? 'https://drive.google.com/drive/folders/' + file.id
      : 'https://drive.google.com/file/d/' + file.id + '/view'),
    file.parents && file.parents.length ? 'https://drive.google.com/drive/folders/' + file.parents[0] : ''
  ];
}

function driveSearchHeaders_() { return ['No.', '種類', '名前', '更新日時', '開くURL', '親フォルダURL']; }
function normalizeDriveName_(value) { return String(value || '').normalize('NFKC').toLowerCase(); }
function isDriveSearchId_(value) { return /^[A-Za-z0-9_-]+$/.test(String(value || '')); }
function escapeDriveQuery_(value) { return String(value).replace(/\\/g, '\\\\').replace(/'/g, "\\'"); }
function safeDriveSearchText_(value) { const text = String(value); return /^[=']/.test(text) ? "'" + text : text; }

function ensureDriveSearchSize_(sheet, rows, columns) {
  if (sheet.getMaxColumns() < columns) sheet.insertColumnsAfter(sheet.getMaxColumns(), columns - sheet.getMaxColumns());
  if (sheet.getMaxRows() < rows) sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
}

function clearDriveSearchTail_(sheet, row, columns) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= row) sheet.getRange(row, 1, lastRow - row + 1, columns).clearContent();
}

function showDriveSearchStatus_(sheet, state, label) {
  sheet.getRange('D1').setValue(label + ' ｜ 確認 ' + state.checked.toLocaleString('ja-JP')
    + '件 ／ 該当 ' + state.hits.toLocaleString('ja-JP') + '件');
}
