/**
 * 共有ドライブ検索 v3.1（2026-11-12）
 * B1: 名前の部分一致検索 / B2: 検索フォルダ / 3行目: 見出し / 4行目以降: 結果
 *
 * 設定: DRIVE_IDを変更し、「サービス +」でGoogle Drive API v3を追加。
 * setupDriveSearchを一度実行後、シートを再読み込みする。
 * 「共有ドライブ検索」メニューの「フォルダを選択」「新しく検索」を使用。
 * 約4分で区切り、「続きから検索」で再開する。自動トリガーは使用しない。
 * 一時停止・終了はページ処理の区切りで反映。実行中のDrive APIは即時中断できない。
 * D1:F1に状態を常設表示。緑の「検索完了」になるまでは途中結果。
 * v3の保存形式を維持。「状態を更新」で中断・条件変更も確認できる。
 * 検索・再開待ちの間はB1/B2、結果行、作業シートを編集・並べ替えない。
 *
 * B2は //（共有ドライブ一覧全体）、//ドライブ名/営業/見積、URL、id:フォルダID。
 * / と /営業/見積 は設定したDRIVE_IDを起点とする。選択後は // 形式で表示。
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
  STATE_KEY: 'SHARED_DRIVE_NAME_SEARCH_V3',
  CONTROL_KEY: 'SHARED_DRIVE_SEARCH_CONTROL_V3',
  SELECTION_KEY: 'SHARED_DRIVE_FOLDER_SELECTION_V3',
  LEGACY_SELECTION_KEY: 'SHARED_DRIVE_FOLDER_SELECTION_V2',
  QUEUE_SHEET_NAME: '_DriveSearchQueueV3',
  QUEUE_SIGNATURE: 'SHARED_DRIVE_SEARCH_V3_QUEUE',
  ALL_DRIVES_ID: '__SHARED_DRIVES__',
  PARENT_BATCH_SIZE: 20
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('共有ドライブ検索')
    .addItem('フォルダを選択', 'showFolderPicker')
    .addSeparator()
    .addItem('新しく検索', 'startDriveSearch')
    .addItem('続きから検索', 'continueDriveSearch')
    .addItem('一時停止', 'pauseDriveSearch')
    .addItem('終了（再開しない）', 'stopDriveSearch')
    .addSeparator()
    .addItem('状態を更新', 'refreshDriveSearchStatus')
    .addToUi();
  refreshDriveSearchStatus_(true);
}

// セル編集だけでは検索しない。前回の完了表示を現在の条件と区別する。
function onEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  if (range.getSheet().getName() !== DRIVE_SEARCH.SHEET_NAME) return;
  const lastRow = range.getRow() + range.getNumRows() - 1;
  const lastColumn = range.getColumn() + range.getNumColumns() - 1;
  const touchesB1 = range.getRow() <= 1 && lastRow >= 1 && range.getColumn() <= 2 && lastColumn >= 2;
  const touchesB2 = range.getRow() <= 2 && lastRow >= 2 && range.getColumn() <= 6 && lastColumn >= 2;
  if (touchesB1 || touchesB2) refreshDriveSearchStatus_(true);
}

// 既存B1/B2とv2・v3の結果は保持。旧版の2行目見出しが完全一致した場合だけ移行する。
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
    setupDriveSearchStatusArea_(sheet);
    if (legacy) {
      sheet.getRange(2, 1, Math.max(sheet.getLastRow(), 3) - 1, 6).clearContent();
      const props = PropertiesService.getDocumentProperties();
      props.deleteProperty(DRIVE_SEARCH.STATE_KEY);
      props.deleteProperty(DRIVE_SEARCH.SELECTION_KEY);
      props.deleteProperty(DRIVE_SEARCH.CONTROL_KEY);
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
    refreshDriveSearchStatusLocked_(sheet, PropertiesService.getDocumentProperties());
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
  let props;
  let state;
  let processing = false;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = getDriveSearchSheet_(ss);
    props = PropertiesService.getDocumentProperties();
    if (restart) {
      markDriveSearchPreparing_(sheet, props);
      showDriveSearchStatus_(sheet, { status: 'preparing', checked: 0, hits: 0 },
        '新しい検索を準備中です。表示中の行は前回の結果です。');
      SpreadsheetApp.flush();
    }
    requireDriveSearchApi_();
    const keyword = sheet.getRange('B1').getDisplayValue().trim();
    if (!keyword) throw new Error('B1に検索する文字を入力してください。');
    const key = normalizeDriveName_(keyword);
    const owner = Session.getTemporaryActiveUserKey();
    const defaultDriveId = String(DRIVE_SEARCH.DRIVE_ID).trim();
    let queue;

    if (restart) {
      const scope = resolveDriveSearchScope_(sheet);
      const mode = scope.kind === 'folder' ? 'subtree' : scope.kind;
      if (mode !== 'drive') queue = getDriveSearchQueue_(ss, true);
      writeDriveSearchSelection_(sheet, scope);
      state = {
        version: 3, jobId: Utilities.getUuid(), owner: owner, status: 'running', stopped: false,
        startedAt: Date.now(), finishedAt: null,
        defaultDriveId: defaultDriveId, driveId: scope.driveId, sheetId: sheet.getSheetId(),
        keywordHash: driveSearchHash_(keyword), scopeInputHash: driveSearchHash_(scope.path),
        scopeId: scope.id, mode: mode, token: '', nextRow: DRIVE_SEARCH.RESULT_FIRST_ROW,
        checked: 0, hits: 0, done: false, queueSheetId: queue ? queue.getSheetId() : null,
        queueReadIndex: 0, queueLength: 0, activeBatch: [],
        driveListToken: '', driveListDone: mode !== 'allDrives', activeDriveId: ''
      };
      withDriveSearchControlLock_(() => {
        saveDriveSearchState_(props, state);
        props.deleteProperty(DRIVE_SEARCH.CONTROL_KEY);
      });
    } else {
      const saved = props.getProperty(DRIVE_SEARCH.STATE_KEY);
      if (!saved) throw new Error('再開するv3の検索がありません。旧版の進捗は引き継げないため「新しく検索」を実行してください。');
      state = JSON.parse(saved);
      if (state.owner !== owner) throw new Error('検索を開始したGoogleアカウントで再開してください。');
      if (state.version !== 3 || state.defaultDriveId !== defaultDriveId ||
          state.sheetId !== sheet.getSheetId() || state.keywordHash !== driveSearchHash_(keyword) ||
          state.scopeInputHash !== driveSearchHash_(sheet.getRange('B2').getDisplayValue())) {
        throw new Error('B1・B2または設定が変わっています。「新しく検索」を実行してください。');
      }
      const scope = resolveDriveSearchScope_(sheet);
      if (state.scopeId !== scope.id || state.driveId !== scope.driveId) {
        throw new Error('検索場所が変わっています。「新しく検索」を実行してください。');
      }
      withDriveSearchControlLock_(() => {
        const request = readDriveSearchControl_(props, state);
        if (state.stopped || (request && request.action === 'stop')) {
          throw new Error('この検索は終了しています。続ける場合は「新しく検索」を実行してください。');
        }
        if (request && request.action === 'pause') props.deleteProperty(DRIVE_SEARCH.CONTROL_KEY);
        delete state.lastAttempt;
        delete state.lastError;
        if (!state.done) { state.status = 'running'; saveDriveSearchState_(props, state); }
      });
      if (state.done) { finishDriveSearchRun_(ss, sheet, props, state); return; }
      if (state.mode !== 'drive') queue = getDriveSearchQueue_(ss, false, state.queueSheetId);
    }

    processing = true;
    clearDriveSearchTail_(sheet, state.nextRow, 6);
    if (queue) {
      clearDriveSearchTail_(queue, state.queueLength + 2, 1);
      if (state.mode === 'subtree' && !state.queueLength) {
        appendDriveSearchQueue_(queue, 0, [state.scopeId]);
        SpreadsheetApp.flush();
        state.queueLength = 1;
        saveDriveSearchState_(props, state);
      }
    }
    showDriveSearchStatus_(sheet, state);
    SpreadsheetApp.flush();

    while (!state.done && Date.now() - started < DRIVE_SEARCH.RUN_LIMIT_MS) {
      if (readDriveSearchControl_(props, state)) {
        finishDriveSearchRun_(ss, sheet, props, state);
        return;
      }
      if (state.mode === 'allDrives' && !state.driveListDone) {
        const options = driveSearchDrivesOptions_();
        if (state.driveListToken) options.pageToken = state.driveListToken;
        const page = Drive.Drives.list(options);
        const ids = (page.drives || []).map(drive => drive.id);
        appendDriveSearchQueue_(queue, state.queueLength, ids);
        state.queueLength += ids.length;
        state.driveListToken = page.nextPageToken || '';
        state.driveListDone = !state.driveListToken;
        state.done = state.driveListDone && !state.queueLength;
        SpreadsheetApp.flush();
        saveDriveSearchState_(props, state);
        showDriveSearchStatus_(sheet, state, '共有ドライブ一覧を取得中です。まだ検索は完了していません。');
        SpreadsheetApp.flush();
      } else if (prepareDriveSearchPage_(queue, state)) {
        processDriveSearchPage_(sheet, queue, props, state, key);
      } else {
        state.done = true;
        saveDriveSearchState_(props, state);
      }
      if (readDriveSearchControl_(props, state)) {
        finishDriveSearchRun_(ss, sheet, props, state);
        return;
      }
    }
    finishDriveSearchRun_(ss, sheet, props, state);
  } catch (error) {
    // 未確定のローカルstateは保存しない。確定済みcheckpointの状態表示だけ更新する。
    if (sheet && props) reportDriveSearchError_(sheet, props, state, processing, error);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function prepareDriveSearchPage_(queue, state) {
  if (state.mode === 'drive') return true;
  if (state.mode === 'allDrives') {
    if (!state.activeDriveId) {
      if (state.queueReadIndex >= state.queueLength) return false;
      state.activeDriveId = queue.getRange(state.queueReadIndex + 2, 1).getDisplayValue();
      if (!isDriveSearchId_(state.activeDriveId)) throw new Error('検索用作業シートが変更されています。「新しく検索」でやり直してください。');
      state.queueReadIndex++;
      state.token = '';
    }
  } else if (!state.activeBatch.length) {
    const length = Math.min(DRIVE_SEARCH.PARENT_BATCH_SIZE, state.queueLength - state.queueReadIndex);
    if (length <= 0) return false;
    state.activeBatch = queue.getRange(state.queueReadIndex + 2, 1, length, 1)
      .getDisplayValues().map(row => row[0]);
    if (state.activeBatch.some(id => !isDriveSearchId_(id))) {
      throw new Error('検索用作業シートが変更されています。「新しく検索」でやり直してください。');
    }
    state.queueReadIndex += length;
    state.token = '';
  }
  return true;
}

function processDriveSearchPage_(sheet, queue, props, state, key) {
  const options = driveSearchListOptions_(state.mode === 'allDrives' ? state.activeDriveId : state.driveId);
  options.fields = 'nextPageToken,incompleteSearch,files(id,name,mimeType,modifiedTime,webViewLink,parents)';
  if (state.mode === 'subtree') {
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
  const children = state.mode === 'subtree' ? files
    .filter(file => file.mimeType === 'application/vnd.google-apps.folder').map(file => file.id) : [];
  appendDriveSearchQueue_(queue, state.queueLength, children);
  state.checked += files.length;
  state.hits += rows.length;
  state.nextRow += rows.length;
  state.queueLength += children.length;
  state.token = page.nextPageToken || '';
  if (state.mode === 'subtree' && !state.token) state.activeBatch = [];
  if (state.mode === 'allDrives' && !state.token) state.activeDriveId = '';
  state.done = state.mode === 'drive' ? !state.token
    : !state.token && state.queueReadIndex >= state.queueLength;
  SpreadsheetApp.flush();
  saveDriveSearchState_(props, state);
  showDriveSearchStatus_(sheet, state);
  SpreadsheetApp.flush();
}

function appendDriveSearchQueue_(queue, length, ids) {
  if (!ids.length) return;
  if (ids.some(id => !isDriveSearchId_(id))) throw new Error('Driveから不正なフォルダIDが返されました。検索をやり直してください。');
  ensureDriveSearchSize_(queue, length + ids.length + 1, 1);
  queue.getRange(length + 2, 1, ids.length, 1).setValues(ids.map(id => [id]));
}

// この2関数は本体が保持するDocumentLockを取得しない。
function pauseDriveSearch() { return requestDriveSearchControl_('pause'); }
function stopDriveSearch() { return requestDriveSearchControl_('stop'); }

function requestDriveSearchControl_(action) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let message;
  let result;
  withDriveSearchControlLock_(() => {
    const props = PropertiesService.getDocumentProperties();
    const saved = props.getProperty(DRIVE_SEARCH.STATE_KEY);
    if (!saved) throw new Error('操作する検索がありません。');
    const state = JSON.parse(saved);
    if (state.owner !== Session.getTemporaryActiveUserKey()) throw new Error('検索を開始したGoogleアカウントで操作してください。');
    const sheet = getDriveSearchSheet_(ss);
    if (state.sheetId !== sheet.getSheetId()) throw new Error('検索先のシートが変わっています。');
    if (state.lastAttempt && state.lastAttempt.kind === 'preparing') {
      message = '検索の準備中です。「検索中」の表示になってから操作してください。';
      showDriveSearchStatus_(sheet, { status: 'preparing', checked: 0, hits: 0 });
      SpreadsheetApp.flush();
      result = { action: action, requested: false };
      return;
    }
    if (state.stopped || (state.done && state.status === 'done')) {
      message = state.stopped ? 'この検索は終了しています。' : '検索は完了しています。';
      showDriveSearchStatus_(sheet, state);
      SpreadsheetApp.flush();
      result = { action: state.stopped ? 'stop' : action, requested: false };
      return;
    }
    const previous = readDriveSearchControl_(props, state);
    const effective = previous && previous.action === 'stop' ? 'stop' : action;
    setDriveSearchProperty_(props, DRIVE_SEARCH.CONTROL_KEY,
      { jobId: state.jobId, owner: state.owner, action: effective });
    const paused = ['paused', 'error', 'interrupted'].includes(state.status);
    message = effective === 'stop'
      ? (paused ? '終了しました。続きから検索はできません。' : '終了要求を受け付けました。現在のページ処理後に終了します。')
      : (paused ? '一時停止しています。続きから検索できます。' : '一時停止要求を受け付けました。現在のページ処理後に停止します。');
    if (paused && effective === 'stop') {
      state.stopped = true;
      state.status = 'stopped';
      saveDriveSearchState_(props, state);
    }
    const displayState = Object.assign({}, state);
    if (!paused) displayState.status = effective === 'stop' ? 'stop_requested' : 'pause_requested';
    showDriveSearchStatus_(sheet, displayState);
    SpreadsheetApp.flush();
    result = { action: effective, requested: true };
  });
  ss.toast(message, '共有ドライブ検索', 8);
  return result;
}

function readDriveSearchControl_(props, state) {
  const saved = props.getProperty(DRIVE_SEARCH.CONTROL_KEY);
  if (!saved) return null;
  const request = JSON.parse(saved);
  return request.jobId === state.jobId && request.owner === state.owner &&
    (request.action === 'pause' || request.action === 'stop') ? request : null;
}

// 終了表示とcontrol更新の競合を短いScriptLockで防ぐ。Drive APIはこのlock内では呼ばない。
function finishDriveSearchRun_(ss, sheet, props, state) {
  let label;
  withDriveSearchControlLock_(() => {
    const request = readDriveSearchControl_(props, state);
    if (request && request.action === 'stop') {
      state.stopped = true;
      state.status = 'stopped';
      label = '未完了のまま終了しました。続きから検索はできません。';
    } else if (state.done) {
      state.status = 'done';
      if (typeof state.finishedAt !== 'number') state.finishedAt = Date.now();
      label = '検索完了：指定した範囲の検索処理が最後まで終了しました。';
    } else if (request && request.action === 'pause') {
      state.status = 'paused';
      state.pauseReason = 'manual';
      label = '未完了：一時停止しました。「続きから検索」で再開してください。';
    } else {
      state.status = 'paused';
      state.pauseReason = 'time';
      label = '未完了：時間の区切りで一時停止しました。「続きから検索」で再開してください。';
    }
    saveDriveSearchState_(props, state);
    showDriveSearchStatus_(sheet, state);
    SpreadsheetApp.flush();
  });
  ss.toast(label, '共有ドライブ検索', 8);
}

function withDriveSearchControlLock_(callback) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('操作の受付処理が重なっています。もう一度実行してください。');
  try { return callback(); } finally { lock.releaseLock(); }
}

function showFolderPicker() {
  getDriveSearchSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const html = HtmlService.createHtmlOutputFromFile('FolderPicker').setWidth(760).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '検索するフォルダを選択');
}

// HTMLから呼ぶ公開関数。Dateを返さず、現在フォルダの直下だけ取得する。
function getFolderPickerPage(folderId, pageToken) {
  requireDriveSearchApi_();
  const sheet = getDriveSearchSheet_(SpreadsheetApp.getActiveSpreadsheet());
  let info;
  let warning = '';
  if (!folderId) {
    try { info = resolveDriveSearchScope_(sheet); }
    catch (error) {
      info = getDriveSearchVirtualRoot_();
      warning = 'B2の場所を開けないため共有ドライブ一覧を表示します。' + error.message;
    }
  } else {
    info = getDriveSearchFolderInfo_(String(folderId));
  }
  let page;
  let entries;
  if (info.kind === 'allDrives') {
    const options = driveSearchDrivesOptions_();
    if (pageToken) options.pageToken = String(pageToken);
    page = Drive.Drives.list(options);
    entries = (page.drives || []).map(drive => ({ id: drive.id, name: drive.name, kind: 'drive' }));
  } else {
    const options = driveSearchListOptions_(info.driveId);
    options.q += " and '" + escapeDriveQuery_(info.id) + "' in parents"
      + " and mimeType = 'application/vnd.google-apps.folder'";
    options.pageSize = 100;
    options.orderBy = 'name_natural';
    options.fields = 'nextPageToken,incompleteSearch,files(id,name)';
    if (pageToken) options.pageToken = String(pageToken);
    page = Drive.Files.list(options);
    if (page.incompleteSearch) throw new Error('フォルダ一覧を取得しきれませんでした。もう一度開いてください。');
    entries = (page.files || []).map(file => ({ id: file.id, name: file.name, kind: 'folder' }));
  }
  return {
    rootId: DRIVE_SEARCH.ALL_DRIVES_ID, rootName: '共有ドライブ一覧',
    folderId: info.id, folderName: info.name, kind: info.kind, driveName: info.driveName,
    path: info.path, breadcrumbs: info.breadcrumbs, folders: entries,
    nextPageToken: page.nextPageToken || '', warning: warning
  };
}

function selectSearchFolder(folderId) {
  const lock = acquireDriveSearchLock_();
  try {
    requireDriveSearchApi_();
    const sheet = getDriveSearchSheet_(SpreadsheetApp.getActiveSpreadsheet());
    const info = getDriveSearchFolderInfo_(String(folderId || ''));
    writeDriveSearchSelection_(sheet, info);
    refreshDriveSearchStatusLocked_(sheet, PropertiesService.getDocumentProperties());
    return { id: info.id, path: info.path };
  } finally {
    lock.releaseLock();
  }
}

function requireDriveSearchApi_() {
  if (typeof Drive === 'undefined') throw new Error('「サービス +」でGoogle Drive API（v3）を追加してください。');
}

function getDriveSearchVirtualRoot_() {
  return {
    id: DRIVE_SEARCH.ALL_DRIVES_ID, name: '共有ドライブ一覧', kind: 'allDrives',
    driveId: '', driveName: '', path: '//',
    breadcrumbs: [{ id: DRIVE_SEARCH.ALL_DRIVES_ID, name: '共有ドライブ一覧' }]
  };
}

function getDriveSearchDriveInfo_(id) {
  if (!isDriveSearchId_(id) || id === DRIVE_SEARCH.ALL_DRIVES_ID) throw new Error('共有ドライブIDが正しくありません。');
  const drive = Drive.Drives.get(id, { fields: 'id,name' });
  return {
    id: drive.id, name: drive.name, kind: 'drive', driveId: drive.id, driveName: drive.name,
    path: '//' + drive.name,
    breadcrumbs: getDriveSearchVirtualRoot_().breadcrumbs.concat([{ id: drive.id, name: drive.name }])
  };
}

function getDriveSearchDefaultRoot_() {
  const id = String(DRIVE_SEARCH.DRIVE_ID).trim();
  if (!isDriveSearchId_(id)) throw new Error('DRIVE_IDに起点となる共有ドライブIDを設定するか、選択画面から場所を選んでください。');
  return getDriveSearchDriveInfo_(id);
}

// 選択画面で確定したIDを優先し、同名・末尾空白を保護。v2の選択情報は初回だけ利用。
function resolveDriveSearchScope_(sheet) {
  const rawInput = sheet.getRange('B2').getDisplayValue();
  const props = PropertiesService.getDocumentProperties();
  const currentSelection = props.getProperty(DRIVE_SEARCH.SELECTION_KEY);
  const saved = currentSelection || props.getProperty(DRIVE_SEARCH.LEGACY_SELECTION_KEY);
  if (saved) {
    const selection = JSON.parse(saved);
    if (selection.sheetId === sheet.getSheetId() &&
        (currentSelection || selection.driveId === String(DRIVE_SEARCH.DRIVE_ID).trim()) &&
        (selection.pathHash ? selection.pathHash === driveSearchHash_(rawInput) : selection.path === rawInput)) {
      return getDriveSearchFolderInfo_(selection.id);
    }
  }
  const input = rawInput.trim();
  if (/^[A-Za-z]:[\\/]/.test(input)) {
    throw new Error('G:\\などPCのパスは使えません。//ドライブ名/営業/見積、フォルダURL、または選択画面を使用してください。');
  }
  let id = '';
  if (/^id:/i.test(input)) {
    id = input.slice(3).trim();
    if (!isDriveSearchId_(id)) throw new Error('id:の後ろにフォルダIDを入力してください。');
  } else if (/^https?:\/\//i.test(input)) {
    const host = input.match(/^https?:\/\/([^/]+)/i);
    if (!host || host[1].toLowerCase() !== 'drive.google.com') throw new Error('Google DriveのフォルダURLを入力してください。');
    if (/\/drive\/(?:u\/\d+\/)?shared-drives(?:[/?#]|$)/.test(input)) return getDriveSearchVirtualRoot_();
    const match = input.match(/\/folders\/([A-Za-z0-9_-]+)/) || input.match(/[?&]id=([A-Za-z0-9_-]+)/);
    if (!match) throw new Error('共有ドライブのフォルダURL、または共有ドライブ一覧のURLを入力してください。');
    id = match[1];
  } else if (isDriveSearchId_(input) && input.length >= 10) {
    id = input;
  }
  if (id) return getDriveSearchFolderInfo_(id);

  const path = input.replace(/\\/g, '/');
  const segments = path.split('/').filter(part => part !== '');
  let root;
  if (path.startsWith('//')) {
    if (!segments.length) return getDriveSearchVirtualRoot_();
    root = findDriveSearchDriveByName_(segments.shift());
  } else {
    root = getDriveSearchDefaultRoot_();
  }
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
  return getDriveSearchFolderInfo_(currentId);
}

function findDriveSearchDriveByName_(name) {
  const options = driveSearchDrivesOptions_();
  const matches = new Map();
  do {
    const page = Drive.Drives.list(options);
    (page.drives || []).filter(drive => drive.name === name).forEach(drive => matches.set(drive.id, drive));
    if (matches.size > 1) throw new Error('「' + name + '」という同名の共有ドライブが複数あります。選択画面またはURLで指定してください。');
    options.pageToken = page.nextPageToken || '';
  } while (options.pageToken);
  if (!matches.size) throw new Error('共有ドライブ「' + name + '」が見つかりません。');
  return getDriveSearchDriveInfo_(matches.values().next().value.id);
}

// Files.getで実体を確認。共有ドライブのルートはDrives.getでも確認できる。
function getDriveSearchFolderInfo_(folderId) {
  if (folderId === DRIVE_SEARCH.ALL_DRIVES_ID) return getDriveSearchVirtualRoot_();
  if (!isDriveSearchId_(folderId)) throw new Error('フォルダIDが正しくありません。');
  const fields = { supportsAllDrives: true, fields: 'id,name,mimeType,trashed,parents,driveId' };
  let file;
  try { file = Drive.Files.get(folderId, fields); }
  catch (error) {
    try { return getDriveSearchDriveInfo_(folderId); } catch (unused) { throw error; }
  }
  if (file.mimeType !== 'application/vnd.google-apps.folder' || file.trashed) {
    throw new Error('ゴミ箱以外の通常のフォルダを指定してください。');
  }
  if (!file.driveId) {
    try { return getDriveSearchDriveInfo_(folderId); }
    catch (unused) { throw new Error('マイドライブ等は対象外です。共有ドライブ内のフォルダを指定してください。'); }
  }
  const root = getDriveSearchDriveInfo_(file.driveId);
  if (folderId === root.id) return root;
  const chain = [];
  const seen = new Set();
  while (file.id !== root.id) {
    if (seen.has(file.id) || chain.length >= 500) throw new Error('フォルダ階層を確認できません。フォルダURLを確認してください。');
    seen.add(file.id);
    if (file.mimeType !== 'application/vnd.google-apps.folder' || file.trashed || file.driveId !== root.id) {
      throw new Error('共有ドライブ内の有効なフォルダ階層を確認できません。');
    }
    chain.push({ id: file.id, name: file.name });
    if (!file.parents || !file.parents.length) throw new Error('共有ドライブまでの親フォルダを確認できません。');
    if (file.parents[0] === root.id) break;
    file = Drive.Files.get(file.parents[0], fields);
  }
  chain.reverse();
  return {
    id: folderId, name: chain[chain.length - 1].name, kind: 'folder', driveId: root.id, driveName: root.name,
    path: root.path + '/' + chain.map(folder => folder.name).join('/'),
    breadcrumbs: root.breadcrumbs.concat(chain)
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

function driveSearchDrivesOptions_() {
  return { pageSize: 100, fields: 'nextPageToken,drives(id,name)' };
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
    throw new Error('先にsetupDriveSearchを実行し、v3のシート配置へ更新してください。');
  }
  return sheet;
}

function acquireDriveSearchLock_() {
  const lock = LockService.getDocumentLock();
  if (!lock || !lock.tryLock(1000)) throw new Error('別の処理が実行中です。終了後に実行してください。');
  return lock;
}

function saveDriveSearchState_(props, state) {
  const updatedAt = Date.now();
  setDriveSearchProperty_(props, DRIVE_SEARCH.STATE_KEY, Object.assign({}, state, { updatedAt: updatedAt }));
  state.updatedAt = updatedAt;
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

// E1/F1に別のデータがある場合は、結合で消さず利用者に移動を案内する。
function setupDriveSearchStatusArea_(sheet) {
  const area = sheet.getRange('D1:F1');
  if (area.isPartOfMerge()) {
    const merges = area.getMergedRanges();
    if (merges.length !== 1 || merges[0].getA1Notation() !== 'D1:F1') {
      throw new Error('D1:F1に別のセル結合があります。結合を解除してから初期設定してください。');
    }
  } else {
    if (!sheet.getRange('E1:F1').isBlank()) {
      throw new Error('E1/F1にデータがあります。別のセルへ移動してからsetupDriveSearchを実行してください。');
    }
    area.merge();
  }
  area.setWrap(true).setFontWeight('bold').setFontSize(11)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sheet.setRowHeight(1, 112);
  sheet.getRange('A1:C1').setVerticalAlignment('middle');
}

// この機能はDrive APIを呼ばず、保存済みの進捗と実行ロックだけを確認する。
function refreshDriveSearchStatus() { return refreshDriveSearchStatus_(false); }

function refreshDriveSearchStatus_(quiet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DRIVE_SEARCH.SHEET_NAME);
  if (!sheet) {
    if (!quiet) throw new Error('先にsetupDriveSearchを実行してください。');
    return;
  }
  const lock = LockService.getDocumentLock();
  if (!lock || !lock.tryLock(1000)) {
    if (!quiet) ss.toast('検索または別の操作が実行中です。状態は処理の区切りで更新されます。', '共有ドライブ検索', 8);
    return;
  }
  try {
    // 旧版の別レイアウトへD1を書き込まない。
    getDriveSearchSheet_(ss);
    refreshDriveSearchStatusLocked_(sheet, PropertiesService.getDocumentProperties());
  } catch (error) {
    if (!quiet) throw error;
  } finally {
    lock.releaseLock();
  }
}

// 呼び出し元がDocumentLockを保持する。進捗のtoken/行番号/キュー位置は変更しない。
function refreshDriveSearchStatusLocked_(sheet, props) {
  withDriveSearchControlLock_(() => {
    const saved = props.getProperty(DRIVE_SEARCH.STATE_KEY);
    const state = saved ? JSON.parse(saved) : null;
    if (!state || state.version !== 3 || state.sheetId !== sheet.getSheetId()) {
      showDriveSearchStatus_(sheet, null);
    } else {
      const request = readDriveSearchControl_(props, state);
      if (state.lastAttempt && state.lastAttempt.kind === 'preparing') {
        state.lastAttempt.kind = 'interrupted';
        state.lastAttempt.message = '新しい検索の準備中に実行が中断しました。「新しく検索」を実行してください。';
        saveDriveSearchState_(props, state);
      }
      if (!state.stopped && request && request.action === 'stop') {
        state.stopped = true;
        state.status = 'stopped';
        saveDriveSearchState_(props, state);
      } else if (state.status === 'running') {
        // 強制終了ではcatchが走らない。ロック取得成功により実行がないことを確認済み。
        // 最終ページが保存済みでも、完了の確定前なら「続きから検索」で最終確認する。
        state.status = request && request.action === 'pause' ? 'paused' : 'interrupted';
        if (state.status === 'paused') state.pauseReason = 'manual';
        saveDriveSearchState_(props, state);
      }
      showDriveSearchStatus_(sheet, state);
    }
    SpreadsheetApp.flush();
  });
}

function driveSearchInputsMatch_(sheet, state) {
  if (state.sheetId !== sheet.getSheetId() ||
      state.defaultDriveId !== String(DRIVE_SEARCH.DRIVE_ID).trim() ||
      state.keywordHash !== driveSearchHash_(sheet.getRange('B1').getDisplayValue().trim()) ||
      state.scopeInputHash !== driveSearchHash_(sheet.getRange('B2').getDisplayValue())) return false;
  const saved = PropertiesService.getDocumentProperties().getProperty(DRIVE_SEARCH.SELECTION_KEY);
  if (saved) {
    const selection = JSON.parse(saved);
    // 同じパス名でも、選択画面から別IDの同名フォルダを選んだ場合を区別する。
    if (selection.sheetId === state.sheetId && selection.pathHash === state.scopeInputHash &&
        selection.id !== state.scopeId) return false;
  }
  return true;
}

function markDriveSearchPreparing_(sheet, props) {
  withDriveSearchControlLock_(() => {
    const saved = props.getProperty(DRIVE_SEARCH.STATE_KEY);
    const previous = saved ? JSON.parse(saved) : null;
    if (!previous || previous.version !== 3 || previous.sheetId !== sheet.getSheetId()) return;
    // 準備中の強制終了後に、前回の緑の完了表示が戻ることを防ぐ。
    previous.lastAttempt = {
      kind: 'preparing', keywordHash: driveSearchHash_(sheet.getRange('B1').getDisplayValue().trim()),
      scopeInputHash: driveSearchHash_(sheet.getRange('B2').getDisplayValue()),
      message: '', at: Date.now()
    };
    saveDriveSearchState_(props, previous);
  });
}

function reportDriveSearchError_(sheet, props, localState, processing, error) {
  const message = String(error && error.message || error).slice(0, 500);
  let display = { status: 'notStarted', checked: 0, hits: 0, lastError: message };
  try {
    withDriveSearchControlLock_(() => {
      try {
        const saved = props.getProperty(DRIVE_SEARCH.STATE_KEY);
        const checkpoint = saved ? JSON.parse(saved) : null;
        if (processing && checkpoint && localState && checkpoint.jobId === localState.jobId) {
          const request = readDriveSearchControl_(props, checkpoint);
          checkpoint.status = checkpoint.stopped || (request && request.action === 'stop') ? 'stopped' : 'error';
          if (checkpoint.status === 'stopped') checkpoint.stopped = true;
          checkpoint.lastError = message;
          delete checkpoint.lastAttempt;
          display = checkpoint;
          saveDriveSearchState_(props, checkpoint);
        } else if (checkpoint && checkpoint.sheetId === sheet.getSheetId()) {
          // 開始前の失敗は前回の検索カーソル・完了フラグを壊さず、別の通知情報として保持。
          display = Object.assign({}, checkpoint, { status: 'notStarted', lastError: message, lastAttempt: null });
          if (checkpoint.owner === Session.getTemporaryActiveUserKey()) {
            checkpoint.lastAttempt = {
              keywordHash: driveSearchHash_(sheet.getRange('B1').getDisplayValue().trim()),
              scopeInputHash: driveSearchHash_(sheet.getRange('B2').getDisplayValue()),
              message: message, at: Date.now()
            };
            display.lastAttempt = checkpoint.lastAttempt;
            saveDriveSearchState_(props, checkpoint);
          }
        }
      } finally {
        // 保存と表示を同じロック内で行い、後着の終了要求を古いエラー表示で上書きしない。
        try {
          showDriveSearchStatus_(sheet, display);
          SpreadsheetApp.flush();
        } catch (displayError) {
          // Sheets自体が利用不可の場合は元の例外をApps Script側へ返す。
        }
      }
    });
  } catch (reportError) {
    // 進捗保存自体の失敗でも、元のエラーを別の例外で隠さない。
  }
}

function showDriveSearchStatus_(sheet, state, detail) {
  state = state || { status: 'idle', checked: 0, hits: 0 };
  let status = state.stopped ? 'stopped' : state.status;
  let errorText = state.lastError || '';
  const attempt = state.lastAttempt;
  const controlStatus = ['stopped', 'stop_requested', 'pause_requested'].includes(status);
  if (!controlStatus && attempt && attempt.keywordHash === driveSearchHash_(sheet.getRange('B1').getDisplayValue().trim()) &&
      attempt.scopeInputHash === driveSearchHash_(sheet.getRange('B2').getDisplayValue())) {
    status = attempt.kind === 'preparing' ? 'preparing' : 'notStarted';
    errorText = attempt.message;
  } else if (!controlStatus && state.jobId && !['preparing', 'notStarted'].includes(status) && !driveSearchInputsMatch_(sheet, state)) {
    status = 'changed';
  } else if (status === 'running') {
    const request = readDriveSearchControl_(PropertiesService.getDocumentProperties(), state);
    if (request) status = request.action === 'stop' ? 'stop_requested' : 'pause_requested';
  }
  const complete = status === 'done' && state.done === true && state.stopped !== true;
  const views = {
    idle: ['【未検索】検索を開始してください', '#eeeeee', 'B1に検索語、B2に場所を指定して「新しく検索」を実行してください。'],
    preparing: ['【未完了】検索の準備中', '#d9eaf7', '表示中の行は前回の結果です。新しい検索はまだ完了していません。'],
    running: ['【未完了】検索中 — 表示中の一覧は途中結果です', '#d9eaf7', '緑の「検索完了」が表示されるまで、検索は終わっていません。'],
    paused: [state.pauseReason === 'time' ? '【未完了】自動一時停止（時間の区切り）' : '【未完了】一時停止中', '#fff2cc', 'まだ続きがあります。「続きから検索」を押してください。'],
    pause_requested: ['【未完了】一時停止を要求中', '#fff2cc', '現在のページ処理後に停止します。「一時停止中」の表示を待ってください。'],
    stop_requested: ['【未完了】終了を要求中', '#f4cccc', '現在のページ処理後に終了します。まだ停止は確定していません。'],
    stopped: ['【未完了】検索を終了しました（再開不可）', '#f4cccc', '表示中は途中結果です。必要な場合は「新しく検索」でやり直してください。'],
    error: ['【未完了】エラーで中断', '#f4cccc', '原因を確認して「続きから検索」。再開不可の案内がある場合は「新しく検索」。'],
    interrupted: ['【未完了】前回の実行が中断しています', '#f4cccc', '完了を確認できていません。「続きから検索」で再開・最終確認してください。'],
    changed: ['【条件変更・未検索】表示中は前回の条件での結果です', '#fff2cc', '現在の条件では検索していません。「新しく検索」を押してください。'],
    notStarted: ['【未完了】今回の操作を実行できませんでした', '#f4cccc', '表示中の行は前回の結果です。原因を修正して、もう一度操作してください。']
  };
  const view = complete
    ? ['【検索完了】選択した範囲の検索が最後まで終了しました', '#d9ead3', 'この検索の全ページ・全対象範囲を確認済みです。']
    : (views[status] || views.interrupted);
  const hits = Number(state.hits || 0).toLocaleString('ja-JP');
  const checked = Number(state.checked || 0).toLocaleString('ja-JP');
  let counts = (complete ? '検索結果 ' : '保存済みの途中結果 ') + hits + '件 ／ 確認済み ' + checked + '件';
  if (['idle', 'preparing', 'changed', 'notStarted'].includes(status)) {
    counts = status === 'idle' ? '保存された検索完了の記録はありません。' : '表示中の行を、今回の条件での検索完了結果として使わないでください。';
  }
  let guidance = detail && status === 'running' ? detail : view[2];
  if (errorText && ['error', 'notStarted'].includes(status)) guidance += ' 原因：' + errorText.slice(0, 90);
  const time = complete ? '完了日時：' + formatDriveSearchStatusTime_(state.finishedAt)
    : '状態の保存日時：' + formatDriveSearchStatusTime_(attempt && status === 'notStarted' ? attempt.at : state.updatedAt);
  sheet.getRange('D1:F1').setBackground(view[1]).setFontColor('#202124')
    .setFontWeight('bold').setFontSize(11).setWrap(true).setVerticalAlignment('middle');
  sheet.getRange('D1').setValue([view[0], counts, guidance, time].join('\n'))
    .setNote('緑の「検索完了」の場合だけ、選択した範囲の検索処理が完了しています。\n'
      + '件数はファイルとフォルダの合計です。総数が未確定のため進捗率は表示しません。\n'
      + 'エラー直前の未保存行が残る場合は、再開時に最後の保存位置から整理されます。'
      + (errorText ? '\nエラー詳細：' + errorText : ''));
}

function formatDriveSearchStatusTime_(value) {
  if (typeof value !== 'number') return '記録なし（未検索または旧版の記録）';
  const timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(new Date(value), timezone, 'yyyy/MM/dd HH:mm:ss');
}
