/**
 * my単語帳 — Google Apps Script v6（JSONP対応）
 *
 * v5 の追加機能:
 *  ・学習記録（復習スケジュール・学習履歴など）の保存と読み込み
 *    → シート「学習記録」に保存（端末をまたいで記録を共有）
 *  ・単語の全削除（バックアップを残す／完全に削除）
 * v6 の追加機能:
 *  ・「単元」列（既存のシートにも自動で追加）
 *  ・列の並び順が違うシートでも、見出しの名前で正しい列に書き込む
 *  ・★・編集・削除は、行番号の単語が違う場合は単語名で行を探し直す（別の端末で行がずれても安全）
 *
 * 【デプロイ設定】
 * ・種類: ウェブアプリ
 * ・次のユーザーとして実行: 「自分（○○@gmail.com）」
 * ・アクセスできるユーザー: 「全員」
 *
 * ★ 自分のGoogleアカウントでデプロイ → 自分のドライブにシート作成
 * ★ 方法3: 各自がGASを登録して自分のスプレッドシートを使う
 */

const GAS_VERSION = 6;
const SHEET_NAME = 'my単語帳';
const PROGRESS_SHEET = '学習記録';
const PROGRESS_TMP   = '学習記録_送信中';
const CELL_MAX = 40000;   // 1セルに入れる最大文字数
const HEADER     = ['英単語','日本語訳','品詞','IPA','英英定義','例文',
                    'CEFR','セクション','登録日','習得','単元'];
// カードの項目 → 見出し名（別名も可）
const FIELD_COLS = {
  word:['英単語'], meaning:['日本語訳','品詞・意味'], pos:['品詞'], ipa:['IPA','発音記号 (IPA)','発音記号'],
  def:['英英定義'], ex:['例文'], cefr:['CEFR','CEFRレベル'], section:['セクション'],
  addedAt:['登録日','日時'], star:['習得','習得済み','SRSレベル'], unit:['単元']
};

// ── JSON / JSONP レスポンス生成 ─────────────────────────
function makeResp(e, data) {
  const json     = JSON.stringify({ status:'ok', ...data });
  const callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ── スプレッドシート取得 or 作成 ────────────────────────
function getSheet() {
  const files = DriveApp.getFilesByName(SHEET_NAME);
  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME);
    const sh = ss.getActiveSheet();
    sh.setName('単語帳');
    sh.appendRow(HEADER);
    sh.getRange(1, 1, 1, HEADER.length)
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 130);
    sh.setColumnWidth(2, 180);
    sh.setColumnWidth(5, 250);
    sh.setColumnWidth(6, 250);
  }
  const sh = ss.getSheetByName('単語帳') || ss.getActiveSheet();
  ensureUnitColumn(sh);
  return sh;
}

// 既存のシートに「単元」列が無ければ追加
function ensureUnitColumn(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  if (head.indexOf('単元') < 0) {
    sh.getRange(1, lastCol + 1).setValue('単元')
      .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
  }
}
// 見出し名 → 列番号（0始まり）
function colMap(sh) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const m = {};
  for (const [k, names] of Object.entries(FIELD_COLS)) {
    for (const n of names) { const i = head.indexOf(n); if (i >= 0) { m[k] = i; break; } }
  }
  m._width = head.length;
  return m;
}
// 行番号の単語が match と違えば、単語名で行を探し直す（見つからなければ -1）
function resolveRow(sh, row, match) {
  if (!match) return row;
  const m = colMap(sh); if (m.word === undefined) return row;
  const last = sh.getLastRow(); if (last < 2) return -1;
  if (row >= 2 && row <= last && String(sh.getRange(row, m.word + 1).getValue()) === match) return row;
  const words = sh.getRange(2, m.word + 1, last - 1, 1).getValues();
  for (let i = 0; i < words.length; i++) if (String(words[i][0]) === match) return i + 2;
  return -1;
}
// カード → 1行分の配列（見出しの位置に合わせる）
function rowFromCard(m, c) {
  const row = new Array(m._width).fill('');
  const v = {
    word: c.word || '', meaning: c.meaning || '', pos: c.pos || '', ipa: c.ipa || '',
    def: c.def || '', ex: c.ex || '', cefr: c.cefr || 'B1', section: c.section || 'My Words',
    addedAt: c.addedAt || new Date().toLocaleDateString('ja-JP'),
    star: (c.star === true || c.star === '1' || c.star === 'true') ? '1' : '0', unit: c.unit || ''
  };
  for (const [k, i] of Object.entries(m)) if (k !== '_width' && v[k] !== undefined) row[i] = v[k];
  return row;
}

// ── GET ────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'load';

    if (action === 'ping') {
      const sh  = getSheet();
      const url = sh.getParent().getUrl();
      return makeResp(e, { sheetUrl: url, version: GAS_VERSION });
    }

    // ── 学習記録: 読み込み ──
    if (action === 'loadProgress') {
      const ss = getSheet().getParent();
      const ps = ss.getSheetByName(PROGRESS_SHEET);
      if (!ps || ps.getLastRow() < 1) return makeResp(e, { progress: null });
      const vals = ps.getRange(1, 1, ps.getLastRow(), 1).getValues();
      const json = vals.map(r => String(r[0] || '')).join('');
      return makeResp(e, { progress: json || null });
    }

    // ── 学習記録: 分割して受け取り、全部そろったら保存 ──
    if (action === 'saveProgress') {
      const p = e.parameter;
      const id = p.sid, part = parseInt(p.part), total = parseInt(p.total);
      const ss = getSheet().getParent();
      let tmp = ss.getSheetByName(PROGRESS_TMP);
      if (!tmp) { tmp = ss.insertSheet(PROGRESS_TMP); tmp.hideSheet(); }
      // 別の送信の残りは消す
      if (part === 0) tmp.clear();
      tmp.getRange(part + 1, 1, 1, 2).setValues([[id, p.data || '']]);
      // 全パートそろったか確認
      const rows = tmp.getRange(1, 1, total, 2).getValues();
      const done = rows.every(r => r[0] === id);
      if (!done) return makeResp(e, { saved: false, received: part });
      const json = rows.map(r => String(r[1])).join('');
      let ps = ss.getSheetByName(PROGRESS_SHEET);
      if (!ps) ps = ss.insertSheet(PROGRESS_SHEET);
      ps.clear();
      const cells = [];
      for (let i = 0; i < json.length; i += CELL_MAX) cells.push([json.slice(i, i + CELL_MAX)]);
      if (cells.length) ps.getRange(1, 1, cells.length, 1).setValues(cells);
      tmp.clear();
      return makeResp(e, { saved: true, length: json.length });
    }

    if (action === 'load') {
      const sh     = getSheet();
      const values = sh.getDataRange().getValues();
      const cards  = [];
      if (values.length < 2) return makeResp(e, { cards: [], sheetUrl: sh.getParent().getUrl() });

      const header = values[0].map(h => String(h).trim());
      const col = {
        word:    Math.max(header.indexOf('英単語'), 0),
        meaning: header.indexOf('日本語訳') >= 0 ? header.indexOf('日本語訳') : header.indexOf('品詞・意味'),
        pos:     header.indexOf('品詞') >= 0 ? header.indexOf('品詞') : -1,
        ipa:     header.indexOf('IPA'),
        def:     header.indexOf('英英定義'),
        ex:      header.indexOf('例文'),
        cefr:    header.indexOf('CEFR'),
        section: header.indexOf('セクション'),
        addedAt: header.indexOf('登録日') >= 0 ? header.indexOf('登録日') : header.indexOf('日時'),
        star:    header.indexOf('習得') >= 0 ? header.indexOf('習得') : header.indexOf('SRSレベル'),
        unit:    header.indexOf('単元'),
      };

      for (let i = 1; i < values.length; i++) {
        const r = values[i];
        const word = col.word >= 0 ? (r[col.word] || '') : (r[2] || r[0] || '');
        if (!word) continue;
        cards.push({
          word:    word,
          meaning: col.meaning >= 0 ? (r[col.meaning] || '') : (r[1] || ''),
          pos:     col.pos >= 0 ? (r[col.pos] || '') : '',
          ipa:     col.ipa >= 0 ? (r[col.ipa] || '') : '',
          def:     col.def >= 0 ? (r[col.def] || '') : '',
          ex:      col.ex >= 0 ? (r[col.ex] || '') : '',
          cefr:    col.cefr >= 0 ? (r[col.cefr] || 'B1') : 'B1',
          section: col.section >= 0 ? (r[col.section] || 'My Words') : 'My Words',
          addedAt: col.addedAt >= 0 ? (r[col.addedAt] || '') : '',
          star:    col.star >= 0 ? (r[col.star] === '1' || r[col.star] === 1 || r[col.star] === true) : false,
          unit:    col.unit >= 0 ? String(r[col.unit] || '') : '',
          row:     i + 1
        });
      }
      const url = sh.getParent().getUrl();
      return makeResp(e, { cards, sheetUrl: url });
    }

    // ── 書き込み操作（GETパラメータでJSONP対応）──
    const sh2 = getSheet();

    if (action === 'add') {
      const p = e.parameter;
      if (p.word) {
        sh2.appendRow(rowFromCard(colMap(sh2), p));
        return makeResp(e, { added: 1, row: sh2.getLastRow() });
      }
      const c = JSON.parse(p.cards || p.card || '{}');
      const cards = Array.isArray(c) ? c : [c];
      const m0 = colMap(sh2);
      cards.forEach(card => sh2.appendRow(rowFromCard(m0, card)));
      return makeResp(e, { added: cards.length, row: sh2.getLastRow() });
    }

    if (action === 'addBulk') {
      const cards = JSON.parse(e.parameter.cards || '[]');
      const m1 = colMap(sh2);
      const rows = cards.map(c => rowFromCard(m1, c));
      if (rows.length) {
        sh2.getRange(sh2.getLastRow() + 1, 1, rows.length, m1._width).setValues(rows);
      }
      return makeResp(e, { added: rows.length });
    }

    if (action === 'star') {
      const row  = resolveRow(sh2, parseInt(e.parameter.row), e.parameter.match);
      const star = e.parameter.star === 'true' || e.parameter.star === '1';
      const m2 = colMap(sh2);
      if (row < 2) return makeResp(e, { missing: true });
      if (m2.star !== undefined) sh2.getRange(row, m2.star + 1).setValue(star ? '1' : '0');
      return makeResp(e, { updated: row, row: row });
    }

    if (action === 'delete') {
      const row = resolveRow(sh2, parseInt(e.parameter.row), e.parameter.match);
      if (row < 2) return makeResp(e, { missing: true });
      sh2.deleteRow(row);
      return makeResp(e, { deleted: row });
    }

    // ── 全削除: バックアップシートに移してから空にする ──
    if (action === 'archiveAll') {
      const ss = sh2.getParent();
      const last = sh2.getLastRow();
      if (last < 2) return makeResp(e, { archived: 0 });
      const name = 'バックアップ_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
      const bk = sh2.copyTo(ss).setName(name);
      sh2.deleteRows(2, last - 1);
      return makeResp(e, { archived: last - 1, backup: bk.getName() });
    }

    // ── 全削除: 完全に削除 ──
    if (action === 'clearAll') {
      const last = sh2.getLastRow();
      if (last >= 2) sh2.deleteRows(2, last - 1);
      return makeResp(e, { cleared: Math.max(0, last - 1) });
    }

    if (action === 'update') {
      const row = resolveRow(sh2, parseInt(e.parameter.row), e.parameter.match);
      if (row < 2) return makeResp(e, { missing: true });
      {
        const c = JSON.parse(e.parameter.cards || '{}');
        const m3 = colMap(sh2);
        if (!c.addedAt && m3.addedAt !== undefined) c.addedAt = sh2.getRange(row, m3.addedAt + 1).getValue();   // 登録日は残す
        sh2.getRange(row, 1, 1, m3._width).setValues([rowFromCard(m3, c)]);
      }
      return makeResp(e, { updated: row, row: row });
    }

    return makeResp(e, { error: 'Unknown action' });
  } catch (err) {
    return makeResp(e, { error: err.toString() });
  }
}

// ── POST（後方互換） ────────────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    const sh     = getSheet();

    if (action === 'add') {
      const c = body.card;
      sh.appendRow(rowFromCard(colMap(sh), c));
      return makeResp(e, { row: sh.getLastRow() });
    }

    if (action === 'star') {
      const row = resolveRow(sh, parseInt(body.row), body.match);
      const star = body.star;
      const m = colMap(sh);
      if (row < 2) return makeResp(e, { missing: true });
      if (m.star !== undefined) sh.getRange(row, m.star + 1).setValue(star ? '1' : '0');
      return makeResp(e, { updated: row, row: row });
    }

    if (action === 'delete') {
      const row = resolveRow(sh, parseInt(body.row), body.match);
      if (row < 2) return makeResp(e, { missing: true });
      sh.deleteRow(row);
      return makeResp(e, { deleted: row });
    }

    return makeResp(e, { error: 'Unknown action' });
  } catch (err) {
    return makeResp(e, { error: err.toString() });
  }
}
