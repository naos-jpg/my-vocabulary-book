/**
 * my単語帳 — Google Apps Script v4（JSONP対応）
 *
 * 【デプロイ設定】
 * ・種類: ウェブアプリ
 * ・次のユーザーとして実行: 「自分（○○@gmail.com）」
 * ・アクセスできるユーザー: 「全員」
 *
 * ★ 自分のGoogleアカウントでデプロイ → 自分のドライブにシート作成
 * ★ 方法3: 各自がGASを登録して自分のスプレッドシートを使う
 */

const SHEET_NAME = 'my単語帳';
const HEADER     = ['英単語','日本語訳','品詞','IPA','英英定義','例文',
                    'CEFR','セクション','登録日','習得'];

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
  return ss.getSheetByName('単語帳') || ss.getActiveSheet();
}

// ── GET ────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'load';

    if (action === 'ping') {
      const sh  = getSheet();
      const url = sh.getParent().getUrl();
      return makeResp(e, { sheetUrl: url });
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
          star:    col.star >= 0 ? (r[col.star] === '1' || r[col.star] === true) : false,
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
        sh2.appendRow([
          p.word || '', p.meaning || '', p.pos || '', p.ipa || '',
          p.def  || '', p.ex || '', p.cefr || 'B1', p.section || 'My Words',
          p.addedAt || new Date().toLocaleDateString('ja-JP'),
          (p.star === '1') ? '1' : '0'
        ]);
        return makeResp(e, { added: 1, row: sh2.getLastRow() });
      }
      const c = JSON.parse(p.cards || p.card || '{}');
      const cards = Array.isArray(c) ? c : [c];
      cards.forEach(card => {
        sh2.appendRow([
          card.word || '', card.meaning || '', card.pos || '', card.ipa || '',
          card.def  || '', card.ex || '', card.cefr || 'B1', card.section || 'My Words',
          card.addedAt || new Date().toLocaleDateString('ja-JP'),
          card.star ? '1' : '0'
        ]);
      });
      return makeResp(e, { added: cards.length, row: sh2.getLastRow() });
    }

    if (action === 'addBulk') {
      const cards = JSON.parse(e.parameter.cards || '[]');
      const rows = cards.map(c => [
        c.word || '', c.meaning || '', c.pos || '', c.ipa || '',
        c.def  || '', c.ex || '', c.cefr || 'B1', c.section || 'My Words',
        c.addedAt || new Date().toLocaleDateString('ja-JP'),
        c.star ? '1' : '0'
      ]);
      if (rows.length) {
        sh2.getRange(sh2.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
      }
      return makeResp(e, { added: rows.length });
    }

    if (action === 'star') {
      const row  = parseInt(e.parameter.row);
      const star = e.parameter.star === 'true' || e.parameter.star === '1';
      if (row >= 2) sh2.getRange(row, 10).setValue(star ? '1' : '0');
      return makeResp(e, { updated: row });
    }

    if (action === 'delete') {
      const row = parseInt(e.parameter.row);
      if (row >= 2) sh2.deleteRow(row);
      return makeResp(e, { deleted: row });
    }

    if (action === 'update') {
      const row = parseInt(e.parameter.row);
      if (row >= 2) {
        const c = JSON.parse(e.parameter.cards || '{}');
        sh2.getRange(row, 1, 1, 10).setValues([[
          c.word || '', c.meaning || '', c.pos || '', c.ipa || '',
          c.def  || '', c.ex || '', c.cefr || 'B1', c.section || 'My Words',
          c.addedAt || '', c.star ? '1' : '0'
        ]]);
      }
      return makeResp(e, { updated: row });
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
      sh.appendRow([
        c.word || '', c.meaning || '', c.pos || '', c.ipa || '',
        c.def  || '', c.ex || '', c.cefr || 'B1', c.section || 'My Words',
        c.addedAt || new Date().toLocaleDateString('ja-JP'),
        c.star ? '1' : '0'
      ]);
      return makeResp(e, { row: sh.getLastRow() });
    }

    if (action === 'star') {
      const { row, star } = body;
      if (row >= 2) sh.getRange(row, 10).setValue(star ? '1' : '0');
      return makeResp(e, { updated: row });
    }

    if (action === 'delete') {
      const { row } = body;
      if (row >= 2) sh.deleteRow(row);
      return makeResp(e, { deleted: row });
    }

    return makeResp(e, { error: 'Unknown action' });
  } catch (err) {
    return makeResp(e, { error: err.toString() });
  }
}
