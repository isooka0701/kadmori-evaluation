const SPREADSHEET_ID = '1iBeN-ZHZ38tYduTZWzBaGBkndMrIbWUP2DzNO8vIOY4';
const SELF_EVAL_SPREADSHEET_ID = '1oxHwBseRXE02_zbWtdq5DKQVqE8hz69AxdR4U4l_gvM';
const LINE_ID_SHEET_NAME = 'LINE_ID';
const CLINIC_NAMES = ['大阪院', '東京院'];

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getStaffList') return getStaffList(e.parameter.clinic);
  if (action === 'getStaffDetail') return getStaffDetail(e.parameter.clinic, e.parameter.staff);
  if (action === 'getSelfEvaluation') return getSelfEvaluationForUI(e.parameter.clinic, e.parameter.staff);
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('カドモリ 1on1評価システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getStaffList(clinic) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(clinic);
  if (!sheet) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  const data = sheet.getDataRange().getDisplayValues();
  const staffMap = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const staffName = row[2];
    if (!staffName) continue;
    staffMap[staffName] = {
      name: staffName, clinic: row[1], evaluator: row[3], date: row[0],
      scores: { 素直: row[4], 行動スピード: row[5], 振り返り前進: row[6], 継続挑戦: row[7], キッカケづくり: row[8], 凡事徹底: row[9] },
      totalRank: row[16], totalComment: row[17]
    };
  }
  return ContentService.createTextOutput(JSON.stringify(Object.values(staffMap))).setMimeType(ContentService.MimeType.JSON);
}

function getStaffDetail(clinic, staff) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(clinic);
  if (!sheet) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  const data = sheet.getDataRange().getDisplayValues();
  const history = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[2] !== staff) continue;
    history.push({
      date: row[0], clinic: row[1], evaluator: row[3],
      scores: { 素直: row[4], 行動スピード: row[5], 振り返り前進: row[6], 継続挑戦: row[7], キッカケづくり: row[8], 凡事徹底: row[9] },
      categoryRanks: { 素直: row[10], 行動スピード: row[11], 振り返り前進: row[12], 継続挑戦: row[13], キッカケづくり: row[14], 凡事徹底: row[15] },
      totalRank: row[16], totalComment: row[17], memo: row[18], docUrl: row[19],
      aiSummary: row[20] || '', aiGood: row[21] || '', aiImprove: row[22] || '', aiAction: row[23] || ''
    });
  }
  history.reverse();
  return ContentService.createTextOutput(JSON.stringify(history)).setMimeType(ContentService.MimeType.JSON);
}

function saveToGoogleDocs(clinic, staffName, transcript, checklist, memo, evaluator, categoryRanks, categoryReasons, totalRank, totalComment, aiSummary) {
  const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const fileName = clinic + '_' + staffName + '_' + date;
  const rootFolderName = 'カドモリ1on1記録';
  let rootFolder;
  const rootFolders = DriveApp.getFoldersByName(rootFolderName);
  rootFolder = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder(rootFolderName);
  let clinicFolder;
  const clinicFolders = rootFolder.getFoldersByName(clinic);
  clinicFolder = clinicFolders.hasNext() ? clinicFolders.next() : rootFolder.createFolder(clinic);
  let staffFolder;
  const staffFolders = clinicFolder.getFoldersByName(staffName);
  staffFolder = staffFolders.hasNext() ? staffFolders.next() : clinicFolder.createFolder(staffName);
  const doc = DocumentApp.create(fileName);
  const body = doc.getBody();
  body.appendParagraph('■ 面談日: ' + date);
  body.appendParagraph('■ 院名: ' + clinic);
  body.appendParagraph('■ スタッフ名: ' + staffName);
  body.appendParagraph('■ 評価者名: ' + (evaluator || ''));
  body.appendParagraph('');
  body.appendParagraph('【文字起こし】').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(transcript || '');
  body.appendParagraph('');
  body.appendParagraph('【チェックリスト結果】').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(checklist || '');
  body.appendParagraph('');
  body.appendParagraph('【面談メモ】').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(memo || 'なし');
  body.appendParagraph('');
  body.appendParagraph('【カテゴリ別評価】').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  const categoryNames = ['素直','行動スピード','振り返り・前進','継続・挑戦','キッカケづくり','凡事徹底'];
  const ranks = (categoryRanks || '').split(',');
  const reasons = (categoryReasons || '').split('|');
  categoryNames.forEach((name, i) => {
    body.appendParagraph(name + ': ' + (ranks[i] || '') + '　理由: ' + (reasons[i] || ''));
  });
  body.appendParagraph('');
  body.appendParagraph('【総合評価】').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('評価: ' + (totalRank || ''));
  body.appendParagraph('コメント: ' + (totalComment || ''));
  body.appendParagraph('');
  body.appendParagraph('【AI評価】').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(aiSummary || 'なし');
  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  staffFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  const docUrl = doc.getUrl();
  saveToSpreadsheet(clinic, staffName, evaluator, date, checklist, memo, categoryRanks, totalRank, totalComment, docUrl, aiSummary);
  return docUrl;
}

function saveToSpreadsheet(clinic, staffName, evaluator, date, checklist, memo, categoryRanks, totalRank, totalComment, docUrl, aiSummary) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(clinic);
  if (!sheet) {
    sheet = ss.insertSheet(clinic);
    sheet.appendRow(['面談日','院名','スタッフ名','評価者名','素直(点)','行動スピード(点)','振り返り・前進(点)','継続・挑戦(点)','キッカケづくり(点)','凡事徹底(点)','素直(評価)','行動スピード(評価)','振り返り・前進(評価)','継続・挑戦(評価)','キッカケづくり(評価)','凡事徹底(評価)','総合評価','総合コメント','面談メモ','議事録リンク','AI要約','AI良かった点','AI改善ポイント','AI次回アクション']);
  }
  const lines = checklist.split('\n');
  const scores = {};
  lines.forEach(line => {
    const match = line.match(/^(.+?):\s*(\d+\/\d+)/);
    if (match) scores[match[1].trim()] = match[2];
  });
  const ranks = (categoryRanks || '').split(',');
  const aiParts = parseAISummary(aiSummary || '');
  sheet.appendRow([
    date, clinic, staffName, evaluator || '',
    scores['素直'] || '', scores['行動スピード'] || '', scores['振り返り・前進'] || '',
    scores['継続・挑戦'] || '', scores['キッカケづくり'] || '', scores['凡事徹底'] || '',
    ranks[0] || '', ranks[1] || '', ranks[2] || '', ranks[3] || '', ranks[4] || '', ranks[5] || '',
    totalRank || '', totalComment || '', memo || '', docUrl,
    aiParts.summary, aiParts.good, aiParts.improve, aiParts.action
  ]);
}

function parseAISummary(aiText) {
  const result = { summary: '', good: '', improve: '', action: '' };
  if (!aiText) return result;

  const summaryMatch = aiText.match(/■\s*面談要約[^\n]*?[\s]+([\s\S]*?)(?=■|$)/);
  const goodMatch = aiText.match(/■\s*良かった点[^\n]*?[\s]+([\s\S]*?)(?=■|$)/);
  const improveMatch = aiText.match(/■\s*改善ポイント[^\n]*?[\s]+([\s\S]*?)(?=■|$)/);
  const actionMatch = aiText.match(/■\s*次回アクション[^\n]*?[\s]+([\s\S]*?)(?=■|$)/);

  if (summaryMatch) result.summary = summaryMatch[1].trim();
  if (goodMatch) result.good = goodMatch[1].trim();
  if (improveMatch) result.improve = improveMatch[1].trim();
  if (actionMatch) result.action = actionMatch[1].trim();

  return result;
}

function doOptions(e) {
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  if (data.action === 'sendLineImage') {
    return sendLineImage(data.imageBase64, data.staffName);
  }

  if (data.action === 'transcribeAudio') {
    try {
      const text = transcribeAudio(data.audioBase64, data.mimeType);
      return ContentService.createTextOutput(JSON.stringify({ success: true, text: text }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (data.action === 'uploadAudioChunk') {
    try {
      saveAudioChunk_(data.uploadId, data.chunkIndex, data.base64Chunk);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (data.action === 'finishAudioUpload') {
    try {
      const audioBase64 = assembleAudioChunks_(data.uploadId, data.totalChunks);
      const text = transcribeAudio(audioBase64, data.mimeType);
      return ContentService.createTextOutput(JSON.stringify({ success: true, text: text }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // LINEからのWebhookイベント（友だち追加・メッセージ受信）
  if (data.events) {
    try {
      const ev = data.events[0];
      const userId = ev && ev.source && ev.source.userId;

      if (userId && ev.type === 'follow') {
        // 友だち追加時：名前の送信をお願いする
        replyLine(ev.replyToken,
          '友だち追加ありがとうございます。\n最初のメッセージとして、スプレッドシートに登録されているお名前（フルネーム）を送ってください。');

      } else if (userId && ev.type === 'message' && ev.message && ev.message.type === 'text') {
        // メッセージ受信時：名前として照合し、登録する
        registerLineUser(userId, ev.message.text, ev.replyToken);
      }
    } catch (err) {
      Logger.log('webhook error: ' + err.message);
    }
    return ContentService.createTextOutput('ok');
  }

  let transcript = fixHomophones_(data.transcript || '');
  try {
    transcript = cleanTranscript_(transcript);
  } catch (err) {
    // クリーンアップに失敗しても、補正前のテキストのまま処理を続行する
  }

  let selfEvalText = '';
  try {
    const selfEval = getSelfEvaluation(data.staff, data.clinic);
    if (selfEval) {
      selfEvalText = `素直: ${selfEval.素直}\n行動スピード: ${selfEval.行動スピード}\n振り返り・前進: ${selfEval.振り返り前進}\n継続・挑戦: ${selfEval.継続挑戦}\nキッカケづくり: ${selfEval.キッカケづくり}\n凡事徹底: ${selfEval.凡事徹底}\n頑張ったこと: ${selfEval.頑張ったこと}\n課題: ${selfEval.課題}`;
    }
  } catch(err) { selfEvalText = ''; }

  let aiSummary = '';
  try {
    aiSummary = generateAISummary(transcript, data.checklistDetail || '', selfEvalText);
  } catch(err) {
    aiSummary = 'AI評価の生成に失敗しました：' + err.message;
  }

  const url = saveToGoogleDocs(
    data.clinic, data.staff, transcript, data.checklistDetail || '',
    data.memo || '', data.evaluator || '', data.categoryRanks || '',
    data.categoryReasons || '', data.totalRank || '', data.totalComment || '', aiSummary
  );
  return ContentService.createTextOutput(JSON.stringify({url: url}))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 音声認識が同音異義語を誤って書き起こしがちな業務用語を補正する。
 * 例: 「報連相（ほうれんそう）」が野菜の「ほうれん草」として誤認識される。
 */
function fixHomophones_(text) {
  return (text || '').replace(/ほうれん草/g, '報連相');
}

/**
 * 音声認識(ブラウザのWeb Speech API)による文字起こしは誤字脱字や
 * 意味の通らない文が混ざりやすいため、AIで自然な日本語に整える。
 * 内容の追加・推測・要約はせず、表記の誤りのみを直す。
 */
function cleanTranscript_(transcript) {
  if (!transcript || !transcript.trim()) return transcript;
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const prompt = `以下は音声認識による1on1面談の文字起こしです。音声認識特有の誤字脱字・同音異義語の誤変換・意味の通らない文を、自然な日本語になるように直してください。

【厳守事項】
・話されていない内容を推測で追加しないでください
・要約や省略はせず、発言の順序や粒度はそのまま保ってください
・タイムスタンプの行(例: 14:32:10)はそのまま残してください
・「かどもり」「カドモリ」という会社名は、必ずカタカナで「カドモリ」と表記してください
・意味が通り誤字がなければ元の文章のまま出力してください
・Markdown記法（**太字**、### 見出しなど）は使わず、プレーンテキストで出力してください

【文字起こし】
${transcript}

修正後の全文のみを出力してください（前置きや説明は不要です）。`;
  const response = UrlFetchApp.fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`,
    { method: 'POST', contentType: 'application/json', payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.1, maxOutputTokens: 65536, thinkingConfig: { thinkingBudget: 0 } } }), muteHttpExceptions: true }
  );
  const result = JSON.parse(response.getContentText());
  if (!result.candidates) return transcript;
  const cleaned = result.candidates[0].content.parts[result.candidates[0].content.parts.length - 1].text;
  return stripMarkdown_(cleaned);
}

function calcTotalScoreFromChecklist_(checklistDetail) {
  let total = 0;
  (checklistDetail || '').split('\n').forEach(line => {
    const match = line.match(/(\d+)\/(\d+)/);
    if (match) total += parseInt(match[1], 10);
  });
  return total;
}

function generateAISummary(transcript, checklistDetail, selfEvaluation) {
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const totalScore = calcTotalScoreFromChecklist_(checklistDetail);
  const prompt = `以下は1on1面談の記録です。評価者の評価を主軸に、スタッフの自己評価を補足として参考にしながら、AIとして分析・フィードバックを生成してください。

【文字起こし】
${transcript}

【評価者チェックリスト結果】
${checklistDetail}

【総合スコア】
${totalScore}/37点（これは確定値です。自分で計算し直したり、別の数字を出したりしないでください）

【スタッフ自己評価（補足）】
${selfEvaluation || 'なし'}

以下の形式で日本語で回答してください：
■ 面談要約（3〜5行、総合スコアに触れる場合は必ず「${totalScore}/37点」とそのまま記載してください）
■ 良かった点
■ 改善ポイント
■ 次回アクション

文字起こしや評価者チェックリストに書かれていない事実(行動規範の名称、制度名、点数など)を推測で作り出さないでください。不明な場合は触れずに省略してください。「かどもり」「カドモリ」という会社名は、必ずカタカナで「カドモリ」と表記してください。

出力形式について：Markdown記法（**太字**、### 見出し、-や*による箇条書きなど）は一切使わず、プレーンテキストのみで書いてください。箇条書きが必要な場合は「・」を使ってください。`;
  const response = UrlFetchApp.fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`,
    { method: 'POST', contentType: 'application/json', payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.2, maxOutputTokens: 65536, thinkingConfig: { thinkingBudget: 0 } } }) }
  );
  const result = JSON.parse(response.getContentText());
  const rawText = result.candidates[0].content.parts[result.candidates[0].content.parts.length - 1].text;
  return stripMarkdown_(rawText);
}

function stripMarkdown_(text) {
  return (text || '')
    .replace(/^#{1,6}\s*/gm, '')
    .replace(/\*\*(.*?)\*\*/g, '$1')
    .replace(/^\s*[\*\-]\s+/gm, '・')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function getResponseHeader_(response, name) {
  const headers = response.getHeaders();
  const key = Object.keys(headers).find(k => k.toLowerCase() === name.toLowerCase());
  return key ? headers[key] : null;
}

function getAudioChunkFolder_() {
  const folderName = '一時_音声アップロード';
  const folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

function saveAudioChunk_(uploadId, chunkIndex, base64Chunk) {
  const folder = getAudioChunkFolder_();
  const paddedIndex = ('000000' + chunkIndex).slice(-6);
  folder.createFile(uploadId + '_' + paddedIndex, base64Chunk, MimeType.PLAIN_TEXT);
}

function assembleAudioChunks_(uploadId, totalChunks) {
  const folder = getAudioChunkFolder_();
  const parts = [];
  for (let i = 0; i < totalChunks; i++) {
    const paddedIndex = ('000000' + i).slice(-6);
    const fileName = uploadId + '_' + paddedIndex;
    const files = folder.getFilesByName(fileName);
    if (!files.hasNext()) throw new Error('アップロードされたデータの一部が見つかりません(' + (i + 1) + '/' + totalChunks + ')');
    const file = files.next();
    parts.push(file.getBlob().getDataAsString());
    file.setTrashed(true);
  }
  return parts.join('');
}

function transcribeAudio(audioBase64, mimeType) {
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const bytes = Utilities.base64Decode(audioBase64);

  // 1. 再開可能アップロードセッションを開始
  const startResponse = UrlFetchApp.fetch(
    `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${GEMINI_API_KEY}`,
    {
      method: 'post',
      headers: {
        'X-Goog-Upload-Protocol': 'resumable',
        'X-Goog-Upload-Command': 'start',
        'X-Goog-Upload-Header-Content-Length': String(bytes.length),
        'X-Goog-Upload-Header-Content-Type': mimeType
      },
      contentType: 'application/json',
      payload: JSON.stringify({ file: { display_name: '1on1_audio_' + new Date().getTime() } }),
      muteHttpExceptions: true
    }
  );
  const uploadUrl = getResponseHeader_(startResponse, 'X-Goog-Upload-URL');
  if (!uploadUrl) throw new Error('音声アップロードの開始に失敗しました: ' + startResponse.getContentText());

  // 2. 音声本体をアップロード
  const uploadResponse = UrlFetchApp.fetch(uploadUrl, {
    method: 'post',
    headers: {
      'X-Goog-Upload-Command': 'upload, finalize',
      'X-Goog-Upload-Offset': '0'
    },
    contentType: mimeType,
    payload: bytes,
    muteHttpExceptions: true
  });
  const fileInfo = JSON.parse(uploadResponse.getContentText());
  let file = fileInfo.file;
  if (!file || !file.uri) throw new Error('音声アップロードに失敗しました: ' + uploadResponse.getContentText());

  // 3. ファイルがACTIVEになるまで待機
  let attempts = 0;
  while (file.state === 'PROCESSING' && attempts < 45) {
    Utilities.sleep(2000);
    const statusResponse = UrlFetchApp.fetch(
      `https://generativelanguage.googleapis.com/v1beta/${file.name}?key=${GEMINI_API_KEY}`,
      { muteHttpExceptions: true }
    );
    file = JSON.parse(statusResponse.getContentText());
    attempts++;
  }
  if (file.state !== 'ACTIVE') throw new Error('音声の処理がタイムアウトしました');

  // 4. アップロード済みファイルを参照して文字起こし(繰り返しループが出た場合は温度を変えて再試行)
  const transcriptionPrompt = 'この音声を日本語で文字起こししてください。話者の発言をそのまま書き起こし、要約や説明は付けないでください。聞き取れない・自信がない箇所は、推測で言葉を埋めず「[聞き取れません]」と記載してください。実際に発話されていない単語や固有名詞を作り出さないでください。同じ文字・単語・記号を不自然に繰り返さないでください。無音や雑音の箇所は無理に埋めず省略してください。「かどもり」「カドモリ」という会社名は、必ずカタカナで「カドモリ」と表記してください。';
  const temperaturesToTry = [0.15, 0.6];
  let lastText = '';
  for (let i = 0; i < temperaturesToTry.length; i++) {
    const genResponse = UrlFetchApp.fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{
            parts: [
              { text: transcriptionPrompt },
              { file_data: { mime_type: mimeType, file_uri: file.uri } }
            ]
          }],
          generationConfig: { temperature: temperaturesToTry[i], maxOutputTokens: 65536, thinkingConfig: { thinkingBudget: 0 } }
        }),
        muteHttpExceptions: true
      }
    );
    const result = JSON.parse(genResponse.getContentText());
    if (!result.candidates) throw new Error('文字起こしに失敗しました: ' + genResponse.getContentText());
    lastText = result.candidates[0].content.parts[result.candidates[0].content.parts.length - 1].text;
    if (!isRepetitiveText_(lastText)) return lastText;
  }
  throw new Error('音声の文字起こしが繰り返しループになり、正常に生成できませんでした。もう一度お試しください。');
}

function isRepetitiveText_(text) {
  if (!text || text.length < 30) return false;
  const sample = text.slice(0, 400);
  for (const unitLen of [1, 2, 3, 4]) {
    const unit = sample.slice(0, unitLen);
    let repeatCount = 0;
    for (let pos = 0; pos + unitLen <= sample.length; pos += unitLen) {
      if (sample.slice(pos, pos + unitLen) === unit) repeatCount++;
      else break;
    }
    if (repeatCount * unitLen >= 60) return true;
  }
  return false;
}

function testGemini() {
  try {
    const result = generateAISummary('テスト文字起こし：今月は接客の改善に取り組みました。', '素直: 4/6\n行動スピード: 5/6\n振り返り・前進: 3/6', 'よく頑張れたと思います。');
    Logger.log('結果: ' + result);
  } catch(e) { Logger.log('エラー: ' + e.message); }
}

function testApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  Logger.log('APIキー: ' + key);
}

function getSelfEvaluation(staffName, clinic) {
  const ss = SpreadsheetApp.openById(SELF_EVAL_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('スタッフ自己評価');
  const data = sheet.getDataRange().getDisplayValues();
  let bestMatch = null;
  let bestDiff = Infinity;
  const now = new Date();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][11] === staffName && data[i][1] === clinic) {
      const formDate = new Date(data[i][2]);
      const diff = Math.abs(formDate - now);
      if (diff < 7 * 24 * 60 * 60 * 1000 && diff < bestDiff) {
        bestDiff = diff;
        bestMatch = {
          素直: data[i][12], 行動スピード: data[i][13], 振り返り前進: data[i][14],
          継続挑戦: data[i][15], キッカケづくり: data[i][16], 凡事徹底: data[i][17],
          頑張ったこと: data[i][9], 課題: data[i][10],
          素直項目: data[i][3], 行動スピード項目: data[i][4], 振り返り前進項目: data[i][5],
          継続挑戦項目: data[i][6], キッカケづくり項目: data[i][7], 凡事徹底項目: data[i][8]
        };
      }
    }
  }
  return bestMatch;
}

/**
 * 評価画面から直近1週間以内のスタッフ自己評価を参照するためのエンドポイント。
 */
function getSelfEvaluationForUI(clinic, staffName) {
  let result = null;
  try {
    result = getSelfEvaluation(staffName, clinic);
  } catch (err) {
    Logger.log('getSelfEvaluationForUI error: ' + err.message);
  }
  return ContentService.createTextOutput(JSON.stringify(result || {})).setMimeType(ContentService.MimeType.JSON);
}

function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const formId = '1kc-yNOdph3AG8BP98oJ6qEHq44xa249cuPT2YT-p-H4';
  const responseUrl = `https://docs.google.com/forms/d/${formId}/edit#responses`;
  const maxCounts = [6, 6, 6, 6, 6, 7];
  for (let i = 0; i < 6; i++) {
    const cellValue = sheet.getRange(row, 5 + i).getValue();
    const checkedCount = cellValue ? cellValue.toString().split(',').length : 0;
    sheet.getRange(row, 13 + i).setValue(checkedCount + '/' + maxCounts[i]);
  }
  sheet.getRange(row, 19).setValue(responseUrl);
}

function showLastUserId() {
  const uid = PropertiesService.getScriptProperties().getProperty('LAST_USER_ID');
  Logger.log('保存されたuserId: ' + (uid || '（まだありません）'));
}

/**
 * 「LINE_ID」シートを取得（なければヘッダー付きで新規作成）
 */
function getLineIdSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(LINE_ID_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LINE_ID_SHEET_NAME);
    sheet.appendRow(['スタッフ名', 'userId', '登録日時']);
  }
  return sheet;
}

/**
 * スタッフ名からLINEのuserIdを取得する。見つからなければnull。
 */
function getLineUserId(staffName) {
  const sheet = getLineIdSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === staffName) return data[i][1];
  }
  return null;
}

/**
 * 「大阪院」「東京院」シートのスタッフ名一覧から、送られてきた名前と
 * 完全一致するスタッフを探す。
 */
function findStaffByName_(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  for (const clinic of CLINIC_NAMES) {
    const sheet = ss.getSheetByName(clinic);
    if (!sheet) continue;
    const data = sheet.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === name) return { clinic: clinic, name: data[i][2] };
    }
  }
  return null;
}

/**
 * LINEから届いたテキストメッセージを名前として照合し、
 * 一致すればuserIdを登録、本人へ結果を返信する。
 */
function registerLineUser(userId, text, replyToken) {
  const name = (text || '').trim();
  const staff = findStaffByName_(name);

  if (!staff) {
    replyLine(replyToken,
      '「' + name + '」という名前が見つかりませんでした。\nスプレッドシートに登録されている表記とスペースの有無まで含めて、同じお名前で送り直してください。');
    return;
  }

  const sheet = getLineIdSheet_();
  const data = sheet.getDataRange().getValues();
  let updated = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === staff.name) {
      sheet.getRange(i + 1, 2).setValue(userId);
      sheet.getRange(i + 1, 3).setValue(new Date());
      updated = true;
      break;
    }
  }
  if (!updated) {
    sheet.appendRow([staff.name, userId, new Date()]);
  }

  replyLine(replyToken,
    staff.name + ' さん、登録が完了しました。\n次回の面談評価から、このアカウントに自動でカードが届きます。');
}

/**
 * LINEのReply APIでメッセージを返信する
 */
function replyLine(replyToken, text) {
  if (!replyToken) return;
  const TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + TOKEN },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [{ type: 'text', text: text }]
    }),
    muteHttpExceptions: true
  });
}

/**
 * スタッフ名から自動でuserIdを解決し、評価カード画像をLINEで送信する。
 */
function sendLineImage(imageBase64, staffName) {
  try {
    if (!imageBase64 || !staffName) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: '画像またはスタッフ名がありません' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const userId = getLineUserId(staffName);
    if (!userId) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: staffName + 'さんはLINE未登録です。本人にLINE公式アカウントを友だち追加のうえ、お名前を送ってもらってください。'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const base64Data = imageBase64.replace(/^data:image\/png;base64,/, '');
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', staffName + '_評価カード_' + new Date().getTime() + '.png');

    const folderName = 'カドモリ評価カード画像';
    let folder;
    const folders = DriveApp.getFoldersByName(folderName);
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = file.getId();
    const imageUrl = 'https://lh3.googleusercontent.com/d/' + fileId;

    const TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
    const lineUrl = 'https://api.line.me/v2/bot/message/push';
    const payload = {
      to: userId,
      messages: [
        {
          type: 'image',
          originalContentUrl: imageUrl,
          previewImageUrl: imageUrl
        }
      ]
    };

    const response = UrlFetchApp.fetch(lineUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + TOKEN },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    if (code === 200) {
      return ContentService.createTextOutput(JSON.stringify({ success: true, imageUrl: imageUrl }))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'LINE送信エラー: ' + response.getContentText() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}