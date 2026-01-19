// ==========================================
// 設定エリア
// ==========================================
// APIキーはスクリプトプロパティから読み込み
const API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");

// ★重要: スプレッドシートのファイル名（正確に合わせてください）
const SPREADSHEET_FILE_NAME = "UI_Inventory_DB";

// シート名設定
const SHEET_NAME = "UI_Gallery";
const CONFIG_SHEET_NAME = "Config";
const SCREEN_MASTER_SHEET_NAME = "Screen_Master";
const PROMPT_MASTER_SHEET_NAME = "Prompt_Master";

// 実行時間の制限（秒）
const MAX_EXECUTION_TIME_SEC = 240;
// ==========================================

// 定期実行用関数
function processNewImages() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(1000);
  } catch (e) {
    console.warn("🔒 ロック中のためスキップ");
    return;
  }

  const startTime = new Date().getTime();

  // ★相対パスでスプレッドシート取得
  let ss;
  try {
    ss = getRelativeSpreadsheet();
  } catch (e) {
    console.error(e.message);
    return;
  }

  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

  try {
    // --- 1. スイッチ確認 ---
    if (configSheet) {
      const switchStatus = configSheet.getRange(2, 2).getValue();
      if (switchStatus !== "ON") {
        console.log("😴 スイッチOFF");
        updateStatusMessage(configSheet, "");
        return;
      }
    }

    const sheet = ss.getSheetByName(SHEET_NAME);
    const masterSheet = ss.getSheetByName(SCREEN_MASTER_SHEET_NAME);

    // ★相対パスでInboxフォルダ取得
    let inboxFolder;
    try {
      inboxFolder = getRelativeInboxFolder();
    } catch (e) {
      console.error(e.message);
      return;
    }

    const promptInstructions = loadPromptMasterInstructions(ss);

    // --- 2. Masterから未処理リスト (targetIds) を作成 ---
    const masterData = masterSheet.getDataRange().getValues();
    const idColIdx = masterData[0].indexOf("Screen_ID");
    const dateColIdx = masterData[0].indexOf("Last_Processed");

    if (idColIdx === -1 || dateColIdx === -1) {
      throw new Error("Screen_Masterに Screen_ID または Last_Processed 列がありません。");
    }

    const targetIds = [];
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][idColIdx] && !masterData[i][dateColIdx]) {
        targetIds.push({ row: i + 1, id: String(masterData[i][idColIdx]) });
      }
    }

    // ★【プランB】起動メッセージの表示（高速版）
    updateStatusMessage(configSheet, `🚀 起動中... 残り ${targetIds.length} 画面`);

    // --- 3. 既登録チェック用リスト作成 ---
    const galleryData = sheet.getDataRange().getValues();
    const registeredPaths = new Set();
    for (let i = 1; i < galleryData.length; i++) {
      if (galleryData[i][1]) registeredPaths.add(String(galleryData[i][1]));
    }

    // --- 4. 処理ループ (鉄壁の完遂フラグ版) ---
    let processedTotal = 0;
    let timeLimitReached = false;
    let hasFilesRemaining = false;
    const rootFolderName = inboxFolder.getName();

    for (const target of targetIds) {
      if (timeLimitReached) {
        hasFilesRemaining = true;
        break;
      }

      const screenId = target.id;
      const folders = inboxFolder.getFoldersByName(screenId);
      if (!folders.hasNext()) continue;

      const folder = folders.next();
      if (folder.getName().startsWith("🚫")) continue;

      const files = folder.getFiles();

      // ★フォルダ開始時に「完遂フラグ」を立てる
      let isFolderFullyProcessed = true;

      while (files.hasNext()) {
        const currentTime = new Date().getTime();
        if ((currentTime - startTime) / 1000 > MAX_EXECUTION_TIME_SEC) {
          timeLimitReached = true;
          hasFilesRemaining = true;
          isFolderFullyProcessed = false; // 未完としてマーク
          break;
        }

        const file = files.next();
        const fileName = file.getName();
        if (!file.getMimeType().includes("image")) continue;

        const relativePath = `${rootFolderName}/${screenId}/${fileName}`;
        if (registeredPaths.has(relativePath)) continue;

        if (processedTotal % 3 === 0) {
          updateStatusMessage(configSheet, `🔄 処理中... (${processedTotal}完了)`);
        }

        console.log(`Processing [${screenId}] ${fileName}...`);

        try {
          const result = callGeminiVisionAPI_Dynamic(file.getBlob(), promptInstructions);
          const uniqueId = Utilities.getUuid().slice(0, 8);
          const today = new Date();

          // 書き込み (Created_Date列が13番目の想定で、12番目に空文字を配置)
          sheet.appendRow([uniqueId, relativePath, screenId, result.category, "", result.specificName, result.tags, "", "", "", "", "", today, "", ""]);

          SpreadsheetApp.flush();
          registeredPaths.add(relativePath);
          processedTotal++;
          Utilities.sleep(3000); // 429エラー(API制限)対策
        } catch (e) {
          console.error(`❌ Error in [${screenId}] ${fileName}: ${e.message}`);
          isFolderFullyProcessed = false; // 1つでもコケたらこのフォルダは「未完」

          // API制限(429)の場合は中断
          if (e.message.includes("Resource exhausted")) {
            timeLimitReached = true;
            break;
          }
        }
      }

      // --- 判定：フォルダ内が完全に完了した時だけ日付を記入 ---
      if (isFolderFullyProcessed) {
        masterSheet.getRange(target.row, dateColIdx + 1).setValue(new Date());
        SpreadsheetApp.flush();
        console.log(`✅ ${screenId} の全画像を処理完了`);
      }
    }

    // --- 5. 終了処理 ---
    if (!timeLimitReached && !hasFilesRemaining) {
      if (processedTotal === 0 && !timeLimitReached) {
        console.log(`🎉 完了。`);
        updateStatusMessage(configSheet, "");
        configSheet.getRange(2, 2).setValue("OFF");
        SpreadsheetApp.flush();
      } else {
        updateStatusMessage(configSheet, `⏸ 一時停止。`);
      }
    } else {
      updateStatusMessage(configSheet, `⏳ 時間切れ休憩中...`);
    }
  } catch (e) {
    console.error("予期せぬエラー: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// AppSheet連携用：再生成関数（安全版）
function regenerateSingleImage(uniqueId, relativePath, customInstruction) {
  console.log(`★再生成開始: ID=${uniqueId}`);

  const pathParts = relativePath.split("/");
  if (pathParts.length < 3) {
    console.error("❌ パス形式エラー");
    return;
  }
  const folderName = pathParts[1];
  const fileName = pathParts[2];

  let ss;
  try {
    ss = getRelativeSpreadsheet();
  } catch (e) {
    console.error(e.message);
    return;
  }
  const sheet = ss.getSheetByName(SHEET_NAME);

  try {
    const inbox = getRelativeInboxFolder();
    const targetFolders = inbox.getFoldersByName(folderName);
    if (!targetFolders.hasNext()) {
      console.error(`❌ フォルダなし: ${folderName}`);
      return;
    }
    const targetFolder = targetFolders.next();

    const files = targetFolder.getFilesByName(fileName);
    if (!files.hasNext()) {
      console.error(`❌ ファイルなし: ${fileName}`);
      return;
    }
    const file = files.next();

    const result = callGeminiVisionAPI_Dynamic(file.getBlob(), customInstruction);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colMap = {};
    headers.forEach((h, i) => (colMap[h] = i + 1));
    const data = sheet.getDataRange().getValues();
    const idColIndex = (colMap["Unique_ID"] || colMap["ID"] || colMap["UI_ID"] || 1) - 1;
    let targetRow = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIndex]) === String(uniqueId)) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow !== -1) {
      const colCategory = colMap["Category"] || 4;
      const colSpecific = colMap["Specific_Name"] || colMap["SpecificName"] || 6;
      const colTags = colMap["Tags"] || 7;
      sheet.getRange(targetRow, colCategory).setValue(result.category);
      sheet.getRange(targetRow, colSpecific).setValue(result.specificName);
      sheet.getRange(targetRow, colTags).setValue(result.tags);
      SpreadsheetApp.flush();
      console.log("✅ 更新完了");
    }
  } catch (e) {
    console.error("❌ Error: " + e.message);
  }
}

// ==========================================
// ★ヘルパー関数
// ==========================================

function getRelativeInboxFolder() {
  const parent = DriveApp.getFileById(ScriptApp.getScriptId()).getParents().next();
  const folders = parent.getFoldersByName("_INBOX");
  if (!folders.hasNext()) throw new Error(`同じ階層に "_INBOX" フォルダが見つかりません。`);
  return folders.next();
}

function getRelativeSpreadsheet() {
  const parent = DriveApp.getFileById(ScriptApp.getScriptId()).getParents().next();
  const files = parent.getFilesByName(SPREADSHEET_FILE_NAME);
  if (!files.hasNext()) throw new Error(`同じ階層に "${SPREADSHEET_FILE_NAME}" が見つかりません。`);
  return SpreadsheetApp.open(files.next());
}

function loadPromptMasterInstructions(ss) {
  const sheet = ss.getSheetByName(PROMPT_MASTER_SHEET_NAME);
  if (!sheet) return "";
  const data = sheet.getDataRange().getValues();
  let instructions = "";
  for (let i = 1; i < data.length; i++) {
    const category = data[i][0];
    const text = data[i][1];
    if (category && text) instructions += `- **${category}の場合**: ${text}\n`;
  }
  return instructions;
}

function callGeminiVisionAPI_Dynamic(imageBlob, instructionBlock) {
  const model = "gemini-2.0-flash";
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${API_KEY}`;
  const finalPrompt = `
  あなたはUIデザインシステムの構築を支援するAIです。画像を解析し、以下のステップでJSONを出力してください。
  【Step 1: カテゴリ判定】
  画像がどのカテゴリ(Atom, Component, Unit, Dialog, Modal, Table)に属するか判定してください。
  **重要ルール:**
  - **Button (ボタン)** は必ず「Component」に分類すること。
  - **Table (テーブル)** の一部(ヘッダーや行)も「Table」に分類すること。
  【Step 2: 詳細タグ・説明生成】
  判定したカテゴリに応じ、以下のガイドラインに従って情報を生成してください。
  ユーザーからの追加指示がある場合は、必ず "description" フィールドに反映してください。
  ${instructionBlock}
  【出力JSON形式】
  {
    "category": "カテゴリ名",
    "specificName": "名称（日本語）",
    "tags": "タグ（日本語）",
    "description": "画像の説明文。"
  }`;

  const payload = {
    contents: [
      {
        parts: [{ text: finalPrompt }, { inline_data: { mime_type: imageBlob.getContentType(), data: Utilities.base64Encode(imageBlob.getBytes()) } }],
      },
    ],
    generationConfig: { response_mime_type: "application/json" },
  };

  const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());
  if (json.error) throw new Error(json.error.message);
  return JSON.parse(json.candidates[0].content.parts[0].text.replace(/```json|```/g, "").trim());
}

function updateStatusMessage(configSheet, message) {
  if (configSheet && message !== undefined) {
    try {
      configSheet.getRange(2, 3).setValue(message);
      SpreadsheetApp.flush();
    } catch (e) {}
  }
}
