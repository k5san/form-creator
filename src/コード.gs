function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("JSONからGoogleForm作成")
    .addItem("JSONを直接入力してフォーム作成", "showDirectJsonDialog")
    .addItem("JSONをアップロードしてフォーム作成", "showUploadDialog")
    .addItem("アップロード済みJSONからフォーム作成", "showFilePicker")
    .addToUi();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("B2").setValue("");
}

function showFilePicker() {
  const html = HtmlService.createHtmlOutputFromFile("filepicker")
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, "JSONファイルを選択");
}

function showUploadDialog() {
  const html = HtmlService.createHtmlOutputFromFile("uploadForm")
    .setWidth(400)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, "JSONファイルをアップロード");
}

function uploadFile(data, name, deleteAfter) {
  const ui = SpreadsheetApp.getUi();

  let createdFile = null;
  let createdFormId = null;

  try {
    if (!data || typeof data !== "string") {
      throw new Error("ファイルデータを正しく受け取れていません。");
    }

    const typeMatch = data.match(/^data:(.*?);/);
    if (!typeMatch) {
      throw new Error("ContentType を取得できませんでした。");
    }

    const contentType = typeMatch[1];
    const base64 = data.split(",")[1];
    const bytes = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(bytes, contentType, name);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const folderName = sheet.getRange("B1").getValue();

    if (!folderName) {
      throw new Error("B1セルにフォルダ名が入力されていません。");
    }

    const folders = DriveApp.getFoldersByName(folderName);
    if (!folders.hasNext()) {
      throw new Error("指定されたフォルダ '" + folderName + "' が見つかりません。");
    }

    const folder = folders.next();
    createdFile = folder.createFile(blob);

    Logger.log("[INFO] アップロード済ファイルID: " + createdFile.getId());

    // JSON → フォーム作成
    createFormFromDriveFile(createdFile.getId());

    // ★削除フラグが有効ならアップロードした JSON を削除
    if (deleteAfter && createdFile) {
      createdFile.setTrashed(true);
      Logger.log("[INFO] JSONファイルを削除しました");
    }

    ui.alert("フォーム作成完了");

  } catch (e) {
    ui.alert("[ERROR] アップロードまたはフォーム生成に失敗: " + e.message);
    Logger.log("[ERROR] " + e.stack);
  }
}

function showDirectJsonDialog() {
  const html = HtmlService.createHtmlOutputFromFile("directInput")
    .setWidth(450)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, "JSON直接入力");
}

function uploadFromText(jsonText, name, saveFlag) {
  const ui = SpreadsheetApp.getUi();
  let createdFile = null;
  let targetFolder = null;

  try {
    // JSON 妥当性チェック
    JSON.parse(jsonText);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const folderName = sheet.getRange("B1").getValue();

    // ①B1フォルダを試す
    if (folderName) {
      try {
        const folders = DriveApp.getFoldersByName(folderName);
        if (folders.hasNext()) {
          targetFolder = folders.next();
          Logger.log("[INFO] B1指定フォルダ使用: " + folderName);
        }
      } catch (e) {
        Logger.log("[WORN] B1フォルダアクセス不可: " + e.message);
      }
    } else {
      Logger.log("[WORN] B1フォルダ名未入力");
    }

    // ②fallback: Spreadsheetと同一フォルダ
    if (!targetFolder) {
      const ssFile = DriveApp.getFileById(
        SpreadsheetApp.getActiveSpreadsheet().getId()
      );
      const parents = ssFile.getParents();
      if (!parents.hasNext()) {
        throw new Error("スプレッドシートの親フォルダが取得できません");
      }
      targetFolder = parents.next();
      Logger.log("[INFO] fallback: スプレッドシートと同一フォルダ使用");
    }

    const blob = Utilities.newBlob(jsonText, "application/json", name);

    if (saveFlag) {
      // 保存
      createdFile = targetFolder.createFile(blob);
      Logger.log("[INFO] JSON保存: " + createdFile.getId());
      createFormFromDriveFile(createdFile.getId());
    } else {
      // 一時ファイル → フォーム生成後削除
      createdFile = targetFolder.createFile(blob);
      Logger.log("[INFO] 一時JSON: " + createdFile.getId());
      createFormFromDriveFile(createdFile.getId());
      createdFile.setTrashed(true);
      Logger.log("[INFO] 一時JSONを削除しました");
    }

    ui.alert("フォーム作成完了（直接入力）");

  } catch (e) {
    ui.alert("[ERROR] JSON解析またはフォーム生成失敗: " + e.message);
    Logger.log("[ERROR] " + e.stack);
  }
}

/**
 * B1セルで指定したフォルダ内の .json を列挙
 */
function getJsonFileList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const folderName = sheet.getRange("B1").getValue();
  const ui = SpreadsheetApp.getUi();

  if (!folderName) {
    ui.alert("[ERROR] フォルダ名がB1セルに設定されていません");
    return;
  }

  const folders = DriveApp.getFoldersByName(folderName);
  if (!folders.hasNext()) {
    ui.alert("[ERROR] フォルダ '" + folderName + "' が見つかりません");
    return;
  }

  const folder = folders.next();
  const files = folder.getFiles();
  const fileList = [];

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    if (fileName.toLowerCase().endsWith(".json")) {
      fileList.push({ id: file.getId(), name: fileName });
    }
  }

  if (fileList.length === 0) {
    Logger.log("[WORN] 対象の.jsonファイルが見つかりませんでした");
    ui.alert("[WORN] 対象の.jsonファイルが見つかりませんでした");
  }

  return fileList;
}

/**
 * 指定した Drive の JSON からフォームを生成
 */
function createFormFromDriveFile(fileId) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("B2").setValue("");

  let form = null;
  let createdFormId = null;

  try {
    const file = DriveApp.getFileById(fileId);
    const jsonText = file.getBlob().getDataAsString();
    const data = JSON.parse(jsonText);

    // 入力バリデーション（最低限）
    assertTruthy(data.title, "title が未指定です");
    assertArray(data.questions, "questions は配列である必要があります");

    form = FormApp.create(data.title);
    createdFormId = form.getId();

    if (data.description) {
      form.setDescription(data.description);
    }

    // choicesTemplates を連想配列化（description 対応）
    // 形式： name -> { choices: [...], description: "..." }
    const choicesTemplates = {};
    if (Array.isArray(data.choicesTemplates)) {
      data.choicesTemplates.forEach(t => {
        assertTruthy(t && t.name, "choicesTemplates の name が未指定です: " + JSON.stringify(t));
        choicesTemplates[t.name] = {
          choices: t.choices || [],
          description: t.description || ""
        };
      });
    }

    // itemTemplates を連想配列化（rows 用）
    // 形式： name -> [ "row1", "row2", ... ]
    const itemTemplates = {};
    if (Array.isArray(data.itemTemplates)) {
      data.itemTemplates.forEach(t => {
        assertTruthy(t && t.name, "itemTemplates の name が未指定です: " + JSON.stringify(t));
        assertArray(t.items, "itemTemplates の items は配列である必要があります: " + JSON.stringify(t));
        itemTemplates[t.name] = t.items || [];
      });
    }

    // 設問を順に作成
    data.questions.forEach(q => {
      assertTruthy(q && q.type, "設問に type がありません: " + JSON.stringify(q));

      let item = null;
      // 参照テンプレートの説明を蓄積（choices/columns など複数参照した場合は連結）
      const templateDescriptions = [];

      switch (String(q.type).toLowerCase()) {
        case "multiplechoice": {
          item = form.addMultipleChoiceItem();
          const labels = resolveChoiceLabels(q.choices, choicesTemplates, q.title, templateDescriptions);
          assertNoDuplicates(labels, q.title);
          const choices = labels.map(label => item.createChoice(label));
          item.setChoices(choices);
          break;
        }
        case "checkbox": {
          item = form.addCheckboxItem();
          const labels = resolveChoiceLabels(q.choices, choicesTemplates, q.title, templateDescriptions);
          assertNoDuplicates(labels, q.title);
          const choices = labels.map(label => item.createChoice(label));
          item.setChoices(choices);
          break;
        }
        case "list": {
          item = form.addListItem();
          const labels = resolveChoiceLabels(q.choices, choicesTemplates, q.title, templateDescriptions);
          assertNoDuplicates(labels, q.title);
          item.setChoiceValues(labels);
          break;
        }
        case "text": {
          item = form.addTextItem();
          break;
        }
        case "paragraph": {
          item = form.addParagraphTextItem();
          break;
        }
        case "scale": {
          assertNumber(q.min, "scale の min が数値ではありません: " + JSON.stringify(q));
          assertNumber(q.max, "scale の max が数値ではありません: " + JSON.stringify(q));
          item = form.addScaleItem()
            .setBounds(q.min, q.max)
            .setLabels(q.minLabel || "", q.maxLabel || "");
          break;
        }
        case "date": {
          item = form.addDateItem();
          break;
        }
        case "time": {
          item = form.addTimeItem();
          break;
        }
        case "grid": {
          item = form.addGridItem();
          const rowLabels = resolveRowLabels(q.rows, itemTemplates, q.title);
          const columnLabels = resolveChoiceLabels(q.columns, choicesTemplates, q.title, templateDescriptions);
          assertNoDuplicates(rowLabels, q.title + "（rows）");
          assertNoDuplicates(columnLabels, q.title + "（columns）");
          item.setRows(rowLabels);
          item.setColumns(columnLabels);
          break;
        }
        case "checkboxgrid": {
          item = form.addCheckboxGridItem();
          const rowLabels = resolveRowLabels(q.rows, itemTemplates, q.title);
          const columnLabels = resolveChoiceLabels(q.columns, choicesTemplates, q.title, templateDescriptions);
          assertNoDuplicates(rowLabels, q.title + "（rows）");
          assertNoDuplicates(columnLabels, q.title + "（columns）");
          item.setRows(rowLabels);
          item.setColumns(columnLabels);
          break;
        }
        case "pagebreak": {
          item = form.addPageBreakItem();
          break;
        }
        default:
          Logger.log("[WORN] 未対応の設問タイプ: " + q.type);
          SpreadsheetApp.getUi().alert("[WORN] 未対応の設問タイプ: " + q.type);
          return; // 未対応は以降の setTitle を避けるため return
      }

      // 共通属性（helpText への description 追記を含む）
      if (item) {
        item.setTitle(q.title || "");
        if (q.required !== undefined) {
          item.setRequired(!!q.required);
        }

        const desc = templateDescriptions
          .map(s => (s || "").trim())
          .filter(s => s.length > 0)
          .join("\n");

        if (q.helpText || desc) {
          const ht = (q.helpText || "") + (q.helpText && desc ? "\n" : "") + (desc || "");
          item.setHelpText(ht);
        }
      }
    });

    ui.alert("作成完了");
    if (form.getEditUrl()) {
      sheet.getRange("B2").setValue(form.getEditUrl());
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("[ERROR] エラー: " + e.message);
    Logger.log("[ERROR] " + e.stack);

    // 途中まで作られたフォームを削除してクリーンアップ
    if (createdFormId) {
      try {
        DriveApp.getFileById(createdFormId).setTrashed(true);
      } catch (ignore) {}
    }
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getRange("B2").setValue("出力失敗");
  }
}

/* ---------------- ヘルパー群 ---------------- */

/**
 * choicesDef が
 *  - テンプレ名（string）
 *  - 配列（["A","B"] または [{label,score}, ...]）
 * のどちらでも、最終的に「文字列ラベルの配列」に正規化する。
 * templates は name -> { choices: [...], description: "..." }
 * description を使った場合は descriptions[] に追記する。
 */
function resolveChoiceLabels(choicesDef, templates, contextTitle, descriptions) {
  if (choicesDef && typeof choicesDef === "string" && templates[choicesDef]) {
    const tpl = templates[choicesDef];
    if (descriptions && tpl.description) {
      descriptions.push(String(tpl.description));
    }
    return normalizeLabels(tpl.choices, contextTitle);
  }
  if (Array.isArray(choicesDef)) {
    return normalizeLabels(choicesDef, contextTitle);
  }
  throw new Error("選択肢指定が不正です（" + (contextTitle || "無題") + "）: " + JSON.stringify(choicesDef));
}

/**
 * rowsDef が
 *  - テンプレ名（string）
 *  - 配列（["A","B", ...]）
 * のどちらでも、最終的に「文字列ラベルの配列」に正規化する。
 * rows は {label, score} のようなオブジェクトは想定しない（アイテム名のみ）。
 */
function resolveRowLabels(rowsDef, itemTemplates, contextTitle) {
  if (rowsDef && typeof rowsDef === "string" && itemTemplates[rowsDef]) {
    return normalizeRowItems(itemTemplates[rowsDef], contextTitle);
  }
  if (Array.isArray(rowsDef)) {
    return normalizeRowItems(rowsDef, contextTitle);
  }
  throw new Error("rows 指定が不正です（" + (contextTitle || "無題") + "）: " + JSON.stringify(rowsDef));
}

/**
 * 文字列配列 or {label, score} 配列を「ラベルの配列」に正規化（choices 用）。
 */
function normalizeLabels(arr, contextTitle) {
  return arr.map(v => {
    if (v == null) {
      throw new Error("null/undefined の選択肢があります（" + (contextTitle || "無題") + "）");
    }
    if (typeof v === "object") {
      if (typeof v.label !== "string" || v.label.trim() === "") {
        throw new Error("label が不正です（" + (contextTitle || "無題") + "）: " + JSON.stringify(v));
      }
      return v.label;
    }
    return String(v);
  });
}

/**
 * rows 用：純テキスト配列のみを許可（オブジェクトは不可）。
 */
function normalizeRowItems(arr, contextTitle) {
  return arr.map(v => {
    if (v == null) {
      throw new Error("null/undefined の行要素があります（" + (contextTitle || "無題") + "）");
    }
    if (typeof v === "object") {
      throw new Error("rows は文字列のみ対応です（" + (contextTitle || "無題") + "）: " + JSON.stringify(v));
    }
    return String(v);
  });
}

/**
 * 配列内の重複を検出してエラーにする。
 */
function assertNoDuplicates(values, contextTitle) {
  const seen = new Set();
  const dups = new Set();
  values.forEach(v => {
    if (seen.has(v)) {
      dups.add(v);
    } else {
      seen.add(v);
    }
  });
  if (dups.size > 0) {
    throw new Error("質問に重複した選択肢/行の値があります（" + (contextTitle || "無題") + "）: " + Array.from(dups).join(", "));
  }
}

/* ------ 軽量バリデーション ------ */

function assertTruthy(value, message) {
  if (!value) throw new Error(message);
}
function assertArray(value, message) {
  if (!Array.isArray(value)) throw new Error(message);
}
function assertNumber(value, message) {
  if (typeof value !== "number") throw new Error(message);
}
