function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("JSONからGoogleForm作成")
    .addItem("JSONを直接入力してフォーム作成", "showDirectJsonDialog")
    .addItem("JSONをアップロードしてフォーム作成", "showUploadDialog")
    .addItem("アップロード済みJSONからフォーム作成", "showFilePicker")
    .addToUi();
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

    const folder = DriveApp.getRootFolder();
    createdFile = folder.createFile(blob);

    Logger.log("[INFO] アップロード済ファイルID: " + createdFile.getId());

    // JSON → フォーム作成
    createFormFromDriveFile(createdFile.getId());

    // ★削除フラグが有効ならアップロードした JSON を削除
    if (deleteAfter && createdFile) {
      createdFile.setTrashed(true);
      Logger.log("[INFO] JSONファイルを削除しました");
    }

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

  try {
    // JSON 妥当性チェック
    JSON.parse(jsonText);

    const targetFolder = DriveApp.getRootFolder();
    Logger.log("[INFO] My Drive 直下に出力します");

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

  } catch (e) {
    ui.alert("[ERROR] JSON解析またはフォーム生成失敗: " + e.message);
    Logger.log("[ERROR] " + e.stack);
  }
}

/**
 * マイドライブ直下の .json を列挙
 */
function getJsonFileList() {
  const ui = SpreadsheetApp.getUi();
  const folder = DriveApp.getRootFolder(); // 実行ユーザーのマイドライブ直下
  const files = folder.getFiles();

  const fileList = [];

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();

    if (!fileName) {
      Logger.log("[WARN] 無名ファイルを検出、スキップしました");
      continue;
    }

    if (fileName.toLowerCase().endsWith(".json")) {
      fileList.push({
        id: file.getId(),
        name: fileName,
      });
    }
  }

  if (fileList.length === 0) {
    Logger.log("[WARN] マイドライブ直下に JSON ファイルがありません");
    ui.alert("[WARN] マイドライブ直下に JSON ファイルがありません");
  }

  return fileList;
}


/**
 * 指定した Drive の JSON からフォームを生成
 */
function createFormFromDriveFile(fileId) {
  const ui = SpreadsheetApp.getUi();

  let form = null;
  let createdFormId = null;

  try {
    // 作成中インジケータを表示（後続の結果ダイアログで上書きされる）
    showProgressDialog();

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

    // 作成結果を通知（ファイル名・編集URL・回答URL を表示）
    const editUrl = form.getEditUrl();
    const publishedUrl = (typeof form.getPublishedUrl === "function" && form.getPublishedUrl()) || "";
    const formFileName = (createdFormId && DriveApp.getFileById(createdFormId).getName()) || data.title || "新規フォーム";
    showFormResultDialog({
      fileName: formFileName,
      editUrl: editUrl || "なし",
      publishedUrl: publishedUrl || "未公開"
    });
  } catch (e) {
    SpreadsheetApp.getUi().alert("[ERROR] エラー: " + e.message);
    Logger.log("[ERROR] " + e.stack);

    // 途中まで作られたフォームを削除してクリーンアップ
    if (createdFormId) {
      try {
        DriveApp.getFileById(createdFormId).setTrashed(true);
      } catch (ignore) {}
    }
  }
}

/* ---------------- ヘルパー群 ---------------- */

/** 作成中であることを示す簡易ダイアログを表示。次のダイアログで上書きされる想定。 */
function showProgressDialog() {
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial, sans-serif; font-size:13px; line-height:1.6;">' +
      '<p><strong>フォームを作成中です...</strong></p>' +
      '<div style="display:flex;align-items:center;gap:8px;">' +
        '<div style="width:14px;height:14px;border:2px solid #999;border-top-color:#4285f4;border-radius:50%;animation:spin 0.9s linear infinite;"></div>' +
        '<span>しばらくお待ちください</span>' +
      '</div>' +
      '<style>@keyframes spin{from{transform:rotate(0deg);}to{transform:rotate(360deg);}}</style>' +
    '</div>'
  )
    .setWidth(300)
    .setHeight(160);

  SpreadsheetApp.getUi().showModalDialog(html, "作成中");
}

/**
 * 作成結果を表示し、OK ボタンでフォームを開くダイアログを出す。
 * window.open をボタンクリックで実行することでポップアップブロックを避ける。
 */
function showFormResultDialog(params) {
  const fileName = params.fileName || "新規フォーム";
  const editUrl = params.editUrl || "";
  const publishedUrl = params.publishedUrl || "";

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial, sans-serif; font-size:13px; line-height:1.6;">' +
      '<p><strong>作成完了</strong></p>' +
      '<p>ファイル名: ' + fileName + '<br>' +
      '編集URL: ' + (editUrl ? '<a href="' + editUrl + '" target="_blank">' + editUrl + '</a>' : 'なし') + '<br>' +
      '回答URL: ' + (publishedUrl ? '<a href="' + publishedUrl + '" target="_blank">' + publishedUrl + '</a>' : '未公開') + '<br>' +
      '</p>' +
      '<p><button onclick="google.script.host.close();">OK</button></p>' +
    '</div>'
  )
    .setWidth(640)
    .setHeight(240);

  SpreadsheetApp.getUi().showModalDialog(html, "作成結果");
}

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
