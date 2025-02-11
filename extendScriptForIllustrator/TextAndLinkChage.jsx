// =============== ユーティリティ関数群 ===============

/**
 * mm(ミリメートル) を pt(ポイント) に変換する関数
 * @param {number} mm
 * @returns {number} pt
 */
function mmToPt(mm) {
  return mm * 2.83464567; // 1mm = 約2.83464567pt
}

/**
 * 文字列を長さ順（昇順）にソートするヘルパー関数
 * findTexts, replaceTextsのペアを同時に並び替える
 * @param {string[]} findTexts
 * @param {string[]} replaceTexts
 */
function sortTextsByLengthAsc(findTexts, replaceTexts) {
  var pairs = [];
  for (var i = 0; i < findTexts.length; i++) {
    pairs.push({ findText: findTexts[i], replaceText: replaceTexts[i] });
  }
  // 文字の長さでソート
  pairs.sort(function (a, b) {
    return a.findText.length - b.findText.length;
  });
  // 並び替え後に元の配列に戻す
  for (var i = 0; i < pairs.length; i++) {
    findTexts[i] = pairs[i].findText;
    replaceTexts[i] = pairs[i].replaceText;
  }
}

/**
 * CSVファイルを読み込み、2次元配列にして返す
 * @param {File} file
 * @returns {string[][]}
 */
function readCSV(file) {
  var csvFile = File(file);
  var data = [];
  if (csvFile.open("r")) {
    while (!csvFile.eof) {
      var line = csvFile.readln();
      data.push(line.split(","));
    }
    csvFile.close();
  } else {
    alert("CSVファイルを開くことができませんでした。");
  }
  return data;
}

/**
 * CSV内容確認ダイアログ（必要に応じて呼び出し）
 * @param {string[][]} data
 * @param {number} currentRow
 */
function showCSVDialog(data, currentRow) {
  var dialog = new Window("dialog", "CSV内容の確認");
  dialog.orientation = "column";
  dialog.size = [400, 700];

  var listBox = dialog.add("listbox", undefined, "", {
    numberOfColumns: 1,
    showHeaders: false,
    columnTitles: [],
  });
  listBox.preferredSize = { width: 380, height: 630 };

  for (var col = 0; col < data[0].length; col++) {
    if (currentRow < data.length && col < data[currentRow].length) {
      var text = data[0][col] + " → " + data[currentRow][col];
      listBox.add("item", text);
    }
  }

  var buttonGroup = dialog.add("group");
  buttonGroup.orientation = "row";
  buttonGroup.alignment = "center";

  var cancelButton = buttonGroup.add("button", undefined, "キャンセル", {
    name: "cancel",
  });
  var okButton = buttonGroup.add("button", undefined, "OK", { name: "ok" });

  okButton.onClick = function () {
    dialog.close(1);
  };
  cancelButton.onClick = function () {
    dialog.close(0);
  };
  return dialog.show();
}

/**
 * テキストフレーム1つを対象に検索置換を行う
 * @param {TextFrame} textFrame
 * @param {string} findText
 * @param {string} replaceText
 * @param {boolean} excludeLockedLayers
 * @param {boolean} excludeHiddenLayers
 * @returns {number} 置換件数
 */
function processTextFrame(
  textFrame,
  findText,
  replaceText,
  excludeLockedLayers,
  excludeHiddenLayers
) {
  var count = 0;

  // ロック・非表示レイヤーチェック
  if (excludeLockedLayers && textFrame.layer.locked) return count;
  if (excludeHiddenLayers && !textFrame.layer.visible) return count;

  var fullText = textFrame.contents;
  var startIndex = fullText.indexOf(findText);

  while (startIndex !== -1) {
    var endIndex = startIndex + findText.length;
    var textRange = textFrame.textRange;
    textRange.start = startIndex;
    textRange.end = endIndex;

    // 現在の文字属性を保持しておく
    var prevCharAttrs = textRange.characterAttributes;
    // 置換実行
    textRange.contents = replaceText;
    // 新しいテキストレンジを取得して属性を引き継ぐ
    var newRange = textFrame.textRanges[textFrame.textRanges.length - 1];
    newRange.characterAttributes = prevCharAttrs;

    count++;
    // 次の検索開始位置を設定
    fullText = textFrame.contents;
    startIndex = fullText.indexOf(findText, startIndex + replaceText.length);
  }
  return count;
}

/**
 * CSVデータ1行の中から TEXT_ で始まる列をテキスト置換する
 * @param {Document} doc
 * @param {string[]} findTexts
 * @param {string[]} replaceTexts
 * @param {boolean} excludeLockedLayers
 * @param {boolean} excludeHiddenLayers
 * @param {boolean} skipBlankReplace
 * @returns {number} 総置換件数
 */
function processCSVRow(
  doc,
  findTexts,
  replaceTexts,
  excludeLockedLayers,
  excludeHiddenLayers,
  skipBlankReplace
) {
  var replaceCount = 0;
  var textFrames = doc.textFrames;

  for (var i = 0; i < findTexts.length; i++) {
    var findText = findTexts[i];
    var replaceText = replaceTexts[i];

    // スキップする条件
    if (!findText) continue; // 検索文字列が空ならスキップ
    if (skipBlankReplace && !replaceText) {
      // 置換文字列が空のときはスキップするかどうか
      continue;
    }

    // 実際の置換実行
    for (var j = 0; j < textFrames.length; j++) {
      replaceCount += processTextFrame(
        textFrames[j],
        findText,
        replaceText || "",
        excludeLockedLayers,
        excludeHiddenLayers
      );
    }
  }
  return replaceCount;
}

/**
 * グループやクリッピングマスク内も含めてリンクアイテムを再帰的に処理する
 * @param {PageItem} item
 * @param {string[]} findTexts
 * @param {string[]} replaceTexts
 * @param {boolean} excludeLockedLayers
 * @param {boolean} excludeHiddenLayers
 * @returns {number} 置換件数
 */
function processItem(
  item,
  findTexts,
  replaceTexts,
  excludeLockedLayers,
  excludeHiddenLayers
) {
  var replaceCount = 0;

  // ロック・非表示レイヤーチェック
  if (excludeLockedLayers && item.layer.locked) return replaceCount;
  if (excludeHiddenLayers && !item.layer.visible) return replaceCount;

  if (item.typename === "PlacedItem") {
    var originalFileName = decodeURI(item.file.name);
    for (var j = 0; j < findTexts.length; j++) {
      if (!findTexts[j] || !replaceTexts[j]) continue;
      // リンクファイル名が一致したらrelink
      if (originalFileName === findTexts[j]) {
        var newFile = new File(item.file.path + "/" + replaceTexts[j]);
        if (newFile.exists) {
          item.relink(newFile);
          replaceCount++;
        }
        break;
      }
    }
  } else if (item.typename === "GroupItem" || item.typename === "Layer") {
    // グループやレイヤーなら、その子要素も再帰的に処理
    for (var k = 0; k < item.pageItems.length; k++) {
      replaceCount += processItem(
        item.pageItems[k],
        findTexts,
        replaceTexts,
        excludeLockedLayers,
        excludeHiddenLayers
      );
    }
  }
  return replaceCount;
}

/**
 * AIドキュメント全体に対してリンク置換を行う
 * @param {Document} doc
 * @param {string[]} findTexts
 * @param {string[]} replaceTexts
 * @param {boolean} excludeLockedLayers
 * @param {boolean} excludeHiddenLayers
 */
function processLinks(
  doc,
  findTexts,
  replaceTexts,
  excludeLockedLayers,
  excludeHiddenLayers
) {
  var replaceCount = 0;
  for (var i = 0; i < doc.pageItems.length; i++) {
    replaceCount += processItem(
      doc.pageItems[i],
      findTexts,
      replaceTexts,
      excludeLockedLayers,
      excludeHiddenLayers
    );
  }
  return replaceCount;
}

/**
 * オーバーフローを解消する
 * @param {Document} doc
 * @param {boolean} excludeLockedLayers
 * @param {boolean} excludeHiddenLayers
 * @param {boolean} showAlerts
 */
function overflowFix(
  doc,
  excludeLockedLayers,
  excludeHiddenLayers,
  showAlerts
) {
  var overflowCount = 0;
  var selectedFrames = [];
  var textFrames = doc.textFrames;

  for (var i = 0; i < textFrames.length; i++) {
    var textFrame = textFrames[i];
    if (excludeLockedLayers && textFrame.layer.locked) continue;
    if (excludeHiddenLayers && !textFrame.layer.visible) continue;

    // オーバーフローしているかを判定
    var totalCharacterCount = textFrame.contents.length;
    var visibleCharacterCount = 0;
    for (var k = 0; k < textFrame.lines.length; k++) {
      visibleCharacterCount += textFrame.lines[k].contents.length;
    }
    if (totalCharacterCount > visibleCharacterCount) {
      overflowCount++;
      selectedFrames.push(textFrame);

      // 文字縮小ループ(強引…)
      var horizontalScale =
        textFrame.textRange.characterAttributes.horizontalScale;
      while (
        totalCharacterCount > visibleCharacterCount &&
        horizontalScale > 15
      ) {
        horizontalScale -= 1;
        textFrame.textRange.characterAttributes.horizontalScale =
          horizontalScale;
        visibleCharacterCount = 0;
        for (var l = 0; l < textFrame.lines.length; l++) {
          visibleCharacterCount += textFrame.lines[l].contents.length;
        }
      }
    }
  }

  // 結果表示
  if (overflowCount > 0) {
    doc.selection = selectedFrames;
    if (showAlerts) {
      alert(overflowCount + "個のオーバーフローを解消しました。");
    }
  } else {
    if (showAlerts) {
      alert("オーバーフローテキストはありませんでした。");
    }
  }

  // 位置修正処理の呼び出し（overflowFix完了後にそのまま呼んでいる）
  if (showAlerts) {
    alert("位置修正を開始します");
  }
  adjustPosition(doc, showAlerts);
  if (showAlerts) {
    alert("位置修正を完了しました");
  }
}

/**
 * 位置修正の処理
 * @param {Document} doc
 * @param {boolean} showAlerts
 */
function adjustPosition(doc, showAlerts) {
  // レイヤー名・移動先座標の定義(本当に必要最低限の変更のみ)
  var elements = [
    { name: "レイヤー名を指定1", x: 301.05, y: 40.222 },
    { name: "レイヤー名を指定2", x: 445.5, y: 184.998 },
    { name: "レイヤー名を指定3", x: 360.77, y: 332.379 },
    { name: "レイヤー名を指定4", x: 445.5, y: 184.933 },
    { name: "レイヤー名を指定5", x: 217.14, y: 129.824 },
  ];

  var artboard = doc.artboards[0];
  var artboardRect = artboard.artboardRect;
  var movedObjects = false;

  for (var i = 0; i < elements.length; i++) {
    var element = elements[i];
    var targetLayer;
    try {
      targetLayer = doc.layers.getByName(element.name);
    } catch (e) {
      targetLayer = null;
    }

    if (targetLayer) {
      for (var j = 0; j < targetLayer.pageItems.length; j++) {
        var item = targetLayer.pageItems[j];
        var typeName = item.typename;
        // PathItem / GroupItem / CompoundPathItem のみ移動を実行
        if (
          typeName === "PathItem" ||
          typeName === "GroupItem" ||
          typeName === "CompoundPathItem"
        ) {
          var xPosition = artboardRect[0] + mmToPt(element.x);
          var yPosition = artboardRect[1] - mmToPt(element.y);

          // 中心合わせをする場合
          item.position = [
            xPosition - item.width / 2,
            yPosition + item.height / 2,
          ];
          movedObjects = true;
        }
      }
    } else {
      // 対象レイヤーがなかった場合
      if (showAlerts) {
        alert("「" + element.name + "」レイヤーが存在しませんでした。");
      }
    }
  }

  if (!movedObjects && showAlerts) {
    alert("移動対象のオブジェクトが見つかりませんでした。");
  }
  // ファイル保存前にすべてのオブジェクトを選択解除
  doc.selection = null;
}

// =============== ダイアログ生成関数群 ===============

/**
 * ダイアログを作成し、各種パスやチェックボックスの設定を取得する
 * @param {string} initialCSVPath
 * @param {string} initialAIPath
 * @param {string} initialSavePath
 * @param {string} initialTemplateFolderPath
 * @returns {Object} ユーザー選択の設定情報 or null(キャンセル)
 */
function createUnifiedDialog(
  initialCSVPath,
  initialAIPath,
  initialSavePath,
  initialTemplateFolderPath
) {
  var dialog = new Window("dialog", "テキスト・リンクを置き換える方法を選択");
  dialog.orientation = "column";
  dialog.alignChildren = "left"; // 全体の整列を左揃え設定

  // === チェックボックスグループ ===
  var checkboxGroup = dialog.add("group");
  checkboxGroup.orientation = "column";
  checkboxGroup.alignment = "left";
  checkboxGroup.alignChildren = "left";

  var showAlertsCheckbox = checkboxGroup.add(
    "checkbox",
    undefined,
    "ステップバイ処理（連続処理の場合はチェックを外す）"
  );
  showAlertsCheckbox.value = false;

  var closeFileAfterProcessingCheckbox = checkboxGroup.add(
    "checkbox",
    undefined,
    "処理終了後ファイルを閉じる"
  );
  closeFileAfterProcessingCheckbox.value = false;

  var replaceTextCheckbox = checkboxGroup.add(
    "checkbox",
    undefined,
    "テキストを置換る"
  );
  replaceTextCheckbox.value = true;

  var skipBlankReplaceCheckbox = checkboxGroup.add(
    "checkbox",
    undefined,
    "ブランクテキストの置換えはスキップ"
  );
  skipBlankReplaceCheckbox.value = true;

  var replaceLinksCheckbox = checkboxGroup.add(
    "checkbox",
    undefined,
    "リンクを差替える"
  );
  replaceLinksCheckbox.value = true;

  var overflowFixCheckbox = checkboxGroup.add(
    "checkbox",
    undefined,
    "オーバーフローを解消"
  );
  overflowFixCheckbox.value = true;

  var adjustPositionCheckbox = checkboxGroup.add(
    "checkbox",
    undefined,
    "位置移動する"
  );
  adjustPositionCheckbox.value = true;

  var excludeLockedLayersCheckbox = checkboxGroup.add(
    "checkbox",
    undefined,
    "ロック済みレイヤーを除外"
  );
  excludeLockedLayersCheckbox.value = true;

  var excludeHiddenLayersCheckbox = checkboxGroup.add(
    "checkbox",
    undefined,
    "非表示レイヤーを除外"
  );
  excludeHiddenLayersCheckbox.value = true;

  // 空白行を挿入
  checkboxGroup.add("statictext", undefined, "");

  // === 'テンプレ情報をCSVから取得' チェックボックス ===
  var useTemplateFromCSVCheckbox = dialog.add(
    "checkbox",
    undefined,
    "テンプレ.ai情報をCSVから取得"
  );
  useTemplateFromCSVCheckbox.value = true;

  // === テンプレのあるフォルダを選択 ===
  var fileGroupTemplateFolder = dialog.add("group");
  fileGroupTemplateFolder.orientation = "column";
  fileGroupTemplateFolder.alignChildren = "left";
  fileGroupTemplateFolder.add(
    "statictext",
    undefined,
    "テンプレのあるフォルダを選択:"
  );
  var templateFolderPathGroup = fileGroupTemplateFolder.add("group");
  templateFolderPathGroup.orientation = "row";
  templateFolderPathGroup.alignment = "left";

  var templateFolderPath = templateFolderPathGroup.add(
    "edittext",
    undefined,
    initialTemplateFolderPath
  );
  templateFolderPath.preferredSize = [300, 20];
  var browseTemplateFolderButton = templateFolderPathGroup.add(
    "button",
    undefined,
    "参照..."
  );
  browseTemplateFolderButton.onClick = function () {
    var templateFolder = Folder.selectDialog("テンプレのあるフォルダを選択");
    if (templateFolder) {
      templateFolderPath.text = templateFolder.fsName;
    }
  };

  // === '1ファイル毎にテンプレaiを選択' チェックボックス ===
  var selectAIperFileCheckbox = dialog.add(
    "checkbox",
    undefined,
    "1ファイル毎にテンプレaiを選択"
  );
  selectAIperFileCheckbox.value = false;

  // === テンプレートaiを選択 ===
  var fileGroupAI = dialog.add("group");
  fileGroupAI.orientation = "column";
  fileGroupAI.alignChildren = "left";
  fileGroupAI.add("statictext", undefined, "テンプレートaiを選択:");
  var aiPathGroup = fileGroupAI.add("group");
  aiPathGroup.orientation = "row";
  var aiFilePath = aiPathGroup.add("edittext", undefined, initialAIPath);
  aiFilePath.preferredSize = [300, 20];
  var browseAIButton = aiPathGroup.add("button", undefined, "参照...");
  browseAIButton.onClick = function () {
    var aiFile = File.openDialog("テンプレaiを選択");
    if (aiFile) {
      aiFilePath.text = aiFile.fsName;
    }
  };

  // チェックボックスの状態によってUIを切り替える
  function updateUI() {
    var isEnabled = !useTemplateFromCSVCheckbox.value;
    selectAIperFileCheckbox.enabled = isEnabled;
    aiFilePath.enabled = isEnabled;
    browseAIButton.enabled = isEnabled;

    templateFolderPath.enabled = !isEnabled;
    browseTemplateFolderButton.enabled = !isEnabled;
  }
  // 初期呼び出し
  updateUI();
  useTemplateFromCSVCheckbox.onClick = updateUI;

  // === 処理結果の保存先フォルダを選択 ===
  var fileGroupSave = dialog.add("group");
  fileGroupSave.orientation = "column";
  fileGroupSave.alignChildren = "left";
  fileGroupSave.add("statictext", undefined, "処理後の保存先フォルダを選択:");
  var savePathGroup = fileGroupSave.add("group");
  savePathGroup.orientation = "row";
  var saveFolderPath = savePathGroup.add(
    "edittext",
    undefined,
    initialSavePath
  );
  saveFolderPath.preferredSize = [300, 20];
  var browseSaveButton = savePathGroup.add("button", undefined, "参照...");
  browseSaveButton.onClick = function () {
    var saveFolder = Folder.selectDialog("処理後の保存先フォルダを選択");
    if (saveFolder) {
      saveFolderPath.text = saveFolder.fsName;
    }
  };

  var createNewFolderCheckbox = fileGroupSave.add(
    "checkbox",
    undefined,
    "新規に保存フォルダを作る"
  );
  createNewFolderCheckbox.value = true;

  // === CSVファイルを選択 ===
  var fileGroupCSV = dialog.add("group");
  fileGroupCSV.orientation = "column";
  fileGroupCSV.alignChildren = "left";
  fileGroupCSV.add("statictext", undefined, "CSVファイルを選択:");
  var csvPathGroup = fileGroupCSV.add("group");
  csvPathGroup.orientation = "row";
  var csvFilePath = csvPathGroup.add("edittext", undefined, initialCSVPath);
  csvFilePath.preferredSize = [300, 20];
  var browseCSVButton = csvPathGroup.add("button", undefined, "参照...");
  browseCSVButton.onClick = function () {
    var csvFile = File.openDialog("CSVファイルを選択してください");
    if (csvFile) {
      csvFilePath.text = csvFile.fsName;
    }
  };

  // === OK/キャンセルボタン ===
  var buttonGroup = dialog.add("group");
  buttonGroup.orientation = "row";
  buttonGroup.alignment = "center";

  buttonGroup.add("button", undefined, "キャンセル", {
    name: "cancel",
  }).onClick = function () {
    dialog.close(0);
  };

  buttonGroup.add("button", undefined, "OK", { name: "ok" }).onClick =
    function () {
      // 入力チェック
      if (csvFilePath.text === "" || saveFolderPath.text === "") {
        alert("CSVファイル、および保存先フォルダを選択してください。");
        return;
      }
      // テンプレCSV使用しない場合はAIファイル必須
      if (!useTemplateFromCSVCheckbox.value && aiFilePath.text === "") {
        alert("テンプレaiを選択してください。");
        return;
      }
      dialog.close(1);
    };

  var result = dialog.show();
  // OKボタンで閉じた場合
  if (result === 1) {
    return {
      excludeLockedLayers: excludeLockedLayersCheckbox.value,
      excludeHiddenLayers: excludeHiddenLayersCheckbox.value,
      showAlerts: showAlertsCheckbox.value,
      closeFileAfterProcessing: closeFileAfterProcessingCheckbox.value,
      csvFile: csvFilePath.text,
      aiFile: aiFilePath.text,
      savePath: saveFolderPath.text,
      replaceText: replaceTextCheckbox.value,
      skipBlankReplace: skipBlankReplaceCheckbox.value,
      replaceLinks: replaceLinksCheckbox.value,
      overflowFix: overflowFixCheckbox.value,
      adjustPosition: adjustPositionCheckbox.value,
      selectAIperFile: selectAIperFileCheckbox.value,
      createNewSaveFolder: createNewFolderCheckbox.value,
      useTemplateFromCSV: useTemplateFromCSVCheckbox.value,
      templateFolder: templateFolderPath.text,
    };
  }
  return null;
}

/**
 * 新たに追加: 検索置換するAIファイル選択専用のダイアログを表示
 * @param {string} initialAIPath
 * @returns {Object|null}
 */
function selectAIFileDialog(initialAIPath) {
  var dialog = new Window(
    "dialog",
    "検索置換する「イラレファイル」を選択してください"
  );
  dialog.orientation = "column";

  var aiPathGroup = dialog.add("group");
  aiPathGroup.orientation = "row";
  aiPathGroup.alignment = "left";

  var aiFilePath = aiPathGroup.add("edittext", undefined, initialAIPath);
  aiFilePath.preferredSize = [300, 20];

  var browseAIButton = aiPathGroup.add("button", undefined, "参照...");
  browseAIButton.onClick = function () {
    var aiFile = File.openDialog("検索置換するファイルを選択してください");
    if (aiFile) {
      aiFilePath.text = aiFile.fsName;
      // UIの都合でフォーカス移動させるだけ
      aiFilePath.active = true;
      aiFilePath.active = false;
    }
  };

  var buttonGroup = dialog.add("group");
  buttonGroup.orientation = "row";
  buttonGroup.alignment = "center";

  var cancelButton = buttonGroup.add("button", undefined, "キャンセル", {
    name: "cancel",
  });
  var okButton = buttonGroup.add("button", undefined, "OK", { name: "ok" });

  okButton.onClick = function () {
    if (aiFilePath.text === "") {
      alert("AIファイルを選択してください。");
    } else {
      dialog.close(1);
    }
  };
  cancelButton.onClick = function () {
    dialog.close(0);
  };

  var result = dialog.show();
  if (result == 1) {
    return { aiFile: aiFilePath.text };
  } else {
    return null;
  }
}

// =============== メイン処理 ===============

/**
 * メイン処理関数
 */
function main() {
  // 初期パス設定
  var initialCSVPath = "CSVファイルパスを設定";
  var initialAIPath = "aiファイルパスを設定";
  var initialSavePath = "保存先パスを設定";
  var initialTemplateFolderPath = "Templateフォルダパスを設定";

  // ダイアログを開く
  var settings = createUnifiedDialog(
    initialCSVPath,
    initialAIPath,
    initialSavePath,
    initialTemplateFolderPath
  );
  if (settings == null) {
    alert("操作がキャンセルされました。");
    return;
  }

  var showAlerts = settings.showAlerts;
  var closeFileAfterProcessing = settings.closeFileAfterProcessing;
  var csvFile = new File(settings.csvFile);
  var savePath = settings.savePath;
  var excludeLockedLayers = settings.excludeLockedLayers;
  var excludeHiddenLayers = settings.excludeHiddenLayers;
  var selectAIperFile = settings.selectAIperFile;
  var createNewSaveFolder = settings.createNewSaveFolder;
  var useTemplateFromCSV = settings.useTemplateFromCSV;
  var templateFolderPathText = settings.templateFolder;
  var data = readCSV(csvFile);

  // CSVが正しく読み込めない場合
  if (!data || data.length === 0) {
    alert("CSVファイルのデータが無効です。");
    return;
  }

  // CSVの1行目はヘッダ行
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row) {
      if (showAlerts) alert("操作がキャンセルされました。");
      break;
    }

    // テンプレAIファイル名、フォルダ名、ファイル名を取り出し
    var templateFileName = null;
    var folderName = null;
    var fileName = null;

    if (useTemplateFromCSV) {
      // CSV列: [0]:テンプレAIファイル名, [1]:フォルダ名, [2]:保存用ファイル名
      templateFileName = row[0];
      folderName = row[1];
      fileName = row[2];
    } else {
      // CSV列: [0]はテンプレ名(不要), [1]:フォルダ名, [2]:保存用ファイル名
      folderName = row[1];
      fileName = row[2];
    }

    // 保存時のファイル名
    var saveFileName = fileName + ".ai";

    // 保存先のフォルダとファイルオブジェクトを作成
    var saveFile;
    if (createNewSaveFolder) {
      var newFolder = new Folder(savePath + "/" + folderName);
      if (!newFolder.exists) {
        newFolder.create();
      }
      saveFile = new File(newFolder.fsName + "/" + saveFileName);
    } else {
      saveFile = new File(savePath + "/" + saveFileName);
    }

    // テンプレCSV使用の有無で、開くAIファイルを決定
    var currentAIFile = settings.aiFile;
    if (useTemplateFromCSV) {
      // テンプレAIをCSVから取得
      var aiFile = new File(templateFolderPathText + "/" + templateFileName);
      if (!aiFile.exists) {
        alert("テンプレートファイルが存在しません: " + aiFile.fsName);
        continue;
      }
      var doc = app.open(aiFile);
      if (!doc) {
        alert(
          "テンプレートファイルを開くことができませんでした: " + aiFile.fsName
        );
        continue;
      }
    } else {
      // 従来通り
      if (selectAIperFile && i > 1) {
        var aiSettings = selectAIFileDialog(initialAIPath);
        if (aiSettings == null) {
          if (showAlerts) alert("操作がキャンセルされました。");
          break;
        }
        currentAIFile = aiSettings.aiFile;
      }
      var aiFile = new File(currentAIFile);
      if (!aiFile.exists) {
        alert("AIファイルが存在しません: " + aiFile.fsName);
        continue;
      }
      var doc = app.open(aiFile);
      if (!doc) {
        alert("AIファイルを開くことができませんでした: " + aiFile.fsName);
        continue;
      }
    }

    // リンク置換
    if (settings.replaceLinks) {
      var linkFindTexts = [];
      var linkReplaceTexts = [];
      // CSVヘッダ row[0][j] に "LINK_" で始まる列を探す
      for (var j = 3; j < data[0].length; j++) {
        if (data[0][j] && data[0][j].indexOf("LINK_") === 0) {
          linkFindTexts.push(data[0][j].replace("LINK_", ""));
          linkReplaceTexts.push(row[j]);
        }
      }
      sortTextsByLengthAsc(linkFindTexts, linkReplaceTexts);
      processLinks(
        doc,
        linkFindTexts,
        linkReplaceTexts,
        excludeLockedLayers,
        excludeHiddenLayers
      );
    }

    // テキスト置換
    if (settings.replaceText) {
      var textFindTexts = [];
      var textReplaceTexts = [];
      // CSVヘッダ row[0][j] に "TEXT_" で始まる列を探す
      for (var j = 3; j < data[0].length; j++) {
        if (data[0][j] && data[0][j].indexOf("TEXT_") === 0) {
          textFindTexts.push(data[0][j].replace("TEXT_", ""));
          textReplaceTexts.push(row[j]);
        }
      }
      sortTextsByLengthAsc(textFindTexts, textReplaceTexts);
      processCSVRow(
        doc,
        textFindTexts,
        textReplaceTexts,
        excludeLockedLayers,
        excludeHiddenLayers,
        settings.skipBlankReplace
      );
    }

    // オーバーフロー解消
    if (settings.overflowFix) {
      // showAlertsがオフの場合も、confirmダイアログは出ない
      if (!showAlerts || confirm("オーバーフロー解消処理を開始します。")) {
        overflowFix(doc, excludeLockedLayers, excludeHiddenLayers, showAlerts);
      }
    }

    // 位置修正
    if (settings.adjustPosition) {
      adjustPosition(doc, showAlerts);
    }

    // 保存
    var saveOptions = new IllustratorSaveOptions();
    // （元コードのコメント状態を維持）
    // saveOptions.pdfCompatible = true;
    saveOptions.pdfCompatible = false;
    saveOptions.embedICCProfile = false;
    saveOptions.compressed = true;
    try {
      doc.saveAs(saveFile, saveOptions);
    } catch (e) {
      alert("ファイルを保存中にエラーが発生しました: " + saveFileName);
    }

    if (showAlerts) {
      alert("次のファイル名で保存し、ファイルを閉じます。\n" + saveFileName);
    }

    if (closeFileAfterProcessing) {
      doc.close(SaveOptions.DONOTSAVECHANGES);
    }
  }

  // 処理完了後
  if (!showAlerts) {
    alert("確認\nすべての処理が完了しました。");
  }
  if (showAlerts) {
    alert("すべての処理が完了しました。");
  }
}

// =============== スクリプトの実行 ===============

(function () {
  // メイン処理を呼び出し
  main();
})();
