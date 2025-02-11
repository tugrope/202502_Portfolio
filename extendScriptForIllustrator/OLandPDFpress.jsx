try {
  // アクティブなドキュメントが存在するか確認
  if (app.documents.length > 0) {
    // ドキュメントがなくなるまでループ処理
    while (app.documents.length > 0) {
      var doc = app.documents[0]; // 常に最初のドキュメントを処理

      // 現在のファイル名を取得し、"ol"を付加した新しい名前を生成
      var originalFilePath = doc.fullName.fsName; // 元のファイルパスを取得
      var newFilePath = originalFilePath.replace(/\.ai$/, "") + "ol.ai"; // ファイル名に"ol"を追加

      // Illustrator保存オプションを設定
      var saveOptions = new IllustratorSaveOptions();
      saveOptions.pdfCompatible = true; // PDF互換ファイルを有効に設定

      // 新しいファイルとして保存し、処理の完了を確認する
      doc.saveAs(new File(newFilePath), saveOptions);

      // 保存が完了するまでファイルの存在確認
      while (!File(newFilePath).exists) {
        $.sleep(100);
      }

      // #1: 指定されたアクション "ol_Save_Packing01" を実行
      var actionSetName = "forExtendScript";
      var actionName = "ol_Save_Packing01";
      app.doScript(actionName, actionSetName);

      $.sleep(500);

      // #2: アクティブドキュメントを上書き保存
      var overwriteSaveOptions = new IllustratorSaveOptions();
      overwriteSaveOptions.pdfCompatible = true;
      overwriteSaveOptions.embedICCProfile = false;
      overwriteSaveOptions.compressed = true;

      doc.saveAs(doc.fullName, overwriteSaveOptions);

      $.sleep(500);

      // #4: "PressSave_Close" アクションを実行
      var pdfActionName = "PressSave_Close";
      app.doScript(pdfActionName, actionSetName);

      $.sleep(500);
    }

    // 全ての処理が完了したらアラートを表示
    alert("OL処理とPressPDFの書出し完了");
  } else {
    alert("アクティブなドキュメントがありません。");
  }
} catch (e) {
  alert("エラーが発生しました: " + e.message);
}
