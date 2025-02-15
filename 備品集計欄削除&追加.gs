const docHandler = {
  targetText: " ♢ 使用備品", // 検索したい特定の文字列

  deleteSectionAfterHeading: function () {
    const body = DocumentApp.getActiveDocument().getBody();
    let deleteMode = false;

    // 順方向にループ
    for (let i = 0; i < body.getNumChildren(); i++) {
      const element = body.getChild(i);

      // 要素がPARAGRAPHであることを確認
      if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const paragraph = element.asParagraph();

        // 見出し1であるかどうかを確認
        if (paragraph.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
          const text = paragraph.getText();

          // 見出し1に特定の文字列が含まれている場合、削除モードをオンにする
          if (text.includes(this.targetText)) {
            deleteMode = true;
            continue; // 見出し1自体は削除しないので処理をスキップ
          } else if (deleteMode) {
            // 次の見出し1が見つかった場合、削除モードをオフにする
            deleteMode = false;
          }
        }
      }

      // 削除モードがオンの場合
      if (deleteMode) {
        if (i === body.getNumChildren() - 1) {
          // 最後の段落の場合、削除せず空白に置き換える
          element.asParagraph().setText("------------------------------------------------------------------------------------------");
        } else {
          try {
            // 通常の要素は削除
            body.removeChild(element);
            i--; // 削除した要素によりインデックスがずれるのを防ぐ
          } catch (e) {
            Logger.log('要素の削除中にエラーが発生しました: ' + e.message);
          }
        }
      }
    }
  }
};

const TableInserter = {
  // 列ごとの幅を指定する変数（単位: ポイント）
  COLUMN_WIDTHS: [70, 100, 300, 110, 205], // 各列の幅を指定

  // フォントサイズを指定する変数（単位: ポイント）
  FONT_SIZE: 9,

  // 1回のバッチで処理する行数
  BATCH_SIZE: 20,

  // 2次元配列をGoogleドキュメントの末尾に表として挿入する関数（バッチ処理）
  addTableToEndOfDocumentWithColumnWidthsAndFontSize(docId, data) {
    let doc = DocumentApp.openById(docId);
    let body = doc.getBody();
    let table = body.appendTable(); // 最初にテーブルを作成

    // データをバッチごとに処理
    for (let i = 0; i < data.length; i += this.BATCH_SIZE) {
      const batch = data.slice(i, i + this.BATCH_SIZE);

      // バッチ内のデータをテーブルに追加
      batch.forEach((rowData, rowIndex) => {
        const row = table.appendTableRow();
        rowData.forEach((cellData, colIndex) => {
          const cell = row.appendTableCell(cellData.toString());

          // 列幅を設定（列ごとの幅が定義されている場合）
          if (colIndex < this.COLUMN_WIDTHS.length) {
            const width = this.COLUMN_WIDTHS[colIndex];
            cell.setWidth(width);
          }

          // フォントサイズを設定
          const text = cell.getChild(0).asText();
          text.setFontSize(this.FONT_SIZE);
        });
      });

      // バッチ処理後にドキュメントを保存してクローズし、再度開き直す
      if (i + this.BATCH_SIZE < data.length) {
        doc.saveAndClose();
        Utilities.sleep(1000); // 1秒待機（負荷を軽減するため）
        doc = DocumentApp.openById(docId); // ドキュメントを再度開き直し
        body = doc.getBody();
        table = body.getTables().pop(); // 既存のテーブルを取得
      }
    }

    // 最後にドキュメントを保存して終了
    doc.saveAndClose();
  }
};

