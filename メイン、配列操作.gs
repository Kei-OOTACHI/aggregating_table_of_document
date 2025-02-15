function dummyReader() {
  docHandler.deleteSectionAfterHeading();
  const docId = DocumentApp.getActiveDocument().getId();
  const tableData = TableExtractor.extractTableDataWithChildTables(docId)//ドキュメントから表データを抽出して連想配列に変換
  let trimedTableData = TableProcessor.transformDataTo2DArray(tableData);//連想配列を2次元配列に変換
  trimedTableData = extractDateFromHeadingInThirdColumn(trimedTableData);
  trimedTableData = ArrayManipulator.sortArrayByColumns(trimedTableData, hasHeader = true);
  const columnsToMerge = [2, 3, 4, 5];
  trimedTableData = ArrayManipulator.mergeColumns(trimedTableData, columnsToMerge, "／");
  Logger.log(trimedTableData);
  TableInserter.addTableToEndOfDocumentWithColumnWidthsAndFontSize(docId, trimedTableData);
}

function extractDateFromHeadingInThirdColumn(data) {
  return data.map(row => {
    const datePattern = /\d{1,2}\/\d{1,2}/; // 日付形式のパターン（00/00など）
    const match = row[2].match(datePattern);
    row[2] = match ? match[0] : row[2]; // 日付が見つかればそれを返し、見つからなければ元の3列目をそのまま残す
    return row;
  });
}

// JSON⇒2次元配列に変換数するオブジェクト
const TableProcessor = {
  // グローバル変数をオブジェクトのプロパティとして定義
  COLUMN_ORDER: ["department", "itemName", "heading", "time", "summary", "quantity", "usageLocation", "purpose"],
  HEADER_ROW: ["貸出部署", "備品名", "日付", "時間", "活動概要", "個数", "使用場所", "使用目的・備考"],

  // データを2次元配列に変換する関数
  transformDataTo2DArray(dataArray) {
    const result = [];

    // 見出し行を追加
    result.push(this.HEADER_ROW);

    dataArray.forEach(tableData => {
      tableData.forEach(entry => {
        // 活動内容が「備品借用」「備品返却」でない行をスキップ
        if (entry.summary != "備品借用" && entry.summary != "備品返却") return;

        // 活動詳細に子表がある場合のみ処理
        if (entry.details && entry.details.childTable) {
          entry.details.childTable.forEach(childRow => {
            // 備品名が「備品名」の行をスキップ
            if (childRow.itemName === "備品名") return;

            const row = [];

            // COLUMN_ORDERに基づいて各列のデータを追加
            this.COLUMN_ORDER.forEach(column => {
              if (column in childRow) {   //ここ実はfor in ではなくて if in. きもい。array.includes()と同じ意味らしい。
                row.push(childRow[column] || "");
              } else {
                row.push(entry[column] || "");
              }
            });

            // 新しい行を追加
            result.push(row);
          });
        }
      });
    });

    return result;
  }
};

// 2次元配列を並べ替えたり結合したりするオブジェクト
const ArrayManipulator = {
  // 配列を列ごとに並べ替える関数
  sortArrayByColumns(array, hasHeader = false) {
    let header = [];
    let dataToSort = array;

    // 見出し行がある場合は、最初の行を除外
    if (hasHeader) {
      header = array[0]; // 最初の行を見出し行として保持
      dataToSort = array.slice(1); // 見出し行以外を並び替え対象とする
    }

    // 見出し行以外を並び替え
    dataToSort.sort(function (a, b) {
      for (let i = 0; i < a.length; i++) {
        if (a[i] < b[i]) return -1; // aがbより小さい場合、aが先
        if (a[i] > b[i]) return 1;  // aがbより大きい場合、bが先
      }
      return 0; // すべての列が等しい場合、順序を変更しない
    });

    // 見出し行がある場合は、再び最初に追加
    if (hasHeader) {
      dataToSort.unshift(header);
    }

    return dataToSort; // 並び替え後の配列を返す
  },

  // 指定された列を結合する関数
  mergeColumns(array, columnsToMerge, separator) {
    const result = [];

    for (let i = 0; i < array.length; i++) {
      const row = array[i];
      const newRow = [];

      // 指定された列番号のデータを結合
      const mergedData = columnsToMerge
        .map(colIndex => row[colIndex])
        .filter(data => data !== undefined) // undefinedを除外
        .join(separator); // separatorで結合

      // 新しい行に他の列のデータを追加
      for (let k = 0; k < row.length; k++) {
        if (!columnsToMerge.includes(k)) {
          newRow.push(row[k]); // 結合対象でない列をそのまま追加
        } else if (k === columnsToMerge[0]) {
          // 結合対象の最初の列に結合したデータを追加
          newRow.push(mergedData);
        }
      }

      result.push(newRow); // 新しい行を結果に追加
    }

    return result;
  },

  // 配列を指定された列数に成型する関数
  reshapeToNColumns(array, n) {
    const result = [];

    for (let i = 0; i < array.length; i++) {
      const firstColumns = array[i].slice(0, n); // n列目までの要素
      const restColumns = array[i].slice(n); // n列目以降の要素を取得
      let combinedColumns = "";
      // n列目以降を結合（スペースで区切る）
      if (Array.isArray(restColumns)) combinedColumns = restColumns.join(" ");
      // 新しい行として、n列目までの要素と結合した2列目を結果に追加
      result.push([...firstColumns, combinedColumns]);
    }

    return result; // 成型後の2次元配列を返す
  }
};

