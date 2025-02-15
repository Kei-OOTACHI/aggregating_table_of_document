const TABLE_ELEMENT = DocumentApp.ElementType.TABLE;
const TEXT_ELEMENT = DocumentApp.ElementType.TEXT;
const PARAGRAPH_ELEMENT = DocumentApp.ElementType.PARAGRAPH;
const LIST_ITEM_ELEMENT = DocumentApp.ElementType.LIST_ITEM;
const HEADING1 = DocumentApp.ParagraphHeading.HEADING1;
const HEADING2 = DocumentApp.ParagraphHeading.HEADING2;

// 親表の列番号とオブジェクトのプロパティ名の対応関係
const TableExtractor = {
  // 親表の列番号とオブジェクトのプロパティ名の対応関係
  PARENT_COLUMN_MAPPING: {
    0: 'time',
    1: 'assignee',
    2: 'location',
    3: 'summary',
    4: 'details', // 活動詳細 (子表がある可能性あり)
    5: 'department',
    6: 'status'
  },

  // 子表の列番号とオブジェクトのプロパティ名の対応関係
  CHILD_COLUMN_MAPPING: {
    0: 'itemName',
    1: 'quantity',
    2: 'usageLocation',
    3: 'purpose'
  },

  // ドキュメントからテーブルデータを抽出する関数
  extractTableDataWithChildTables(docId) {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    let tableData = [];
    let currentHeading = null; // 現在の見出しを追跡

    // ドキュメント内の全ての要素を処理
    const numChildren = body.getNumChildren();
    for (let i = 0; i < numChildren; i++) {
      const element = body.getChild(i);
      const eType = element.getType();

      if (eType == TABLE_ELEMENT) {
        const table = element.asTable();
        // テーブルが少なくとも1行以上あるかを確認
        if (table.getNumRows() > 0 && table.getRow(0).getNumCells() === Object.keys(this.PARENT_COLUMN_MAPPING).length) {
          const parentTableData = this.processParentTable(table, currentHeading);
          tableData.push(parentTableData);
        }
      } else if (eType == PARAGRAPH_ELEMENT) {
        const headingType = element.asParagraph().getHeading();
        if (headingType == HEADING1 || headingType == HEADING2) {
          const headingText = element.asParagraph().getText();
          currentHeading = headingText; // 現在の見出しを更新
        }
      }
    }

    return tableData;
  },

  // 親表を処理する関数
  processParentTable(table, heading) {
    const numRows = table.getNumRows();
    let parentTableData = [];

    for (let rowIndex = 0; rowIndex < numRows; rowIndex++) {
      const row = table.getRow(rowIndex);
      let rowData = {};

      // 見出しデータを含める
      if (heading) {
        rowData.heading = heading;
      }

      // 各列のデータを処理
      for (let columnIndex = 0; columnIndex < row.getNumCells(); columnIndex++) {
        const columnKey = this.PARENT_COLUMN_MAPPING[columnIndex];
        if (columnKey === 'details') {
          // 5列目（活動詳細）の処理。子表がある場合はそのデータも含める
          rowData[columnKey] = this.processCellWithChildTable(row.getCell(columnIndex));
        } else {
          rowData[columnKey] = this.processCell(row.getCell(columnIndex));
        }
      }

      parentTableData.push(rowData);
    }

    return parentTableData;
  },

  // セルのデータを処理する関数（子表がない場合）
  processCell(cell) {
    const numChildren = cell.getNumChildren();
    let cellData = [];

    for (let childIndex = 0; childIndex < numChildren; childIndex++) {
      const childElement = cell.getChild(childIndex);
      const childElementData = this.processChildCellText(childElement);
      cellData.push(childElementData);
    }

    return cellData.length === 1 ? cellData[0] : cellData.join(" ");
  },

  // セルのデータを処理し、子表があればそのデータも含める関数
  processCellWithChildTable(cell) {
    const numChildren = cell.getNumChildren();
    let cellData = { text: "", childTable: null };

    for (let childIndex = 0; childIndex < numChildren; childIndex++) {
      const childElement = cell.getChild(childIndex);

      if (childElement.getType() == TABLE_ELEMENT) {
        const childTable = childElement.asTable();
        const childTableData = this.processChildTable(childTable);
        cellData.childTable = childTableData;

      } else {
        cellData.text += this.processChildCellText(childElement) + " ";
      }
    }

    cellData.text = cellData.text.trim(); // トリムして余分なスペースを削除

    return cellData;
  },

  // 子表を処理する関数
  processChildTable(table) {
    const numRows = table.getNumRows();
    let childTableData = [];

    for (let rowIndex = 0; rowIndex < numRows; rowIndex++) {
      const row = table.getRow(rowIndex);
      let rowData = {};

      // 各列のデータを処理
      for (let columnIndex = 0; columnIndex < row.getNumCells(); columnIndex++) {
        const columnKey = this.CHILD_COLUMN_MAPPING[columnIndex];
        rowData[columnKey] = this.processCell(row.getCell(columnIndex));
      }

      childTableData.push(rowData);
    }

    return childTableData;
  },

  // 子表以外のセル内テキストを処理する関数
  processChildCellText(childElement) {
    let text = "";

    if (childElement.getType() == TEXT_ELEMENT) {
      text = childElement.asText().getText();
    } else if (childElement.getType() == PARAGRAPH_ELEMENT) {
      text = childElement.asParagraph().getText();
    } else if (childElement.getType() == LIST_ITEM_ELEMENT) {
      text = childElement.asListItem().getText();
    } else {
      Logger.log("Unsupported element type found: " + childElement.getType());
    }

    return text;
  }
};
