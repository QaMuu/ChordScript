const ASS = SpreadsheetApp.getActiveSpreadsheet();
const sheetChord = ASS.getSheetByName("コード");
const sheetBassTab = ASS.getSheetByName("ベースタブ譜");
const sheetBassFret = ASS.getSheetByName("ベースフレット");

let dicBassFret: {[key: string]: Array<number>}

function onOpen() {
  const customMenu = SpreadsheetApp.getUi();
  customMenu
    .createMenu("オリジナル")
    .addItem("コードシートに行を設定", "setRowSize")
    .addItem("コード -> ベースタブ譜", "changeChord_To_BassTab")
    .addToUi();
}

function changeChord_To_BassTab() {
  getDicBassFret();
}

function getDicBassFret() {
  dicBassFret = {};
  const rangeBassFret: string = createRange("A", 1, "E", 53);
  const tempJson: any = sheetBassFret?.getRange(rangeBassFret).getValues();

  for (let index = 0; index < tempJson.length; index++) {
    const targetRow: any = tempJson[index];
    const setCodeRoot: string = String(targetRow[0]);
    const arrayFret: Array<number> = Array<number>(
      Number(targetRow[1]),
      Number(targetRow[2]),
      Number(targetRow[3]),
      Number(targetRow[4])
    );
    dicBassFret[setCodeRoot] = arrayFret;
  }

  //console.log(dicBassFret);
  //console.log(dicBassFret['E1'])
}

function setRowSize() {
  let countChangeRow = 300;

  for (let index = 2; index < countChangeRow; index += 3) {
    sheetChord?.setRowHeight(index, 30);
    sheetChord?.setRowHeight(index + 1, 10);
    sheetChord?.setRowHeight(index + 2, 40);
  }
}

function createRange(
  startCol: string,
  startRow: number,
  endCol: string,
  endRow: number
): string {
  return startCol + startRow + ":" + endCol + endRow;
}
