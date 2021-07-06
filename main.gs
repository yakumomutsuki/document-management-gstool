// Click service menu and enable "Drive API"

const TARGET_ID = 'xxxxxxxxxx';
const SHEET_NAME = '構成管理一覧シート';

const COLORS = {
  YELLOW: '#FFFF00'
}

/**
 * IDをもとに、配下のファイル一覧を取得します
 * @param {string} targetId
 */
const getFilesById = (targetId) => {
  return DriveApp.getFolderById(targetId).getFiles();
}

/**
 * ファイルオブジェクトを元に、配下ファイルの文字列一覧を取得します
 * @param {files} files
 */
const getFileNames = (files) => {
  const result = [];
  while(files.hasNext()){
    const file = files.next();
    result.push(file.getName());
  }

  // 昇順に並び替えを行います
  return result.sort((a,b) => {
    if (a < b) return -1;
    if (a > b) return 1;
    return 0;
  });
}

/**
 * IDをもとに、配下のフォルダ一覧を取得します
 * @param {string} targetId
 */
const getFoldersById = (targetId) => {
  return DriveApp.getFolderById(targetId).getFolders();
}

/**
 * フォルダオブジェクトを元に、配下フォルダのオブジェクト一覧を取得します
 * @param {folder} folders
 */
const getFolderItems = (folders) => {
  const result = [];
  while(folders.hasNext()){
    const folder = folders.next();
    result.push({
      name: folder.getName(),
      id: folder.getId(),
    });
  }

  // 昇順に並び替えを行います
  return result.sort((a,b) => {
    if (a.name < b.name) return -1;
    if (a.name > b.name) return 1;
    return 0;
  });
}

/**
 * 指定されたIDを元に、再帰処理によるフォルダ、ファイルを含む{フォルダ名、ファイル名}を含む
 * オブジェクトを生成します
 * 
 * @return Result<Result>
 */
const recursiveSearchDrive = (targetId) => {
  const folderItems = getFolderItems(getFoldersById(targetId));
  const files = getFileNames(getFilesById(targetId));

  /**
   * Type Result<T> = {
   *   folderName: string,
   *   folders: Array<T>,
   *   files: Array<string>
   * }
   */
  const result = {
    folderName: '',
    folders: [],
    files: [],
  };

  const folderName = DriveApp.getFolderById(targetId).getName();
  if (folderName || files) {
    result.folderName = folderName;
    result.files = files;
  }

  folderItems.forEach(it => {
    const item = recursiveSearchDrive(it.id);
    result.folders.push(item);
  });

  return result;
}

/**
 * シートを作成します
 * @param name
 */
const createSheet = (name) => {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  // シートが存在した場合は削除します
  if (spreadSheet.getSheetByName(name)){
    spreadSheet.deleteSheet(spreadSheet.getSheetByName(name));
  }

  const sheet = spreadSheet.insertSheet(name);
  return sheet;
}


let row = 1;
let col = 1;

/**
 * Driveのフォルダ構成、ファイル一覧をスプレッドシートへ転記します
 * 再帰処理を行うため、何度か呼び出されます
 * 
 * @param {sheet} sheet
 * @param {object} searchResult
 * @param {int} startCol
 * @param {int} startRow
 */
const transcriptionToSheetOfDriveInfo = (sheet, searchResult, startCol, startRow) => {
  col = startCol;
  const {folderName, folders, files,} = searchResult

  // フォルダ名を転記
  sheet.getRange(row, col).setValue(folderName);
  if (files.length === 0 || folders.length === 0 || folders.length > 0) {
    sheet.getRange(row, col).setBackground(COLORS.YELLOW);
  }

  col++;
  folders.forEach(res => {
    // 転記処理を再帰呼び出し
    transcriptionToSheetOfDriveInfo(sheet, res, col, startRow);
  });

  // ファイル一覧が存在しない場合は、カーソル行を進めて処理を抜ける
  if (files.length === 0) {
    col = startCol;
    // 開始の行と現在の行で差分がない場合、行を進める
    if (row - startRow === 0) row++;
    return;
  }
  
  files.forEach(fileName => {
    sheet.getRange(row, col).setValue(fileName);
    row++;
  });
  col = startCol;
}

/**
 * メイン関数
 */
const main = () => {
  const searchResult = recursiveSearchDrive(TARGET_ID);
  const sheet = createSheet(SHEET_NAME);

  transcriptionToSheetOfDriveInfo(sheet, searchResult, 1, 1);

  // To debug
  // console.log(JSON.stringify(searchResult))
}
