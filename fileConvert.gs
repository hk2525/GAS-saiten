/*
エクセルファイルをフォルダ単位でまとめてスプレッドシートに変換する。
実行にはDrive APIのサービス追加が必要だよ。
*/
const excleFolderId = '';//エクセルフォルダのIDを指定してください
const ssFolderId = '';//スプシフォルダのIDを指定してください

function excel2spreadsheet() {
  // Excelファイルが入っているフォルダをidによって取得
  const sourceFolder = DriveApp.getFolderById(excleFolderId);
  // Excelファイルたちを変数に保存
  const excelFiles = sourceFolder.getFiles();
  // 変換されたファイルが格納されるフォルダをidによって取得
  const destFolder = DriveApp.getFolderById(ssFolderId);
  cleanFolder(destFolder);
  // Excelファイルをイテレートして順にスプレッドシートに変換
  while (excelFiles.hasNext()) {
    var file = excelFiles.next();
    convert2ss(file, destFolder);
  }
}

function spreadsheet2excel() {
  const sourceFolder = DriveApp.getFolderById(ssFolderId);
  const ssFiles = sourceFolder.getFiles();
  const destFolder = DriveApp.getFolderById(excleFolderId);
  cleanFolder(destFolder);
  while (ssFiles.hasNext()) {
    var file = ssFiles.next();
    convert2excle(file, destFolder);
  }
}

function convert2ss(file, folder) {
  // 各種オプションを設定
  // mimeTypeをスプレッドシートにする
  options = {
    title: file.getName(),
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: folder.getId() }]
  };

  // Drive APIへfileをファイルをなげる。
  Drive.Files.insert(options, file.getBlob())
}

function convert2excle(file, folder) {
  //ファイル情報を取得
  var id = file.getId();
  var name = file.getName();

  //ファイルのエクスポートURLを生成
  var url = "https://docs.google.com/spreadsheets/d/" + id + "/export?format=xlsx";

  //urlfetchする際のoptionsを宣言
  var options = {
    method: "get",
    headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
  }

  //urlfetch
  var response = UrlFetchApp.fetch(url, options);
  //urlfetchのレスポンスをblobクラスとして取得
  var blob = response.getBlob();
  //取得したblobクラスから新規ファイルを生成
  var newFile = DriveApp.createFile(blob);
  //作成したファイルの名前を変更
  newFile.setName(name);
  //作成したファイルを格納フォルダに移動
  newFile.moveTo(folder);
}

function cleanFolder(folder) {
  //フォルダ内のすべてのファイルを取得
  var files = folder.getFiles();
  //各ファイルに対して繰り返し
  while (files.hasNext()) {
    //ファイルを取得
    var file = files.next();
    //ゴミ箱へ移動
    file.setTrashed(true);
  }
}