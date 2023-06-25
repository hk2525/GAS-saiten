/**
 * 現在開いているシートよりPDFを出力
 */
function createPdf_sendSlack(channel){
  console.log(channel)

  // アクティブなスプレッドシート
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // スプレッドシートID
  const spreadsheetId = spreadsheet.getId();

  // アクティブなシートID
  const gid = spreadsheet.getActiveSheet().getSheetId();

  // スプレッドシートのあるフォルダ
  const folder = DriveApp.getFileById(spreadsheetId).getParents().next();

  // 生徒名を取得
  const name = spreadsheet.getName().replace('理解度テスト採点基準_', '').replace('.xlsm', '');
  console.log(spreadsheet.getName())
  console.log(name)

  // 出力するPDFファイル名
  const fileName = `採点結果_${name}_${spreadsheet.getSheetName()}_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd')}.pdf`;

  // 出力オプション
  var opts = {
    'exportFormat': 'pdf',    // ファイル形式の指定 （pdf / csv / xls / xlsx）
    'format': 'pdf',    // ファイル形式の指定 （pdf / csv / xls / xlsx）
    'size': 'A4',     // 用紙サイズの指定 （legal / letter / A4）
    'portrait': 'true',   // 用紙の向き （true : 縦向き / false : 横向き）
    'fitw': 'true',   // 幅を用紙に合わせるか （true : 合わせる / false : 合わせない）
    'sheetnames': 'false',  // シート名をPDF上部に表示するか （true : 表示する / false : 表示しない）
    'printtitle': 'false',  // スプレッドシート名をPDF上部に表示するか （true : 表示する / false : 表示しない）
    'pagenumbers': 'false',  // ページ番号の有無 （true : 表示する / false : 表示しない）
    'gridlines': 'false',  // グリッドラインの表示有無 （true : 表示する / false : 表示しない）
    'fzr': 'true',   // 固定行の表示有無 （true : 表示する / false : 表示しない）
    'top_margin': 0.8,
    'bottom_margin': 0.8,
    'left_margin': 0.7,
    'right_margin': 0.7,
  };

  let urlExt = [];

  // オプション名と値を「=」で繋げて配列に格納
  for (optName in opts) {
    urlExt.push(optName + '=' + opts[optName]);
  }

  // 各要素を「&」で繋げる
  const options = urlExt.join('&');

  // API使用のためのOAuth認証用トークン
  var token = ScriptApp.getOAuthToken();

  // URLの組み立て
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?gid=${gid}&${options}`;

  // PDF作成
  var response = UrlFetchApp.fetch(
    url, {
    headers: { 'Authorization': 'Bearer ' + token }
  }
  );

  // Blob を作成する
  var blob = response.getBlob().setName(fileName);

  //　PDFを指定したフォルダに保存
  folder.createFile(blob);

  const bot_token = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
  console.log(bot_token)
  const resp = uploadFileToSlack(bot_token, {
    channels: channel.toString(), // 複数渡すならカンマ区切りで
     title: '採点結果です',
     filename: fileName.toString(),
     file: blob,
   });
  console.log(resp)
}

function sendMassage( ) {
  // アクティブなスプレッドシート
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = spreadsheet.getId();
  const gid = spreadsheet.getActiveSheet().getSheetId();

  // 生徒名を取得
  const name = spreadsheet.getName().replace('理解度テスト採点基準_', '').replace('.xlsm', '');
  
  // テスト名を取得
  const testname = spreadsheet.getSheetName();

  //得点を取得
  const scorePos = {
    'Java超入門理解度テスト':'E1',
    '超入門理解度テスト (2回目)':`E1`,
    '超入門理解度テスト (3回目)':`E1`,
    'Java入門理解度テスト_選択式問題':'H2',
    'Java入門理解度テスト_PG作成':`G2`,
    'Java入門理解度テスト2回目_選択式問題 ':`H2`,
    'Java入門理解度テスト2回目_PG作成':'G2',
    'Java基礎理解度テスト_選択式問題':`H2`,
    'Java基礎理解度テスト_記述式問題':'F2',
    'Java基礎理解度テスト_PG作成':`G2`,
    'Java基礎理解度テスト2回目_選択式問題':`H2`,
    'Java基礎理解度テスト2回目_記述式問題':'F2',
    'Java基礎理解度テスト2回目_PG作成':`G2`,
    'JDBC理解度テスト_選択式問題':`H2`,
    'JDBC理解度テスト_PG作成':'G2',
    'JDBC理解度テスト2回目_選択式問題':`E1`,
    'JDBC理解度テスト2回目_PG作成':'G2',
    'JSP理解度テスト_選択式問題':`H2`,
    'JSPサーブレット理解度テスト_PG作成':`G2`,
    'JSP理解度テスト2回目_選択式問題':`H2`,
    'JSPサーブレット理解度テスト2回目_PG作成':`G2`,
    '中間終了理解度テスト_選択式問題':`H2`,
    '中間終了理解度テスト_Java入門':`G2`,
    '中間終了理解度テスト_Java基礎':`F2`,
    '中間終了理解度テスト_JSPサーブレット':`G2`,
  };
  var value = spreadsheet.getRange(scorePos[testname]).getValue();
  console.log(value);
  const message = getMessage(name,testname,value );
  console.log(message);
}

function getMessage(name, testname, score  ) {
  
  const message =`${name}さんお疲れ様です。
採点結果を返却いたします。
■${testname}
得点：${score}点

細かな減点等はありませんでした。`;
}

function uploadFileToSlack(token, payload) {
  const endpoint = "https://www.slack.com/api/files.upload";
  if (payload["file"] !== undefined) {
    payload["token"] = token;
    const response = UrlFetchApp.fetch(endpoint, { method: "post", payload: payload });
    console.log(`Web API (files.upload) response: ${response}`)
    return response;
  } else {
    const response = UrlFetchApp.fetch(endpoint, {
      method: "post",
      contentType: "application/x-www-form-urlencoded",
      headers: { "Authorization": `Bearer ${token}` },
      payload: payload,
    });
    console.log(`Web API (files.upload) response: ${response}`)
    return response;
  }
}