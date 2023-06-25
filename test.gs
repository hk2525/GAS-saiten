function TEST_createPdf() {
  createPdf_sendSlack("")
}


function test_slack_file() {
  const token = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
  const file = DriveApp.getFileById(``);
  const blob = file.getBlob();
  console.log(token)
  const response = uploadFileToSlack(token, {
    channels: '', // 複数渡すならカンマ区切りで
     title: 'file によるアップロード例',
     file: blob,
   });
}

function test_slack() {
  const token = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
  const response = uploadFileToSlack(token, {
    channels: '', // 複数渡すならカンマ区切りで
    title: 'content によるアップロード例',
    content: 'Hello World!',
    filename: 'sample.txt',
    filetype: 'text'
  });
}


// 関数を実行するメニューを追加
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('GAS');
  menu.addItem('PDF生成&Slack投稿', 'createPdf_sendSlack');
  menu.addToUi();
}
