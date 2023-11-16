// Spreadsheetが開かれた時に自動的に実行
function onOpen() {
// 現在開いている、スプレッドシートを取得
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
// メニュー項目を定義
const entries = [{name : "請求書作成",functionName : "create"}];
// 「書類作成」という名前でメニューに追加
spreadsheet.addMenu("書類作成", entries);
}

function createOrGetFolder(folderName) {
    // Googleドライブ内のフォルダを検索
    const folders = DriveApp.getFoldersByName(folderName);
    let folder;

    if (folders.hasNext()) {
        // 指定された名前のフォルダが存在する場合は、そのフォルダを使用
        folder = folders.next();
    } else {
        // フォルダが存在しない場合は、新しいフォルダを作成
        folder = DriveApp.createFolder(folderName);
    }
    return folder;
}

function markAsCompleted(rowNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('一覧');
  sheet.getRange('AA' + rowNumber).setValue('作成済')
}

function create() {

  //スプレッドシートを設定
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();//変数spreadsheetに「アクティブなスプレッドシート」を設定
  const sheet = spreadsheet.getSheetByName('一覧');//変数sheetに「一覧」シートを設定
  const myRange = sheet.getDataRange().getValues();//スプレッドシートのデータを二次元配列として取得
  const template = spreadsheet.getSheetByName('請求書');//変数templateに「請求書」シートを設定
  
  //空の配列を設定
  let client_list = [];
  //顧客先（C列）を空配列に取得
  for (let i=1; i<myRange.length; i++){ 
    let status = myRange[i][26];
    if(status === '' || status !== '作成済'){
      client_list.push(myRange[i][3]);//配列clientにmyRange[i][３]を追加
    }
  }
  
  //商品名ごとに繰り返す
  for (let i=0; i<client_list.length; i++){ 
    //プログラムA-6-1｜空の配列を設定
    let myDate = [];//発行日
    let myPayment = [];//支払期日
    let myID = []; //ID
    let myItem1 = [];//品目１
    let myItem1Price = [];//品目１金額
    let myItem1Qty = []; //品目１個数
    let myItem2 = [];//品目2
    let myItem2Price = [];//品目2金額
    let myItem2Qty = []; //品目2個数
    let myItem3 = [];//品目3
    let myItem3Price = [];//品目3金額
    let myItem3Qty = []; //品目3個数
    let myItem4 = [];//品目4
    let myItem4Price = [];//品目4金額
    let myItem4Qty = []; //品目4個数
    let myItem5 = [];//品目5
    let myItem5Price = [];//品目5金額
    let myItem5Qty = []; //品目5個数
    let myItem6 = [];//品目6
    let myItem6Price = [];//品目6金額
    let myItem6Qty = []; //品目6個数
    let myMemo1 = []; //備考１
    let myMemo2 = []; //備考2
    let myMemo3 = []; //備考3
    let myMemo4 = []; //備考4
    //配列に格納
    for (let k=0; k<myRange.length; k++){ 
        if (myRange[k][3] == client_list[i]){ //myRange[k][3](顧客先)とclient_list[i]が一致すれば
          myDate.push(myRange[k][0]);
          myPayment.push(myRange[k][1]);
          myID.push(myRange[k][2]);
          myItem1.push(myRange[k][4]);
          myItem1Price.push(myRange[k][5]);
          myItem1Qty.push(myRange[k][6]);
          myItem2.push(myRange[k][7]);
          myItem2Price.push(myRange[k][8]);
          myItem2Qty.push(myRange[k][9]);
          myItem3.push(myRange[k][10]);
          myItem3Price.push(myRange[k][11]);
          myItem3Qty.push(myRange[k][12]);
          myItem4.push(myRange[k][13]);
          myItem4Price.push(myRange[k][14]);
          myItem4Qty.push(myRange[k][15]);
          myItem5.push(myRange[k][16]);
          myItem5Price.push(myRange[k][17]);
          myItem5Qty.push(myRange[k][18]);
          myItem6.push(myRange[k][19]);
          myItem6Price.push(myRange[k][20]);
          myItem6Qty.push(myRange[k][21]);
          myMemo1.push(myRange[k][22]);
          myMemo2.push(myRange[k][23]);
          myMemo3.push(myRange[k][24]);
          myMemo4.push(myRange[k][25]);
        }
    }

    //シートを追加して、シート名を各顧客先に変更
    const newsheet = template.copyTo(spreadsheet);//「請求書フォーマット」のシートをコピーする
    newsheet.setName(client_list[i]);//コピーしたシートの名前を「client[i]」にする
    console.log(client_list[i]);
    //貼付
    myDate.length > 0 ? newsheet.getRange('H1').setValue(myDate[0]): null;//発行日
    newsheet.getRange('B4').setValue(client_list[i]);//顧客
    myPayment.length > 0 ? newsheet.getRange('C18').setValue(myPayment[0]): null;//支払期日
    myItem1.length > 0 ? newsheet.getRange('A21').setValue(myItem1[0]): null;//品目1
    myItem1Price.length > 0 ? newsheet.getRange('G21').setValue(myItem1Price[0]): null;//品目1単価
    myItem1Qty.length > 0 ? newsheet.getRange('H21').setValue(myItem1Qty[0]): null;//品目1個数
    myItem2.length > 0 ? newsheet.getRange('A22').setValue(myItem2[0]): null;//品目２
    myItem2Price.length > 0 ? newsheet.getRange('G22').setValue(myItem2Price[0]): null;//品目2単価
    myItem2Qty.length > 0 ? newsheet.getRange('H22').setValue(myItem2Qty[0]): null;//品目2個数
    myItem3.length > 0 ? newsheet.getRange('A23').setValue(myItem3[0]): null;//品目３
    myItem3Price.length > 0 ? newsheet.getRange('G23').setValue(myItem3Price[0]): null;//品目3単価
    myItem3Qty.length > 0 ? newsheet.getRange('H23').setValue(myItem3Qty[0]): null;//品目3個数
    myItem4.length > 0 ? newsheet.getRange('A25').setValue(myItem4[0]): null;//品目4
    myItem4Price.length > 0 ? newsheet.getRange('G25').setValue(myItem4Price[0]): null;//品目4単価
    myItem4Qty.length > 0 ? newsheet.getRange('H25').setValue(myItem4Qty[0]): null;//品目4個数
    myItem5.length > 0 ? newsheet.getRange('A26').setValue(myItem5[0]): null;//品目5
    myItem5Price.length > 0 ? newsheet.getRange('G26').setValue(myItem5Price[0]): null;//品目5単価
    myItem5Qty.length > 0 ? newsheet.getRange('H26').setValue(myItem5Qty[0]): null;//品目5個数
    myItem6.length > 0 ? newsheet.getRange('A27').setValue(myItem6[0]): null;//品目6
    myItem6Price.length > 0 ? newsheet.getRange('G27').setValue(myItem6Price[0]): null;//品目6単価
    myItem6Qty.length > 0 ? newsheet.getRange('H27').setValue(myItem6Qty[0]): null;//品目6個数
    myMemo1.length > 0 ? newsheet.getRange('A33').setValue(myMemo1[0]): null;//備考欄１
    myMemo2.length > 0 ? newsheet.getRange('A34').setValue(myMemo2[0]): null;//備考欄2
    myMemo3.length > 0 ? newsheet.getRange('A35').setValue(myMemo3[0]): null;//備考欄3
    myMemo4.length > 0 ? newsheet.getRange('A36').setValue(myMemo4[0]): null;//備考欄4
    let rowNumber = myID[0] + 1;
    markAsCompleted(rowNumber);
    Utilities.sleep(1000); //1秒待機（待機中に情報を更新）
    SpreadsheetApp.flush(); //挿入したシートの情報更新

    //プログラムA-6-8｜PDF化
    const ssId = spreadsheet.getId();//スプレッドシートIDを取得
    const sheetId = newsheet.getSheetId();//請求書のシートIDを取得
    const folderName = newsheet.getRange('L6').getValue();//newsheetのセルJ2の値（
    let folder = createOrGetFolder(folderName);
    PDFexport(ssId, sheetId, client_list[i], folder);//プログラムBを実行（5つの引数を渡す）
    console.log(client_list[i]);
  }
}

    //プログラムB-0｜PDF化
function PDFexport(ssId, sheetId, client, folder) {
  
  //プログラムB-1｜PDF化の条件設定
  var url = 'https://docs.google.com/spreadsheets/d/'+ ssId +'/export?';
  var opts = {
    exportFormat: 'pdf',      // ファイル形式の指定
    format:       'pdf',      // ファイル形式の指定
    size:         'A4',       // 用紙サイズの指定
    portrait:     'true',     // true縦向き、false 横向き
    fitw:         'true',     // 幅を用紙に合わせるか？
    sheetnames:   'false',    // シート名を PDF 上部に表示するか？
    printtitle:   'false',    // スプレッドシート名をPDF上部に表示するか？
    pagenumbers:  'false',    // ページ番号の有無
    gridlines:    'false',    // グリッドラインの表示有無
    fzr:          'false',    // 固定行の表示有無
    range :       'A1%3AI41',  // 対象範囲「%3A」 = : (コロン)  
    gid:           sheetId    // シート ID を指定 (省略する場合、すべてのシートをダウンロード)
  };
  
  //プログラムB-2｜PDF化のurl作成
  var PDFurl = [];//urlという空配列を設定
  for(optName in opts){
    PDFurl.push(optName + '=' + opts[optName]);//opts配列の各要素を=でつないだものをurl配列に格納
  }
  var options  = PDFurl.join('&');//urlの配列の各要素を&でつなぐ
  
  //プログラムB-3｜PDF化の条件設定
  var token    = ScriptApp.getOAuthToken();//アクセストークンを取得
  var response = UrlFetchApp.fetch(url + options, {headers: {'Authorization': 'Bearer ' +  token}}); //PDFのURLからアクセスする
  var blob = response.getBlob().setName(client + '.pdf');//PDFの名前を「取引先+.pdf」とする
  var newFile = folder.createFile(blob);//PDFを所定のフォルダに保管する
  newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);//共有設定をする：「リンクを知っている人」が「閲覧可能」
}
