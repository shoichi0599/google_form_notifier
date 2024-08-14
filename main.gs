// For the first execution, 
// you get this exeception "Exception: You do not have permission to call FormApp.getActiveForm."
// To give the permission, just call this method and when a pop-up comes up, accept it.
FormApp.getActiveForm()

function onFormSubmit(e) {
  const itemResponses = e.response.getItemResponses();
  let division;
  let employeeNumber;
  let name;
  let objective1;
  let objective1_1;
  let objective1_2 = "";
  let objective1_3 = "";
  let objective2 = "";
  let objective2_1 = "";
  let objective2_2 = "";
  let objective2_3 = "";
  let objective3 = "";
  let objective3_1 = "";
  let objective3_2 = "";
  let objective3_3 = "";
  // 1) 設問追加時
  // ここで”適当な変数名”を定義
  // フォーマット: let 適当な変数名 = "";

  for (var i=0; i < itemResponses.length; i++) {
    let question = itemResponses[i].getItem().getTitle();
    let answer = itemResponses[i].getResponse();

    switch (question) {
    case "所属先Division":
        division = answer;
        break;
    case "社員番号":
        employeeNumber = answer;
        break;
    case "氏名":
        name = answer;
        break;
    case "Objective①":
        objective1 = answer;
        break;
    case "Objective①_Key Result①":
        objective1_1 = answer;
        break;
    case "Objective①_Key Result②":
        objective1_2 = answer;
        break;
    case "Objective①_Key Result③":
        objective1_3 = answer;
        break;
    case "Objective②":
        objective2 = answer;
        break;
    case "Objective②_Key Result①":
        objective2_1 = answer;
        break;
    case "Objective②_Key Result②":
        objective2_2 = answer;
        break;
    case "Objective②_Key Result③":
        objective2_3 = answer;
        break;
    case "Objective③":
        objective3 = answer;
        break;
    case "Objective③_Key Result①":
        objective3_1 = answer;
        break;
    case "Objective③_Key Result②":
        objective3_2 = answer;
        break;
    case "Objective③_Key Result③":
        objective3_3 = answer;
        break;
    // 2) 設問追加時
    // 下記３行を追加
    // case "設問":
    //     ”適当な変数名” = answer;
    //     break;
    default:
        note = answer;
        break;
    }
  }

  // Find the ID of your spread sheet in a URL like below
  // example: https://docs.google.com/spreadsheets/d/{IP of spread sheet}/
  let spreadSheetId = "{Input your spreadsheet ID}";
  let spreadsheet = SpreadsheetApp.openById(spreadSheetId);
  let activeSheet = spreadsheet.getActiveSheet();
  // This function needs to pass the current last row of the spread sheet
  // so that the next function (executed when cliking the approval link) can input an approval result into this answerer's row

  // 3) 設問追加時
  // ここも追加
  // フォーマット=> 設問：${適当な変数名}
  // この変数"body"はメールの受信者がhtmlを表示できないときに利用
  const body = `
  所属先Division：${division}
  社員番号：${employeeNumber}
  氏名：${name}
  Objective①：${objective1}
  Objective①_Key Result①：${objective1_1}
  Objective①_Key Result②：${objective1_2}
  Objective①_Key Result③：${objective1_3}
  Objective②：${objective2}
  Objective②_Key Result①：${objective2_1}
  Objective②_Key Result②：${objective2_2}
  Objective②_Key Result③：${objective2_3}
  Objective③：${objective3}
  Objective③_Key Result①：${objective3_1}
  Objective③_Key Result②：${objective3_2}
  Objective③_Key Result③：${objective3_3}

  ご確認・承認をお願いいたします。
  `;
  const lastRowNo = activeSheet.getLastRow(); 
  const url = `{Input your application URL}/exec?row=${lastRowNo}`;
  
  // 設問追加するときはここも追加
  // フォーマット=> <span>設問：${適当な変数名}</span><br>
  const htmlBody =`
  <span>所属先Division：${division}</span><br>
  <span>社員番号：${employeeNumber}</span><br>
  <span>氏名：${name}</span><br>
  <span>Objective①：${objective1}</span><br>
  <span>Objective①_Key Result①：${objective1_1}</span><br>
  <span>Objective①_Key Result②：${objective1_2}</span><br>
  <span>Objective①_Key Result③：${objective1_3}</span><br>
  <span>Objective②：${objective2}</span><br>
  <span>Objective②_Key Result①：${objective2_1}</span><br>
  <span>Objective②_Key Result②：${objective2_2}</span><br>
  <span>Objective②_Key Result③：${objective2_3}</span><br>
  <span>Objective③：${objective3}</span><br>
  <span>Objective③_Key Result①：${objective3_1}</span><br>
  <span>Objective③_Key Result②：${objective3_2}</span><br>
  <span>Objective③_Key Result③：${objective3_3}</span><br>
  <br>
  <span>ご確認・承認をお願いいたします。</span><br>
  <br>
  <a href="${url}&status=approve">承認</a> <a href="${url}&status=deny">否認</a>
  `;
  const approverEmail = "{Input an approver's email}";
  GmailApp.sendEmail(approverEmail, "テストです", body, {htmlBody: htmlBody});
}

