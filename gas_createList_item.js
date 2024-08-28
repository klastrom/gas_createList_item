
function generateHTML() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // 対象列と出力先列のマッピング
  var inputColumns = ['C', 'D', 'E', 'F'];
  var outputColumns = ['H', 'I', 'J', 'K'];

  // 開始行を指定する配列を作成（規則に基づいて動的に生成）
  var startRows = [3, 10, 17];
  var rowInterval = 6; // 開始行から終了行までの行数 (例: 3〜8行 -> 6行)
  
  // 各範囲ごとに処理を実施
  for (var r = 0; r < startRows.length; r++) {
    var startRow = startRows[r];
    var endRow = startRow + rowInterval - 1;
    
    for (var j = 0; j < inputColumns.length; j++) {
      // 各範囲の値を取得
      var values = sheet.getRange(inputColumns[j] + startRow + ':' + inputColumns[j] + endRow).getValues();

      // HTMLテンプレートを作成
      var htmlOutput = 
        '<li class="fripdesk-personal-recommend__item">\n' +
        '  <a class="fripdesk-personal-recommend__item-link" href="' + values[0][0] + '">\n' +
        '    <div class="fripdesk-personal-recommend__item-image">\n' +
        '      <img src="' + values[5][0] + '" alt="' + values[1][0] +  '">\n' +
        '    </div>\n' +
        '    <div class="fripdesk-personal-recommend__item-title">' + values[1][0] + '</div>\n' +
        '    <p class="fripdesk-personal-recommend__item-detail">' + values[2][0] + '</p>\n' +
        '    <div class="fripdesk-personal-recommend__item-price">\n' +
        '      1' + values[3][0] + '<span class="fripdesk-personal-recommend__item-price-highlight">' + values[4][0] + '</span>円(税抜)～\n' +
        '    </div>\n' +
        '  </a>\n' +
        '</li>';
      
      // 各範囲の出力位置にHTMLを出力
      sheet.getRange(outputColumns[j] + startRow).setValue(htmlOutput);
    }
  }
}


