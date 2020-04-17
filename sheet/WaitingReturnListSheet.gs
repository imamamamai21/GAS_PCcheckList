//////////////////////////////////////////////////////////
////////////////////// 返却待ちPCリスト ////////////////////
//////////////////////////////////////////////////////////

var WaitingReturnListSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('返却待ちPC');
  this.values = this.sheet.getDataRange().getValues();
  this.titleRow = 0;
  this.index = {};
  
  this.createIndex = function() {
    const PCNO = 'CAグループPC番号';
    var me = this;
    var filterData = (function() {
      for(var i = 0; i < me.values.length; i++) {
        if (me.values[i].indexOf(PCNO) > -1) {
          me.titleRow = i + 1;
          return me.values[i];
        }
      }
    }());
    if(!filterData || filterData.length === 0) {
      showTitleError();
      return;
    }
    
    this.index = {
      caPcNo       : filterData.indexOf(PCNO),
      pcNo         : filterData.indexOf('自社PC管理番号'),
      rentalPcNo   : filterData.indexOf('レンタル管理番号'),
      status       : filterData.indexOf('ステータス'),
      employeeNo   : filterData.indexOf('管理者 社員番号'),
      employeeName : filterData.indexOf('管理者名'),
      employeePlace: filterData.indexOf('管理者所属'),
      employeeMail : filterData.indexOf('管理者メアド'),
      maker        : filterData.indexOf('メーカー'),
      product      : filterData.indexOf('製品'),
      model        : filterData.indexOf('モデル'),
      note         : filterData.indexOf('備考'),
      pcUri        : filterData.indexOf('台帳URL'),
      createDate   : filterData.indexOf('登録日'),
      statusEditDate  : filterData.indexOf('ステータス変更日'),
      statusEditedDate: filterData.indexOf('からの経過日'),
      statusEditName  : filterData.indexOf('変更者'),
      superiorMail : filterData.indexOf('上長のメアド'),
      superiorName : filterData.indexOf('上長名'),
      remind       : filterData.indexOf('リマインド'),
      remindDate   : filterData.indexOf('最終送信日'),
      memo         : filterData.indexOf('メモ'),
      row          : filterData.indexOf('行数'),
      checkStatus  : filterData.indexOf('対応済み')
    };
    return this.index;
  }
}
  
WaitingReturnListSheet.prototype = {
  getRowKey: function(target) {
    var targetIndex = this.getIndex()[target];
    var alfabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
    var returnKey = (targetIndex > -1) ? alfabet[targetIndex] : '';
    if (!returnKey || returnKey === '') showTitleError(target);
    return returnKey;
  },
  getIndex: function() {
    return Object.keys(this.index).length ? this.index : this.createIndex();
  }
};
var waitingReturnListSheet = new WaitingReturnListSheet();

/**
 * 返却待ちのデータをアップロードする
 */
function updateWaitingReturnData() {
  var query = 'sub_status in ("返却待ち") and ' + KintoneApi.QUERY_MY_OWNER;
  var fields = [
    KintoneApi.KEY_ID, 'capc_id', 'pc_id', 'rentalid', 'pc_status', 'status_history', 'user_id', 'user_name', 'user_division', 'pc_maker', 'pc_product', 'pc_model', 'appendix', 'created_at'
  ];
  var data = KintoneApi.caApi.api.get(query, fields);
  var wSheet = waitingReturnListSheet;
  
  // シートに書かれたデータ('対応済'を除く)
  var editData = wSheet.values
    .filter(function(value){ return !value[wSheet.getIndex().checkStatus] && value[wSheet.getIndex().caPcNo]; })
    .map(function(value){ return value[wSheet.getIndex().caPcNo]; });
  
  data.forEach(function(value, index) {
    // データと一致したらreturn
    var no = value.capc_id.value;
    var indexOf = editData.indexOf(no);
    if(indexOf > -1) {
      editData.splice(indexOf, 1);
      return;
    }
    var row = wSheet.sheet.getRange('A:A').getValues().filter(String).length + 1;
    
    wSheet.sheet.getRange(wSheet.getRowKey('caPcNo') + row).setValue(no);
    wSheet.sheet.getRange(wSheet.getRowKey('pcNo') + row).setValue(value.pc_id.value);
    wSheet.sheet.getRange(wSheet.getRowKey('rentalPcNo') + row).setValue(value.rentalid.value);
    wSheet.sheet.getRange(wSheet.getRowKey('status') + row).setValue(value.pc_status.value);
    wSheet.sheet.getRange(wSheet.getRowKey('employeeNo') + row).setValue(value.user_id.value);
    wSheet.sheet.getRange(wSheet.getRowKey('employeeName') + row).setValue(value.user_name.value);
    wSheet.sheet.getRange(wSheet.getRowKey('employeePlace') + row).setValue(value.user_division.value);
    wSheet.sheet.getRange(wSheet.getRowKey('maker') + row).setValue(value.pc_maker.value);
    wSheet.sheet.getRange(wSheet.getRowKey('product') + row).setValue(value.pc_product.value);
    wSheet.sheet.getRange(wSheet.getRowKey('model') + row).setValue(value.pc_model.value);
    wSheet.sheet.getRange(wSheet.getRowKey('note') + row).setValue(value.appendix.value);
    wSheet.sheet.getRange(wSheet.getRowKey('createDate') + row).setValue(value.created_at.value.slice(0, 'yyyy-MM-dd'.length));
    wSheet.sheet.getRange(wSheet.getRowKey('pcUri') + row).setValue(KintoneApi.caApi.api.getUri(value[KintoneApi.KEY_ID].value));
    wSheet.sheet.getRange(wSheet.getRowKey('row') + row).setValue('=ROW()');
    wSheet.sheet.getRange(wSheet.getRowKey('remind') + row).setValue(true);
    
    // 変更者など
    var editRow = wSheet.getRowKey('statusEditDate') + row;
    var createRow = wSheet.getRowKey('createDate') + row;
    var statusHistryAry = value.status_history.value.split(',');
    wSheet.sheet.getRange(editRow).setValue(statusHistryAry.length > 0 ? statusHistryAry[0] : '');
    wSheet.sheet.getRange(createRow).setValue(value.created_at.value.slice(0, 'yyyy-MM-dd'.length));
    wSheet.sheet.getRange(wSheet.getRowKey('statusEditedDate') + row).setValue(getDataIf(editRow)); // 本日からの差分を取る
    wSheet.sheet.getRange(wSheet.getRowKey('statusEditName') + row).setValue(statusHistryAry.length > 1 ? statusHistryAry[1] : '');
    
    // メアド取得
    var mail = KintoneApi.manMasterApi.getMail(value.user_id.value)
    wSheet.sheet.getRange(wSheet.getRowKey('employeeMail') + row).setValue(mail);
  });
  // すでに書かれていて,且つデータとして来なかったものはステータスを修正する
  editData.forEach(function(no) {
    for(var i = wSheet.titleRow; i < wSheet.values.length; i++) {
      if(wSheet.values[i][wSheet.getIndex().caPcNo] == no) {
        wSheet.sheet.getRange(wSheet.getRowKey('checkStatus') + (i + 1)).setValue(true);
        wSheet.sheet.getRange(wSheet.getRowKey('statusEditedDate') + (i + 1)).setValue(wSheet.values[i][wSheet.getIndex().statusEditedDate]);
      }
    }
  });
  
  wSheet.sheet.getRange('E1').setValue(Utilities.formatDate(new Date(), "JST", TEXT_DATE));
  wSheet.sheet.getRange('E2').setValue(data.length + '台');
}

/**
 * 返却のリマンドメールを送る
 */
function sendRemindMail_ReturnPc() {
  var index = waitingReturnListSheet.getIndex();
  var today = Utilities.formatDate(new Date(), 'JST', 'MM/dd(E)');
  const REMIND_DAY = 3;
  
  // シートに書かれたデータ('リマンドにチェック''対応済'を除く)
  var remindValues = waitingReturnListSheet.values
    .filter(function(value){
      return !value[index.checkStatus] && value[index.remind]
        && value[index.statusEditedDate] >= REMIND_DAY && value[index.statusEditedDate] % REMIND_DAY === 0
        && value[index.remindDate] != today;
     });
     
  var mailTemplate = mailTemplateSheet.getTemplate('returnRemind');
  remindValues.forEach(function(value) {
    var header = value[index.employeeName] + 'さん\n' + (value[index.superiorName] === '' ? '': 'CC:' + value[index.superiorName] + 'さん\n') + '\n';
    var text = header + mailTemplate.text
        .replace('{userName}', value[index.employeeName])
        .replace('{maker}', value[index.maker])
        .replace('{product}', value[index.product])
        .replace('{model}', value[index.model])
        .replace('{pcNo}', value[index.caPcNo] + (value[index.pcNo] === '' ? '' : ' / 管理番号2：' + value[index.pcNo]) +  (value[index.rentalPcNo] === '' ? '' : ' / レンタルPC番号：' + value[index.rentalPcNo]));
    mailTemplateSheet.sendMail(
      value[index.employeeMail],
      value[index.superiorMail],
      (value[index.remindDate] === '' ? '' : '(再送)') + mailTemplate.title,
      text,
      false
    );
    waitingReturnListSheet.sheet.getRange(waitingReturnListSheet.getRowKey('remindDate') + value[index.row]).setValue(today);
  })
}