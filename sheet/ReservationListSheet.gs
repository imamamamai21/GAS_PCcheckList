//////////////////////////////////////////////////////////
////////////////////// 配布予定PCリスト //////////////////////
//////////////////////////////////////////////////////////

var ReservationListSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('配布予定PC');
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
      rentalPcNo   : filterData.indexOf('レンタルPC管理番号'),
      employeeNo   : filterData.indexOf('管理者 社員番号'),
      employeeName : filterData.indexOf('管理者名'),
      employeePlace: filterData.indexOf('管理者所属'),
      place        : filterData.indexOf('保管場所'),
      maker        : filterData.indexOf('メーカー'),
      product      : filterData.indexOf('製品名'),
      model        : filterData.indexOf('モデル'),
      note         : filterData.indexOf('備考'),
      pcUri        : filterData.indexOf('台帳URL'),
      createDate   : filterData.indexOf('登録日'),
      statusEditDate  : filterData.indexOf('ステータス変更日'),
      statusEditedDate: filterData.indexOf('からの経過日'),
      statusEditName  : filterData.indexOf('変更者'),
      memo         : filterData.indexOf('メモ'),
      checkStatus  : filterData.indexOf('対応済み')
    };
    return this.index;
  }
}
  
ReservationListSheet.prototype = {
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
var reservationListSheet = new ReservationListSheet();

/**
 * 配布予定のデータをアップロードする
 */
function updateReservationData() {
  var query = 'pc_status in ("配布予定") and ' + KintoneApi.QUERY_USE_RENTAL + ' and ' + KintoneApi.QUERY_MY_OWNER;
  var fields = [
    KintoneApi.KEY_ID, 'capc_id', 'rentalid', 'user_id', 'user_name', 'user_division', 'location', 'pc_maker', 'pc_product', 'pc_model', 'appendix', 'status_history', 'created_at'
  ]
  var data = KintoneApi.caApi.api.get(query, fields);
  var r = reservationListSheet;
  
  // シートに書かれたデータ('対応済'を除く)
  var editData = r.values
    .filter(function(value){ return !value[r.getIndex().checkStatus] && value[r.getIndex().caPcNo]; })
    .map(function(value, index){ return value[r.getIndex().caPcNo]; });
  
  data.forEach(function(value, index) {
    // データと一致したらreturn
    var no = value.capc_id.value;
    var i = editData.indexOf(no);
    if(i > -1) {
      editData.splice(i, 1);
      return;
    }
    var row = r.sheet.getRange('A:A').getValues().filter(String).length + 1;
    r.sheet.insertRowAfter(row); // 一番下に追記
    
    r.sheet.getRange(r.getRowKey('caPcNo') + row).setValue(no);
    r.sheet.getRange(r.getRowKey('rentalPcNo') + row).setValue(value.rentalid.value);
    r.sheet.getRange(r.getRowKey('employeeNo') + row).setValue(value.user_id.value);
    r.sheet.getRange(r.getRowKey('employeeName') + row).setValue(value.user_name.value);
    r.sheet.getRange(r.getRowKey('employeePlace') + row).setValue(value.user_division.value);
    r.sheet.getRange(r.getRowKey('place') + row).setValue(value.location.value);
    r.sheet.getRange(r.getRowKey('maker') + row).setValue(value.pc_maker.value);
    r.sheet.getRange(r.getRowKey('product') + row).setValue(value.pc_product.value);
    r.sheet.getRange(r.getRowKey('model') + row).setValue(value.pc_model.value);
    r.sheet.getRange(r.getRowKey('note') + row).setValue(value.appendix.value);
    r.sheet.getRange(r.getRowKey('pcUri') + row).setValue(KintoneApi.caApi.api.getUri(value[KintoneApi.KEY_ID].value));
    
    var editRow = r.getRowKey('statusEditDate') + row;
    var createRow = r.getRowKey('createDate') + row;
    var statushistryAry = value.status_history.value.split(',');
    var ifText = function(row) { return 'DATEDIF(' + row + ',today(), "d")'; };
    r.sheet.getRange(editRow).setValue(statushistryAry.length > 0 ? statushistryAry[0] : '');
    r.sheet.getRange(createRow).setValue(value.created_at.value.slice(0, 'yyyy-MM-dd'.length));
    r.sheet.getRange(r.getRowKey('statusEditedDate') + row).setValue('=if(' + editRow + '<>"", ' + ifText(editRow) + ',' + ifText(createRow) + ')'); // 本日からの差分を取る
    r.sheet.getRange(r.getRowKey('statusEditName') + row).setValue(statushistryAry.length > 1 ? statushistryAry[1] : '');
  });
  // すでに書かれていて,且つデータとして来なかったものはステータスを修正する 
  editData.forEach(function(no) {
    for(var i = r.titleRow; i < r.values.length; i++) {
      if(r.values[i][r.getIndex().caPcNo] === no) {
        r.sheet.getRange(r.getRowKey('checkStatus') + (i + 1)).setValue(true);
      }
    }
  });
  
  r.sheet.getRange('E1').setValue(Utilities.formatDate(new Date(), "JST", TEXT_DATE));
  r.sheet.getRange('E2').setValue(data.length + '台');
}