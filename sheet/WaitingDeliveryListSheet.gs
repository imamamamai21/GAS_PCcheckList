//////////////////////////////////////////////////////////
////////////////////// 納品待ちPCリスト //////////////////////
//////////////////////////////////////////////////////////

var WaitingDeliveryListSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('納品待ちPC');
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
      product      : filterData.indexOf('製品'),
      model        : filterData.indexOf('モデル'),
      note         : filterData.indexOf('備考'),
      pcUri        : filterData.indexOf('台帳URL'),
      createDate   : filterData.indexOf('登録日'),
      memo         : filterData.indexOf('メモ'),
      checkStatus  : filterData.indexOf('対応済み')
    };
    return this.index;
  }
}
  
WaitingDeliveryListSheet.prototype = {
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
var waitingDeliveryListSheet = new WaitingDeliveryListSheet();

/**
 * 納品待ちのデータをアップロードする
 */
function updateWaitingDeliveryData() {
  var query = 'pc_status in ("納品待ち") and ' + KintoneApi.QUERY_USE_RENTAL + ' and ' + KintoneApi.QUERY_MY_OWNER;
  var fields = [
    KintoneApi.KEY_ID, 'capc_id', 'rentalid', 'user_id', 'user_name', 'user_division', 'location', 'pc_maker', 'pc_product', 'pc_model', 'appendix', 'created_at'
  ];
  var data = KintoneApi.caApi.api.get(query, fields);
  var wSheet = waitingDeliveryListSheet;
  
  // シートに書かれたデータ('対応済'を除く)
  var editData = wSheet.values
    .filter(function(value){ return !value[wSheet.getIndex().checkStatus] && value[wSheet.getIndex().caPcNo]; })
    .map(function(value, index){ return value[wSheet.getIndex().caPcNo]; });
  
  data.forEach(function(value, index) {
    // データと一致したらreturn
    var no = value.capc_id.value;
    var i = editData.indexOf(no);
    if(i > -1) {
      editData.splice(i, 1);
      return;
    }
    var row = wSheet.sheet.getRange('A:A').getValues().filter(String).length + 1;
    
    wSheet.sheet.getRange(wSheet.getRowKey('caPcNo') + row).setValue(no);
    wSheet.sheet.getRange(wSheet.getRowKey('rentalPcNo') + row).setValue(value.rentalid.value);
    wSheet.sheet.getRange(wSheet.getRowKey('employeeNo') + row).setValue(value.user_id.value);
    wSheet.sheet.getRange(wSheet.getRowKey('employeeName') + row).setValue(value.user_name.value);
    wSheet.sheet.getRange(wSheet.getRowKey('employeePlace') + row).setValue(value.user_division.value);
    wSheet.sheet.getRange(wSheet.getRowKey('place') + row).setValue(value.location.value);
    wSheet.sheet.getRange(wSheet.getRowKey('maker') + row).setValue(value.pc_maker.value);
    wSheet.sheet.getRange(wSheet.getRowKey('product') + row).setValue(value.pc_product.value);
    wSheet.sheet.getRange(wSheet.getRowKey('model') + row).setValue(value.pc_model.value);
    wSheet.sheet.getRange(wSheet.getRowKey('note') + row).setValue(value.appendix.value);
    wSheet.sheet.getRange(wSheet.getRowKey('createDate') + row).setValue(value.created_at.value.slice(0, 'yyyy-MM-dd'.length));
    wSheet.sheet.getRange(wSheet.getRowKey('pcUri') + row).setValue(KintoneApi.caApi.api.getUri(value[KintoneApi.KEY_ID].value));
  });
  // すでに書かれていて,且つデータとして来なかったものはステータスを修正する
  editData.forEach(function(no) {
    for(var i = wSheet.titleRow; i < wSheet.values.length; i++) {
      if(wSheet.values[i][wSheet.getIndex().caPcNo] == no) {
        wSheet.sheet.getRange(wSheet.getRowKey('checkStatus') + (i + 1)).setValue(true);
      }
    }
  });
  
  wSheet.sheet.getRange('E1').setValue(Utilities.formatDate(new Date(), "JST", TEXT_DATE));
  wSheet.sheet.getRange('E2').setValue(data.length + '台');
}