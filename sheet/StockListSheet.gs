//////////////////////////////////////////////////////////
////////////////////// 動作確認中PCリスト ////////////////////
//////////////////////////////////////////////////////////

var StockListSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('動作確認中PC');
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
      originPcNo   : filterData.indexOf('自社PC管理番号'),
      rentalPcNo   : filterData.indexOf('レンタル管理番号'),
      status       : filterData.indexOf('ステータス'),
      locationName : filterData.indexOf('保管場所'),
      maker        : filterData.indexOf('メーカー'),
      product      : filterData.indexOf('製品'),
      model        : filterData.indexOf('モデル'),
      note         : filterData.indexOf('備考'),
      pcUri        : filterData.indexOf('台帳URL'),
      createDate   : filterData.indexOf('登録日'),
      statusEditDate  : filterData.indexOf('ステータス変更日'),
      statusEditedDate: filterData.indexOf('からの経過日'),
      statusEditName  : filterData.indexOf('変更者'),
      memo         : filterData.indexOf('メモ'),
      checkStatus  : filterData.indexOf('対応済み'),
      row          : filterData.indexOf('行数')
    };
    return this.index;
  }
}
  
StockListSheet.prototype = {
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
var stockListSheet = new StockListSheet();

/**
 * 廃棄予定のデータをアップロードする
 */
function updateStockListData() {
  var index = stockListSheet.getIndex();
  
  // シートに書かれたデータ('対応済'を除く)
  var editData = stockListSheet.values.slice(this.titleRow)
    .filter(function(value) { return value[index.checkStatus] === '' && value[index.caPcNo] != ''; })
    .map(function(value) { return value[index.caPcNo]; });
  
  var query = 'sub_status in ("動作確認中") and ' + KintoneApi.QUERY_MY_OWNER;
  var fields = [
    KintoneApi.KEY_ID, 'pc_id', 'pc_status', 'capc_id', 'rentalid', 'location_name', 'pc_maker', 'pc_product', 'pc_model', 'appendix', 'status_history', 'created_at'
  ];
  var data = KintoneApi.caApi.api.get(query, fields);
  data.forEach(function(value, index) {
    // データと一致したらreturn
    var no = value.capc_id.value;
    var i = editData.indexOf(no);
    var e = editData
    if(i > -1) {
      editData.splice(i, 1);
      return;
    }
    var row = stockListSheet.sheet.getRange('A:A').getValues().filter(String).length + 1;
    
    stockListSheet.sheet.getRange(stockListSheet.getRowKey('caPcNo') + row).setValue(no);
    stockListSheet.sheet.getRange(stockListSheet.getRowKey('originPcNo') + row).setValue(value.pc_id.value);
    stockListSheet.sheet.getRange(stockListSheet.getRowKey('rentalPcNo') + row).setValue(value.rentalid.value);
    stockListSheet.sheet.getRange(stockListSheet.getRowKey('status') + row).setValue(value.pc_status.value);
    stockListSheet.sheet.getRange(stockListSheet.getRowKey('locationName') + row).setValue(value.location_name.value);
    
    stockListSheet.sheet.getRange(stockListSheet.getRowKey('maker') + row).setValue(value.pc_maker.value);
    stockListSheet.sheet.getRange(stockListSheet.getRowKey('product') + row).setValue(value.pc_product.value);
    stockListSheet.sheet.getRange(stockListSheet.getRowKey('model') + row).setValue(value.pc_model.value);
    stockListSheet.sheet.getRange(stockListSheet.getRowKey('note') + row).setValue(value.appendix.value);
    stockListSheet.sheet.getRange(stockListSheet.getRowKey('pcUri') + row).setValue(KintoneApi.caApi.api.getUri(value[KintoneApi.KEY_ID].value));
    stockListSheet.sheet.getRange(stockListSheet.getRowKey('row') + row).setValue('=ROW()');
    
    var editRow = stockListSheet.getRowKey('statusEditDate') + row;
    var createRow = stockListSheet.getRowKey('createDate') + row;
    var statusHistryAry = value.status_history.value.split(',');
    stockListSheet.sheet.getRange(editRow).setValue(statusHistryAry.length > 0 ? statusHistryAry[0] : '');
    stockListSheet.sheet.getRange(createRow).setValue(value.created_at.value.slice(0, 'yyyy-MM-dd'.length));
    stockListSheet.sheet.getRange(stockListSheet.getRowKey('statusEditedDate') + row).setValue(getDataIf(editRow)); // 本日からの差分を取る
    stockListSheet.sheet.getRange(stockListSheet.getRowKey('statusEditName') + row).setValue(statusHistryAry.length > 1 ? statusHistryAry[1] : '');
  });
  // すでに書かれていて,且つデータとして来なかったものはステータスを修正する
  editData.forEach(function(no) {
    for(var i = stockListSheet.titleRow; i < stockListSheet.values.length; i++) {
      if(stockListSheet.values[i][index.caPcNo] == no) {
        stockListSheet.sheet.getRange(stockListSheet.getRowKey('checkStatus') + (i + 1)).setValue(true);
      }
    }
  });
  stockListSheet.sheet.getRange('E1').setValue(Utilities.formatDate(new Date(), "JST", TEXT_DATE));
  stockListSheet.sheet.getRange('E2').setValue(data.length + '台');
}