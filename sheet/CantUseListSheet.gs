//////////////////////////////////////////////////////////
////////////////////// 使用不可PCリスト ////////////////////
//////////////////////////////////////////////////////////

var CantUseListSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('使用不可PC');
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
      status       : filterData.indexOf('サブステータス'),
      locationName : filterData.indexOf('保管場所'),
      maker        : filterData.indexOf('メーカー'),
      product      : filterData.indexOf('製品'),
      model        : filterData.indexOf('モデル'),
      note         : filterData.indexOf('備考'),
      pcUri        : filterData.indexOf('台帳URL'),
      createDate   : filterData.indexOf('登録日'),
      statusEditDate  : filterData.indexOf('ステータス変更日'),
      statusEditedDate: filterData.indexOf('からの経過日'),
      statusEditName: filterData.indexOf('変更者'),
      memo         : filterData.indexOf('メモ'),
      checkStatus  : filterData.indexOf('対応済み'),
      row          : filterData.indexOf('行数')
    };
    return this.index;
  }
}
  
CantUseListSheet.prototype = {
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
var cantUseListSheet = new CantUseListSheet();

/**
 * 使用不可のデータをアップロードする
 */
function updateCantUseData() {
  var index = cantUseListSheet.getIndex();
  
  // シートに書かれたデータ('対応済'を除く)
  var editData = cantUseListSheet.values.slice(this.titleRow)
    .filter(function(value) { return value[index.checkStatus] === '' && value[index.caPcNo] != ''; })
    .map(function(value) { return value[index.caPcNo]; });
  
  var query = 'pc_status in ("使用不可") and sub_status not in ("譲渡予定", "譲渡可能", "廃棄待ち") and ' + KintoneApi.QUERY_MY_OWNER;
  var fields = [
    KintoneApi.KEY_ID, 'pc_id', 'sub_status', 'capc_id', 'rentalid', 'location_name', 'pc_maker', 'pc_product', 'pc_model', 'appendix', 'status_history', 'created_at'
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
    var row = cantUseListSheet.sheet.getRange('A:A').getValues().filter(String).length + 1;
    
    cantUseListSheet.sheet.getRange(cantUseListSheet.getRowKey('caPcNo') + row).setValue(no);
    cantUseListSheet.sheet.getRange(cantUseListSheet.getRowKey('originPcNo') + row).setValue(value.pc_id.value);
    cantUseListSheet.sheet.getRange(cantUseListSheet.getRowKey('rentalPcNo') + row).setValue(value.rentalid.value);
    cantUseListSheet.sheet.getRange(cantUseListSheet.getRowKey('status') + row).setValue(value.sub_status.value);
    cantUseListSheet.sheet.getRange(cantUseListSheet.getRowKey('locationName') + row).setValue(value.location_name.value);
    
    cantUseListSheet.sheet.getRange(cantUseListSheet.getRowKey('maker') + row).setValue(value.pc_maker.value);
    cantUseListSheet.sheet.getRange(cantUseListSheet.getRowKey('product') + row).setValue(value.pc_product.value);
    cantUseListSheet.sheet.getRange(cantUseListSheet.getRowKey('model') + row).setValue(value.pc_model.value);
    cantUseListSheet.sheet.getRange(cantUseListSheet.getRowKey('note') + row).setValue(value.appendix.value);
    cantUseListSheet.sheet.getRange(cantUseListSheet.getRowKey('pcUri') + row).setValue(KintoneApi.caApi.api.getUri(value[KintoneApi.KEY_ID].value));
    cantUseListSheet.sheet.getRange(cantUseListSheet.getRowKey('row') + row).setValue('=ROW()');
    
    var editRow = cantUseListSheet.getRowKey('statusEditDate') + row;
    var createRow = cantUseListSheet.getRowKey('createDate') + row;
    var statusHistryAry = value.status_history.value.split(',');
    cantUseListSheet.sheet.getRange(editRow).setValue(statusHistryAry.length > 0 ? statusHistryAry[0] : '');
    cantUseListSheet.sheet.getRange(createRow).setValue(value.created_at.value.slice(0, 'yyyy-MM-dd'.length));
    cantUseListSheet.sheet.getRange(cantUseListSheet.getRowKey('statusEditedDate') + row).setValue(getDataIf(editRow)); // 本日からの差分を取る
    cantUseListSheet.sheet.getRange(cantUseListSheet.getRowKey('statusEditName') + row).setValue(statusHistryAry.length > 1 ? statusHistryAry[1] : '');
  });
  // すでに書かれていて,且つデータとして来なかったものはステータスを修正する
  editData.forEach(function(no) {
    for(var i = cantUseListSheet.titleRow; i < cantUseListSheet.values.length; i++) {
      if(cantUseListSheet.values[i][index.caPcNo] == no) {
        cantUseListSheet.sheet.getRange(cantUseListSheet.getRowKey('checkStatus') + (i + 1)).setValue(true);
      }
    }
  });
  cantUseListSheet.sheet.getRange('E1').setValue(Utilities.formatDate(new Date(), "JST", TEXT_DATE));
  cantUseListSheet.sheet.getRange('E2').setValue(data.length + '台');
}