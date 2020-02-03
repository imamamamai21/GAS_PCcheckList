//////////////////////////////////////////////////////////
////////////////////// 廃棄予定PCリスト ////////////////////
//////////////////////////////////////////////////////////

var DiscardListSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('廃棄待ちPC');
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
      statusEditName  : filterData.indexOf('変更者'),
      memo         : filterData.indexOf('メモ'),
      checkStatus  : filterData.indexOf('対応済み'),
      row          : filterData.indexOf('行数')
    };
    return this.index;
  }
}
  
DiscardListSheet.prototype = {
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
var discardListSheet = new DiscardListSheet();

/**
 * 廃棄予定のデータをアップロードする
 */
function updateDiscardData() {
  var index = discardListSheet.getIndex();
  
  // シートに書かれたデータ('対応済'を除く)
  var editData = discardListSheet.values.slice(this.titleRow)
    .filter(function(value) { return value[index.checkStatus] === '' && value[index.caPcNo] != ''; })
    .map(function(value) { return value[index.caPcNo]; });
  
  var query = 'sub_status in ("廃棄待ち") and location not in ("SS", "ABTSD") and ' + KintoneApi.QUERY_MY_OWNER;
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
    var row = discardListSheet.sheet.getRange('A:A').getValues().filter(String).length + 1;
    
    discardListSheet.sheet.getRange(discardListSheet.getRowKey('caPcNo') + row).setValue(no);
    discardListSheet.sheet.getRange(discardListSheet.getRowKey('originPcNo') + row).setValue(value.pc_id.value);
    discardListSheet.sheet.getRange(discardListSheet.getRowKey('rentalPcNo') + row).setValue(value.rentalid.value);
    discardListSheet.sheet.getRange(discardListSheet.getRowKey('status') + row).setValue(value.sub_status.value);
    discardListSheet.sheet.getRange(discardListSheet.getRowKey('locationName') + row).setValue(value.location_name.value);
    
    discardListSheet.sheet.getRange(discardListSheet.getRowKey('maker') + row).setValue(value.pc_maker.value);
    discardListSheet.sheet.getRange(discardListSheet.getRowKey('product') + row).setValue(value.pc_product.value);
    discardListSheet.sheet.getRange(discardListSheet.getRowKey('model') + row).setValue(value.pc_model.value);
    discardListSheet.sheet.getRange(discardListSheet.getRowKey('note') + row).setValue(value.appendix.value);
    discardListSheet.sheet.getRange(discardListSheet.getRowKey('pcUri') + row).setValue(KintoneApi.caApi.api.getUri(value[KintoneApi.KEY_ID].value));
    discardListSheet.sheet.getRange(discardListSheet.getRowKey('row') + row).setValue('=ROW()');
    
    var editRow = discardListSheet.getRowKey('statusEditDate') + row;
    var createRow = discardListSheet.getRowKey('createDate') + row;
    var statusHistryAry = value.status_history.value.split(',');
    discardListSheet.sheet.getRange(editRow).setValue(statusHistryAry.length > 0 ? statusHistryAry[0] : '');
    discardListSheet.sheet.getRange(createRow).setValue(value.created_at.value.slice(0, 'yyyy-MM-dd'.length));
    discardListSheet.sheet.getRange(discardListSheet.getRowKey('statusEditedDate') + row).setValue(getDataIf(editRow)); // 本日からの差分を取る
    discardListSheet.sheet.getRange(discardListSheet.getRowKey('statusEditName') + row).setValue(statusHistryAry.length > 1 ? statusHistryAry[1] : '');
  });
  // すでに書かれていて,且つデータとして来なかったものはステータスを修正する
  editData.forEach(function(no) {
    for(var i = discardListSheet.titleRow; i < discardListSheet.values.length; i++) {
      if(discardListSheet.values[i][index.caPcNo] == no) {
        discardListSheet.sheet.getRange(discardListSheet.getRowKey('checkStatus') + (i + 1)).setValue(true);
      }
    }
  });
  discardListSheet.sheet.getRange('E1').setValue(Utilities.formatDate(new Date(), "JST", TEXT_DATE));
  discardListSheet.sheet.getRange('E2').setValue(data.length + '台');
}