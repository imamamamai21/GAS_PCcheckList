//////////////////////////////////////////////////////////
//////////////////// レンタルPC・在庫リスト ////////////////////
//////////////////////////////////////////////////////////


var RentalStokeListSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('レンタルPC/在庫');
  this.values = this.sheet.getDataRange().getValues();
  this.titleRow = 0;
  this.index = {};
  
  this.createIndex = function() {
    const STATUS = 'ステータス';
    var me = this;
    var filterData = (function() {
      for(var i = 0; i < me.values.length; i++) {
        if (me.values[i].indexOf(STATUS) > -1) {
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
      status       : filterData.indexOf(STATUS),
      subStatus    : filterData.indexOf('サブステータス'),
      caPcNo       : filterData.indexOf('CAグループPC番号'),
      rentalPcNo   : filterData.indexOf('レンタルPC管理番号'),
      place        : filterData.indexOf('保管場所'),
      employeeName : filterData.indexOf('管理者名'),
      maker        : filterData.indexOf('メーカー'),
      product      : filterData.indexOf('製品名'),
      model        : filterData.indexOf('モデル'),
      size         : filterData.indexOf('画面サイズ'),
      key          : filterData.indexOf('キー配列'),
      memory       : filterData.indexOf('メモリ'),
      ssd          : filterData.indexOf('SSD'),
      cpu          : filterData.indexOf('CPU'),
      note         : filterData.indexOf('備考'),
      statusEditDate  : filterData.indexOf('ステータス変更日'),
      statusEditedDate: filterData.indexOf('からの経過日'),
      statusEditName: filterData.indexOf('変更者'),
      endDate      : filterData.indexOf('レンタル終了日'),
      money        : filterData.indexOf('月額料金'),
      pcUri        : filterData.indexOf('台帳URL'),
      checkStatus  : filterData.indexOf('対応済み')
    };
    return this.index;
  }
}
  
RentalStokeListSheet.prototype = {
  getRowKey: function(target) {
    var targetIndex = this.getIndex()[target];
    var alfabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
    var returnKey = (targetIndex > -1) ? alfabet[targetIndex] : '';
    if (!returnKey || returnKey === '') showTitleError(target);
    return returnKey;
  },
  getIndex: function() {
    return Object.keys(this.index).length ? this.index : this.createIndex();
  },
  /**
   * jsonに基づき表に書き込む
   */
  editNewData: function(value, row) {
    var statusHistryAry = value.status_history.value.split(',');
    var data = [[
      value.pc_status.value,
      value.sub_status.value,
      value.capc_id.value,
      value.rentalid.value,
      value.location.value,
      value.user_name.value,
      value.pc_maker.value,
      value.pc_product.value,
      value.pc_model.value,
      value.pc_display.value,
      value.keyboard.value,
      value.memory.value,
      value.ssd.value,
      value.cpu.value,
      value.appendix.value,
      KintoneApi.caApi.api.getUri(value[KintoneApi.KEY_ID].value),
      statusHistryAry.length > 0 ? statusHistryAry[0] : '',
      getDataIf(this.getRowKey('statusEditDate') + row), // 本日からの差分
      statusHistryAry.length > 1 ? statusHistryAry[1] : ''
    ]];
    this.sheet.getRange(row, 1, data.length, data[0].length).setValues(data);
  }
};
var rentalStokeListSheet = new RentalStokeListSheet();

/**
 * 在庫のデータをアップロードする
 * ライブラリ SimplitData = https://script.google.com/a/cyberagent.co.jp/d/McQvjVTgsMLeWv2sRD57v3ttFCvUnFrd6/edit?mid=ACjPJvE11NxPrpy3J6s-jIi3gc4dW5SWXwMJoKJzNDSXnWdekMAUuHaYdX-FL_KWvrxRmbwhtS5ajHq3AomNEvlHPIDLQguMlOOAa2wJA_MHH86M7gsEswB-oF8SuUMY_khsJEkWbXzU-qA&uiv=2
 */
function updateStokeData() {
  var data = KintoneApi.caApi.getRentalNotUse([
    KintoneApi.KEY_ID, 'pc_status', 'sub_status', 'capc_id', 'location', 'rentalid', 'user_name', 'pc_maker', 'pc_product', 'pc_model', 'pc_display', 'keyboard', 'appendix', 'status_history', 'memory', 'ssd', 'cpu'
  ]);
  var r = rentalStokeListSheet;
  var simplitSheet = SimplitData.simplitCSVSheet;
  var simplitIndex = simplitSheet.getIndex();
  
  // シートに書かれたデータ('対応済'を除く)
  var editData = r.values
    .map(function(value, index) {
      if(!value[r.getIndex().checkStatus] && value[r.getIndex().caPcNo]) return { caPcNo: value[r.getIndex().caPcNo], statusEditedDate: value[r.getIndex().statusEditedDate],　row: index + 1 };
      else return null;
    }).filter(function(value) {
      return value != null;
    });
  
  data.forEach(function(value, index) {
    var no = value.capc_id.value;
    
    // すでに書かれているデータと一致したらreturn
    var editedIndex = editData.map(function(data, i) {
      if (data.caPcNo === no) return i;
      else return null;
    }).filter(function(data) {
      return data != null;
    });
    if(editedIndex.length > 0) {
      editData.splice(editedIndex[0], 1);
      return;
    }
    var lastRow = r.sheet.getRange('A:A').getValues().filter(String).length + 1;
    r.editNewData(value, lastRow);
    
    // simplitデータ取得
    var endDate = simplitSheet.getTargetData(value.rentalid.value);
    var endValue = 'simplitデータ無し'
    if (endDate) {
      endValue = endDate[simplitIndex.endDate];
      r.sheet.getRange(r.getRowKey('money') + lastRow).setValue(endDate[simplitIndex.money]);
    }
    r.sheet.getRange(r.getRowKey('endDate') + lastRow).setValue(endValue);
  });
  
  // すでに書かれていて,且つデータとして来なかったものはステータスを修正する
  editData.forEach(function(value) {
    r.sheet.getRange(r.getRowKey('statusEditedDate') + (value.row)).setValue(value.statusEditedDate);
    r.sheet.getRange(r.getRowKey('checkStatus') + (value.row)).setValue(true);
  });
  
  r.sheet.getRange('E1').setValue(Utilities.formatDate(new Date(), "JST", TEXT_DATE));
  r.sheet.getRange('E2').setValue(data.length + '台');
}
