///////////////////////////////////////////////////////////////////////
////////////////////////// kintoneのAPIにアクセスする /////////////////////
// kintone仕様書 https://developer.cybozu.io/hc/ja/articles/201941754 //
// 台帳仕様書 https://docs.google.com/spreadsheets/d/1YjgrihpwSU1XuU4dC-4dMaAcQz2aMBOMS6see5__zVo/edit#gid=820501261
///////////////////////////////////////////////////////////////////////

// ↓test
//curl -X 'GET' 'https://zqt3h.cybozu.com/k/v1/record.json?app=9805&id=439203' -H 'X-Cybozu-API-Token: 0Dduu8lXkWCOa4oRhOTVI0voO3o14RSmXeAxYny9' 

var KintoneApiV1 = function() {
  this.allRecords = [];
}
KintoneApiV1.prototype = {
  /**
   * 最新のデータを取得
   * @return {string} 最新のPC情報
   
  getAllRecords: function() {
    if (this.allRecords.length === 0) {
      // いらないかも
      var response = UrlFetchApp.fetch(KINTONE_API_URI + 'records.json?app=' + KINTONE_APP_ID, {
        method : 'get',
        headers: {
          'X-Cybozu-API-Token': KINTONE_API_TOKEN,
          'Authorization': 'Basic ' + KINTONE_API_TOKEN
        }
      });
      var json = JSON.parse(response.getContentText());
      this.allRecords = json.records;
    }
    return this.allRecords;
  },*/
  /**
   * getしてjsonに変換して返す
   * @param {string} query用文字列
   * @param {[string]} fields用文字列の配列
   */
  get: function(queryStr, fieldsAry) {
    var query = encodeURIComponent(queryStr);
    var fields = fieldsAry.length > 0 ? '&' + fieldsAry.map(function(str, i) { return 'fields[' + i + ']=' + encodeURIComponent(str) }).join('&') : '';
    var response = UrlFetchApp.fetch(KINTONE_API_URI + 'records.json?app=' + KINTONE_APP_ID + '&query=' + query + fields, {
      method: 'get',
      headers: {
        'X-Cybozu-API-Token': KINTONE_API_TOKEN,
        'Authorization': 'Basic ' + KINTONE_API_TOKEN
      }
    });
    var json = JSON.parse(response.getContentText());
    //Logger.log(json);
    return json.records;
  },
  /**
   * レンタルPCのステータスが利用中以外のものを返す
   */
  getRentalNotUse: function() {
    var query = QUERY_IS_RENTAL + ' and status not in ("利用中（共有PC）", "利用中（個人）", "紛失", "CA本体管理対象外") and ' + QUERY_MY_ASSIGNED;
    var fields = ['レコード番号', 'no', 'pc_management_no', 'rental_pc_management_no', 'management_employee_no', 'management_employee_name', 'management_employee_division', 'status', 'inventory_base', 'assigned_division', 'maker', 'product_no', 'note', 'model_no', 'registed_at', 'ocs_last_updated_at', 'maam_last_connected_at'];
    var data = this.get(query, fields);
    return data;
  }
};
var kintoneApiV1 = new KintoneApiV1();

/**
 * テスト用：適当なIDのを１レコード取得
 */
function testKinton() {
  var response = UrlFetchApp.fetch(KINTONE_API_URI + 'record.json?app=' + KINTONE_APP_ID + '&id=439189', {
    method : 'get',
    headers: {
      'X-Cybozu-API-Token': KINTONE_API_TOKEN,
      'Authorization': 'Basic ' + KINTONE_API_TOKEN
    }
  });
  var json = JSON.parse(response.getContentText());
  Logger.log(json.record); // PC管理番号
}