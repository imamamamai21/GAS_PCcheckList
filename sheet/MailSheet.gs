/**
 * メールテンプレートシート
 */
var mailTemplateSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('メールテンプレート');
  this.values = this.sheet.getDataRange().getValues();
  this.titleRow = 2;
  this.index = {};
  
  this.createIndex = function() {
    var filterData = this.values[this.titleRow];
    this.index = {
      returnRemind: filterData.indexOf('返却のお願い')
    };
    return this.index;
  };
  
  this.getIndex = function() {
    return Object.keys(this.index).length ? this.index : this.createIndex();
  },

  /**
   * テンプレを返す
   * @param type{string} = index: 'returnRemind'
   */
  this.getTemplate = function(type) {
    return {
      title: this.values[this.titleRow + 3][this.getIndex()[type]],
      text: this.values[this.titleRow + 4][this.getIndex()[type]] + this.getFotter()
    };
  };
  this.getFotter = function() {
   return this.values[9][1];
  };
  this.getMailAdress = function() {
   return this.values[8][1];
  };
  /**
   * メール送信
   * @param to string 送信先
   * @param cc string CC送信先
   * @param title string 件名
   * @param text string 本文
   * @param needCheck boolean 確認popupを出すか否か
   */
  this.sendMail = function(to, cc, title, text, needCheck) {
    if (needCheck && !this.openLastCheckPopup(text)) return false;
    var adress = mailTemplateSheet.getMailAdress();
    
    GmailApp.sendEmail(to, title, text, {
      from: adress,
      replyTo: adress, 
      cc: cc,
      name: '資産管理チーム'
    });
    return true;
  };
  
  /**
   * メールを送ります
   * @return boolean
   */
  this.openLastCheckPopup = function(text) {
    var popup = Browser.msgBox('以下の内容でメールを送信します。よろしいでしょうか？(実際には改行されます)', text, Browser.Buttons.OK_CANCEL);
    return popup === 'ok';
  }
}

var mailTemplateSheet = new mailTemplateSheet();
