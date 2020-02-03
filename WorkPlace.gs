/**
 * Workplaceに投稿する
 */
function postWorkPlaceFeed(text, token) {
  UrlFetchApp.fetch('https://graph.facebook.com/group/feed?', {
    method : 'POST',
    contentType : 'application/json',
    payload : JSON.stringify({
      message: text,
      formatting : 'MARKDOWN',
      access_token: token
    })
  });
}

/**
 * テストに投げるBOT
 * https://ca-group.facebook.com/groups/455956918282865/
 */
function postBotForTest(text) {
  const WORKPLACE_ACCESS_TOKEN__TEST = 'DQVJ2VTJlaVNTcUYyejl2MUNTcnZA3T3Jtb0M1Wmtfd3VsUVVzUldKYUpKNnd5dFZAISHBnRTh0ZAjQ4alE2eEE5dDNyTVJ4NUlnb0k0RWd0ajREMHI3RURMX1FCcV93TkVzekdtQmFSbGRNQU92dWViVVMzR1BmZAzBfYW82LUtQWmNSelRCVWt2RkNGeHI3b3NwckZAhMm5DdGdpNVItTWJ5RWY4dXpybHY5dy1EUG9FZAWJGN0FZAR2JqSldLWVM1LWVabEJqQ3RB';
  postWorkPlaceFeed(text, WORKPLACE_ACCESS_TOKEN__TEST);
}

/**
 * 全シス★資産管理チームに投げるBOT
 * https://ca-group.facebook.com/groups/2063765203720678/
 */
function postBotForArms(text) {
  const WORKPLACE_ACCESS_TOKEN__ARMS = 'DQVJ0V25jY0xoVGlTdmNYNnRQN0ZAfdGYxN3Y2NWxRa3paN0JhVVZAQUlE2QW5zaFdXeHZAPLUZA5alF5VTZAqTzcyZAzVuT2IzbVhXZAVpyb1hpWHc3QnM1WlRqTHA0SHNVUWUtS2hOb2dTbzcwMFNfRjNyMUdOeEtsam81aW44UVNuVUt1WV9kel9Kd2ZACWDNUQURES3ZANbVM1MTlNZA2lYd3N5OW1taTlqUHRKVUNZAdV9Db2hmQ1ZAvSEQ3RS0wQWdqUmpQRnRFaUdGaEJfVFJCMDJiOAZDZD';
  postWorkPlaceFeed(text, WORKPLACE_ACCESS_TOKEN__ARMS);
}
