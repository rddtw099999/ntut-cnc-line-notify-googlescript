//==========================================================================================================================================//
//                                                                                                                                          //
//          @Author : Minkai Chen                                                                                                           //
//          @Date   : 2020/08/13                                                                                                            //
//          @Title  : Line通知小Script，透過getDate()+1 爬取班表相符格子往下5格進行通知                                                            //
//          @Edit   : 08/17 @Minkai Chen : 修正 通知時若不足5人，後面會多出逗號的issues                                                           //
//                                       : 修正 新增方法sendMessage() 將訊息通知改為直接向Line Server發送 POST Method ， 不再使用Heroku webhook     //
//                                       : 報廢 透過 Heroku Webhook 的Get Method                                                              //
//                                       : 微調 通知格式                                                                                      //
//                    08/27 @Minkai Chen : Add Post Image Method                                                                            //
//                                         Test send image with Imgur link and .png extension is success.                                   //
//                    08/28 @Minkai Chen : Add Send Sticker Method                                                                          //
//                                         Sticker Package Id and Sticker Id can be find here                                              //
//                                            https://devdocs.line.me/files/sticker_list.pdf                                                //
//========================================================================================================================================= //
var token = '<發行權杖>'   // 群組 "計中工讀生" 發行權杖  


function doPostMethod() {
  var SpreadSheet = SpreadsheetApp.openById("<試算表ID>");
  var Sheet = SpreadSheet.getSheetByName("工作表1");
  var testDate = new Date();
  var secondDate = new Date();
  secondDate.setDate(testDate.getDate()+1);
  var datestr=Utilities.formatDate(secondDate, "GMT+8", "M/dd")  //轉換為班表之日期格式 (月/日日)
  var data = Sheet.getDataRange().getValues();
  var cols = Sheet.getLastColumn();
  var nametable=[];
  data.forEach(function(e, i){            //迴圈模糊搜尋整個試算表的隔天日期
    for (var j=cols-1; j>=0; j--) {
      var xx=e[j].toString()
      if (xx.indexOf(datestr.toString())>-1) {  //若字串內包含該天日期格式字串
        for (var x=i+2; x<=i+6; x++){
          if(Sheet.getSheetValues((x),(j+1), 1,1)!=''){
            nametable.push(Sheet.getSheetValues((x),(j+1), 1,1))
          }         
        }
        msg= "\n明天 (" + datestr +")，值班的工讀生是:\n[" + nametable.toString() + "]\n請記得準時到班窩~"
        sendMessage(msg)
        break;
      }
    }
  }); 
}


function doSendImage(){
  sendImage('<訊息>', '<圖片URL>');
}

function doSendSticker(){
  sendSticker('<貼圖訊息>','<貼圖PackageId>','<貼圖編號>');
}


function sendMessage(message) {
  var option = {
    method: 'post',
    headers: { Authorization: 'Bearer ' + token },
    payload: {
      message: message
    }
  };
  UrlFetchApp.fetch('https://notify-api.line.me/api/notify', option);
}


function sendImage(message,filelink) {
  var option = {
    method: 'post',
    headers: { Authorization: 'Bearer ' + token },
    payload: {
      message:message,
      imageThumbnail : filelink,
      imageFullsize : filelink,
      imagefile : filelink
    }
  };
  UrlFetchApp.fetch('https://notify-api.line.me/api/notify', option);
}

function sendSticker(msg,pkgid,id) {
  var option = {
    method: 'post',
    headers: { Authorization: 'Bearer ' + token },
    payload: {
      message:msg,
      stickerPackageId:pkgid ,
      stickerId:id
    }
  };
  UrlFetchApp.fetch('https://notify-api.line.me/api/notify', option);
}

