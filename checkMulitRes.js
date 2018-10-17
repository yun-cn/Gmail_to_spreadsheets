function checkMultRes() {


  var immediately_reservation_title = "label:今すぐ予約"
  var booking_is_done_title = "label:予約完了"

  var immediately_reservation = '(' + immediately_reservation_title + ' after: ' + today() + ')'
  var booking_is_done = '(' + booking_is_done_title + ' after: ' + today() + ')'
  var target_query = immediately_reservation + 'OR' + booking_is_done
  var subRegExp = /^(【*.*】)+(★)+([0-9A-Z]{7})/
  Logger.log(target_query)
  // var TEST  = "203898"
  var target_threads = GmailApp.search(target_query)

  if (target_threads.length !== 0) {
    for (var i = 0; i < target_threads.length; i++) {
      var messages = target_threads[i].getMessages()
      for (var m = 0; m < messages.length; m++) {
        var roomCodeTemp = messages[m].getSubject().match(subRegExp)
        if (roomCodeTemp === null) {
          continue;
        }
        var roomCode = roomCodeTemp[3]
        // Logger.log(messages[m].getSubject().match(subRegExp)[2])
        var msg = messages[m].getPlainBody()
        var getMessageTime = Utilities.formatDate(messages[m].getDate(), 'Asia/Tokyo', 'yyyy/MM/dd');
        if (getMessageTime !== today() || messages[m].isStarred() === true)  {
          continue;
        }
        var d = getDatabyMailBody(msg)
        if (d.reference.length >= 2) {
          var messageBody = "```" +  messages[m].getSubject() + '  ' + d.bookingId + '  ' + d.reference + "```"
          postToSlack(messageBody)
          Logger.log(messages[m])
          messages[m].star()
        }
      }
    }
    // var messages = target_threads[1].getMessages()
    // Logger.log(messages[0].getPlainBody())
    // Logger.log(getDatabyMailBody(messages[0].getPlainBody()))
  }
}

function yesterday(){
  var today = new Date();
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  var yesterday = new Date(today.getTime() - MILLIS_PER_DAY);
  var startDate = Utilities.formatDate(yesterday, 'Asia/Tokyo', 'yyyy/MM/dd');
  return startDate;
}

function today() {
  var today = new Date();
  var yesterday = new Date(today.getTime());
  var endDate = Utilities.formatDate(yesterday, 'Asia/Tokyo', 'yyyy/MM/dd');
  return endDate;
}

function getDatabyMailBody( body ) {
  var bookingId = body.match(/予約ID.*/).toString().replace(/[^0-9]/g, "")
  var guestOfNumber = body.match(/人数.*/).toString().replace(/[^0-9]/g, "")
  var usedTimeData
  var usedTempDate = body.match(/利用期間.*/).toString().split('、')
  var referenceDate
  // Logger.log(usedTempDate)
  // Logger.log(usedTempDate.length)
  if (usedTempDate.length >= 2 ) {
    referenceDate = body.match(/利用期間.*/).toString()
    usedTimeData = body.match(/利用期間.*/).toString().split('、')[0]
  } else {
    usedTimeData = body.match(/利用期間.*/).toString()
  }
  var usageTime = usedTimeData.split('〜')
  var myRegExp= /[0-9]{4}\/(0[1-9]|1[0-2]|[0-9])\/(0[1-9]|[1-2][0-9]|3[0-1]|[0-9]) (2[0-3]|[1][0-9]|[0-9]):[0-5][0-9]/
  var checkInDate = usageTime[0].match(myRegExp)
  var checkIn = Utilities.formatDate(new Date(checkInDate[0]), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss')
  var checkInDate = new Date(checkIn)
  var checkOutDate
  var checkOut = usageTime[1].trim()
  if (checkOut.length <= 6) {
    checkOutH = checkOut.split(':')
    checkOutDateTmp = Utilities.formatDate(new Date(checkInDate.getFullYear(), checkInDate.getMonth(), checkInDate.getDate(), checkOutH[0], checkOutH[1], '0'), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss')
    checkOutDate = new Date(checkOutDateTmp).getHours() === 0.0 ? new Date((new Date(checkOutDateTmp)).getTime() + (24 * 60 * 60 * 1000)) : new Date(checkOutDateTmp)
  } else {
    var checkOutDate = usageTime[1].match(myRegExp)
    var checkOut = Utilities.formatDate(new Date(checkOutDate[0]), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss')
    checkOutDate =  new Date(checkOut).getHours() === 0.0 ? new Date((new Date(checkOut)).getTime() + (24 * 60 * 60 * 1000)) : new Date(checkOut)
  }


  if (body.match(/設備・サービス.*/)) {
    var facilities =body.match(/設備・サービス.*/).toString().replace(/設備・サービス/, "").split('、')
   }
  // if (facilities) {
  //   Logger.log(facilities.split('、')[0])
  return {
      bookingId: bookingId,
      guestOfNumber: guestOfNumber,
      // checkInYear: checkInDate.getFullYear() + '年',
      // checkInMonth: checkInDate.getMonth() + 1 + '月',
      // checkInDay: checkInDate.getDate() + '日',
      // checkOutYear: checkOutDate.getFullYear() + '年',
      // checkOutMonth: checkOutDate.getMonth() + 1 + '月',
      // checkOutDay: checkOutDate.getDate() + '日',
      checkInData: Utilities.formatDate(checkInDate,Session.getScriptTimeZone(), 'yyyy/MM/dd'),
      checkOutData: Utilities.formatDate(checkOutDate,Session.getScriptTimeZone(), 'yyyy/MM/dd'),
      checkInTime: (checkInDate.getHours() === 0.0 ? '00' : checkInDate.getHours()) + ':' + '00',
      checkOutTime: (checkOutDate.getHours() === 0.0 ? '24' : checkOutDate.getHours())  + ':' + '00',
      usedTime: Math.abs((checkOutDate - checkInDate) / 36e5),
      reference: referenceDate || ' ' ,
      facilities: facilities || []
   }
 }


 function postToSlack(messageBody) {
   // var postUrl = 'xx';
   //var postUrl = 'xxx';
   var username = 'botbot';  // 通知時に表示されるユーザー名
   var icon = ':hatching_chick:';  // 通知時に表示されるアイコン
   var message = messageBody

   var jsonData =
   {
      "username" : username,
      "icon_emoji": icon,
      "text" : message
   };
   var payload = JSON.stringify(jsonData);

   var options =
   {
     "method" : "post",
     "contentType" : "application/json",
     "payload" : payload
   };

   UrlFetchApp.fetch(postUrl, options);
 }
