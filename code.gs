// リソース→ライブラリで下記を追加する必要があります
//   Moment.js：MHMchiX6c1bwSqGM1PZiW_PxhMjh3Sh48

var EVENT_NAME = '社内勉強会';
var EVENT_CALENDAR_ID = '{メールアドレス}';
var TEMPLATE_FORM_ID = '{フォームのID}';
var BASE_FOLDER_ID = '{フォーム出力先のドライブフォルダID}';
var WEB_PAGE_URL = '{このスクリプトのウェブアプリケーションURL}'

var MAIL_TO = '{宛先To}';
var MAIL_CC = '{宛先Cc}';
var REPLY_TO = '{返信先}';
var FROM_NAME = EVENT_NAME + '実行委員';

var DO_SEND_MAIL = true;

function cron() {
  // 今月のイベント取得
  var event = getEventForThisMonth();
  if (!event) {
    // なければ終了
    Logger.log('There is no event');
    return;
  }
  
  // イベントの日付取得
  Logger.log('event date:' + event.getStartTime());
  var m_event_date = Moment.moment(event.getStartTime());
  var event_date = m_event_date.format('YYYY/MM/DD');

  // イベント日付のフォルダ取得（なければ作る）
  Logger.log('event date:' + event_date);
  var folder_name = event_date;
  var folder = getFolder(folder_name);
  if (!folder) {
    folder = createFolder(folder_name);
    Logger.log('folder created:' + folder.getId());
  }

  // フォーム取得（なければ作る）
  var form_name = EVENT_NAME + '_' + event_date;
  var form = getForm(form_name, folder);
  if (!form) {
    form = createForm(form_name, folder);
    Logger.log('form created:' + form.getId());
    
    // 告知メール送信
    sendOpenMail(form, event);
  }
  Logger.log('folder:' + folder.getId());
  Logger.log('form:' + form.getId());

  // 3日前に開催有無判定
  if (isDaysAgo(3, m_event_date)) {
    // 発表登録あれば
    if (hasPresenter(form)) {
      // 開催メール送信
      sendDetailMail(form, event);

    } else {
      cancelEvent(form, event);
    }
  }
  
  // 前日
  if (isDaysAgo(1, m_event_date)) {
    // 発表登録あれば
    if (hasPresenter(form)) {
      // リマインドメール送信（しつこいのでコメントアウト）
      // sendRemindMail(form, event);
    } else {
      // 3日前までは発表登録があったが直前キャンセルとなった場合は中止。
      cancelEvent(form, event);
    }
  }
}

function isDaysAgo(days, moment) {
  
  var days_ago = moment.diff(Moment.moment(), 'days');
  Logger.log('days:' + days + ' days_ago:' + days_ago);
  if (days_ago == days) {
    return true;
  }
  return false;
}

function hasPresenter(form) {
  var formResponses = form.getResponses();
  for (var i = 0; i < formResponses.length; i++) {
    var formResponse = formResponses[i];
    var itemResponses = formResponse.getItemResponses();
    for (var j = 0; j < itemResponses.length; j++) {
      var itemResponse = itemResponses[j];
      if (itemResponse.getItem().getTitle() == '発表しますか？' && itemResponse.getResponse() == '発表する') {
        return true;
      }
    }
  }
  return false;
}

function getEventForThisMonth() {
  var first_date_of_this_month = Moment.moment().startOf('month');
  var last_date_of_this_month = Moment.moment().endOf('month');

  var calender = CalendarApp.getCalendarById(EVENT_CALENDAR_ID);
  var events = calender.getEvents(new Date(first_date_of_this_month.format('YYYY/MM/DD')), new Date(last_date_of_this_month.format('YYYY/MM/DD')), {search: EVENT_NAME});

  for(var i = 0; i < events.length; i++) {
    // 名称が完全一致した最初のイベントを返す
    if (events[i].getTitle() == EVENT_NAME) {
      return events[i];
    }
  }
  return null;
}

function getForm(form_name, folder) {
  var forms = DriveApp.getFolderById(folder.getId()).getFilesByName(form_name);
  while (forms.hasNext()) {
    // 一つ目のフォームを返す（複数存在した場合は考慮していない）
    var form = FormApp.openById(forms.next().getId());
    return form;
  }
  return null;
}

function createForm(form_name, folder) {

  // フォーム作る
  var template = DriveApp.getFileById(TEMPLATE_FORM_ID);
  var file = template.makeCopy(form_name, folder);
  var form = FormApp.openById(file.getId());
  form.setTitle(form_name);

  return form;
    
}

function getFolder(folder_name) {
  var folders = DriveApp.getFolderById(BASE_FOLDER_ID).getFoldersByName(folder_name);
 　　while (folders.hasNext()) {
    // 一つ目のフォルダIDを返す（複数存在した場合は考慮していない）
    var folder = folders.next();
    return folder;
  }
  return null;
}

function createFolder(folder_name) {
  var folder = DriveApp.getFolderById(BASE_FOLDER_ID).createFolder(folder_name);
  return folder;
}

function sendOpenMail(form, event) {
  var subject = '【' + EVENT_NAME + '】' + Moment.moment(event.getStartTime()).format('YYYY/MM/DD') + ' 参加受付開始！';
  var body_lines = [];
  body_lines.push('各位');
  body_lines.push('');
  body_lines.push('お疲れ様です。' + FROM_NAME + 'です。');
  body_lines.push('');
  body_lines.push('下記日程で' + EVENT_NAME + 'を開催予定です。');
  body_lines.push('');
  body_lines.push('日時：' + Moment.moment(event.getStartTime()).format('M月D日 HH:mm') + ' ~ ' + Moment.moment(event.getEndTime()).format('HH:mm'));
  body_lines.push('場所：' + event.getLocation());
  body_lines.push('');
  body_lines.push('参加される方は下記より登録お願い致します。');
  body_lines.push(form.getPublishedUrl());
  body_lines.push('');
  body_lines.push('以上、よろしくお願いいたします。');
  
  var body = body_lines.join("\n");

  sendMail(subject, body, MAIL_TO);
}

function sendDetailMail(form, event, subject_prefix) {
  var subject_prefix = typeof subject_prefix !== 'undefined' ?  subject_prefix : '';
  var subject = subject_prefix + '【' + EVENT_NAME + '】' + Moment.moment(event.getStartTime()).format('YYYY/MM/DD') + ' 開催！';
  var body_lines = [];
  body_lines.push('各位');
  body_lines.push('');
  body_lines.push('お疲れ様です。' + FROM_NAME + 'です。');
  body_lines.push('');
  body_lines.push(Moment.moment(event.getStartTime()).format('M/D') + 'の' + EVENT_NAME + 'について');
  body_lines.push('WEBページを用意しました。');
  body_lines.push(WEB_PAGE_URL + '?date=' + Moment.moment(event.getStartTime()).format('YYYY/MM/DD'));
  body_lines.push('');
  body_lines.push('引き続き参加者募集中です。下記より登録お願い致します。');
  body_lines.push(form.getPublishedUrl());
  body_lines.push('');
  body_lines.push('以上、よろしくお願いいたします。');

  var body = body_lines.join("\n");

  sendMail(subject, body, MAIL_TO);
}

function sendRemindMail(form, event) {
  sendDetailMail(form, event, '【リマインド】');
}

function sendCancelMail(form, event) {
  var subject = '【' + EVENT_NAME + '】' + Moment.moment(event.getStartTime()).format('YYYY/MM/DD') + ' 中止のお知らせ';
  var body_lines = [];
  body_lines.push('各位');
  body_lines.push('');
  body_lines.push('お疲れ様です。' + FROM_NAME + 'です。');
  body_lines.push('');
  body_lines.push(Moment.moment(event.getStartTime()).format('M/D') + 'に' + EVENT_NAME + 'を予定しておりましたが、');
  body_lines.push('発表者不足のため中止とさせていただきます。');
  body_lines.push('');
  body_lines.push('以上、よろしくお願いいたします。');

  var body = body_lines.join("\n");

  sendMail(subject, body, MAIL_TO);
}

function sendMail(subject, body, mail_to) {
  if (!DO_SEND_MAIL) {
    return;
  }
  
  var options = {
//    from: REPLY_TO, // 変えたいけどエラーになっちゃう
    cc: MAIL_CC,
    replyTo: REPLY_TO,
    name: FROM_NAME
  };

  GmailApp.sendEmail(mail_to, subject, body, options);
}

function cancelEvent(form, event) {
  // 中止メール送信
  sendCancelMail(form, event);
  // フォームを閉じる（回答不可にする）
  closeForm(form);
  // カレンダーから予定削除
  event.deleteEvent();
}

function closeForm(form) {
  form.setAcceptingResponses(false);
}

function doGet(e) {

  var date = e.parameter.date;
  if (typeof date == 'undefined') {
    return HtmlService.createTemplateFromFile("404").evaluate();
  }

  // フォーム取得
  var folder_name = date;
  var folder = getFolder(folder_name);
  if (!folder) {
    return HtmlService.createTemplateFromFile("404").evaluate();
  }
  var form_name = EVENT_NAME + '_' + date;
  var form = getForm(form_name, folder);
  if (!form) {
    return HtmlService.createTemplateFromFile("404").evaluate();
  }

  // フォームの回答結果を取得
  var formResponses = form.getResponses();
  var lines = [];
  for (var i = 0; i < formResponses.length; i++) {
    var line = {};
    var formResponse = formResponses[i];
    var itemResponses = formResponse.getItemResponses();
    for (var j = 0; j < itemResponses.length; j++) {
      var itemResponse = itemResponses[j];
      var title = itemResponse.getItem().getTitle();
      var response = itemResponse.getResponse();
      line[title] = response;
    }

    lines.push(line);
  }

  // 回答結果を表示用に整形
  var presenterLines = [];
  var participant_count = 0;
  for (var i = 0; i < lines.length; i++) {
    // 発表者情報
    if (lines[i]['発表しますか？'] == '発表する') {
      presenterLines.push(lines[i]);
    }

    // 参加人数
    if (lines[i]['参加しますか？'] == '参加する') {
      participant_count++;
    }
  }

  var template = HtmlService.createTemplateFromFile("index");
  template.title = form_name;
  template.date = e.parameter.date;
  template.participant_count = participant_count;
  template.presenters = presenterLines;
  
  return template.evaluate();
}
