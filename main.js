
//　メールフォーマット
//　日時　　：2020年06月01日（月） 10:00 ～ 2020年06月01日（月） 13:00
//　予定　　：会議:【20卒研修】Linux講座
//　参加者　：仲澤 義広,神山 和久,岩﨑 航大,南 総一郎,坂元 法敦,白幡 淳平,赤尾 慎太朗,秋庭 理恵,白木 晴也,畑 彪真,早川 智之,松野 竜也,三好 諒

// メールフォーマット（毎日設定）
// 日時　　：毎日（土日を除く） 10:00 ～ 13:00
// 期間　　：2020年06月10日（水）～2020年06月16日（火）
// 予定　　：会議:【20卒研修】ネットワーク講座
// 参加者　：南 総一郎,坂元 法敦,白幡 淳平,神山 和久,永井 翔平,秋庭 理恵,赤尾 慎太朗,白木 晴也,畑 彪真,早川 智之,松野 竜也,三好 諒

const START = 0;
const MAX = 20;

function main() {

  const sheet = getSpreadSheet();
  const hoge = hoge();

}

function hoge() {

}

function getSpreadSheet() {
  const objSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const objSheet = objSpreadsheet.getSheetByName("シート1");//シート名をここに入力
  const sheet = SpreadsheetApp.setActiveSheet(objSheet);
  return sheet;
}

function getMails() {
  // メールのラベルを指定して検索
  const mails = GmailApp.search('label:door is:unread', START, MAX);
  const row = sheet.getLastRow() + 1;//最後の行探してそれ以降に追加
}

function myFunction() {

  // 読み込み件数の数
  const start = 0;
  const max = 20;



  for (let n in mails) {

    const mail = mails[n];
    const msgs = mail.getMessages();

    for (let m in msgs) {
      const msg = msgs[m];
      // 本文
      const body = msg.getPlainBody();
      // メールタイトル
      const subject = msg.getSubject();

      if (!subject.match(/変更/)) {

        // 改行を使って配列化する
        const bodyAry = body.split("\n");

        // 日時の行を取得し「日時　　：」を取り除いて文字を抽出
        const dateText = bodyAry[0].substring(0).replace('日時　　：', '');

        // 毎日設定の場合
        if (dateText.match(/毎日/)) {

          const objects = everyday(bodyAry, dateText);

          // シートに登録
          sheet.getRange(row, 1).setValue(objects[0]);
          sheet.getRange(row, 2).setValue(objects[1]);
          sheet.getRange(row, 3).setValue(objects[2]);

        } else {

          const objects = aday(bodyAry, dateText);

          sheet.getRange(row, 1).setValue(objects[0]);
          sheet.getRange(row, 2).setValue(objects[1]);
          sheet.getRange(row, 3).setValue(objects[2]);

        }

        row++;

      }
    }

    Utilities.sleep(1000);
    mail.markRead();
    mail.moveToArchive();
  }

}

function convDate(date) {
  // dates[0]=2020年06月10日
  // 開始日と終了日を"2020-6-10"みたいな形にする

  const year = date.slice(0, 4);
  const month = date.slice(5, 7);
  const day = date.slice(8, 10);

  return new Date(year, month - 1, day);

}

function everyday(bodyAry, dateText) {
  const time = dateText.slice(-14);
  Logger.log(time);
  //10:00 ～ 13:00 こんなのが出るはず

  const times = time.split(" ～ ");
  //times[0]=10:00 times[1]=13:00になってるはず
  Logger.log(times[0]);
  Logger.log(times[1]);

  const date = dateText = bodyAry[1].substring(0).replace('期間　　：', '');

  const dates = date.split("～");
  //dates[0]=2020年06月10日（水）date[1]=2020年06月16日（火）期待できる出力

  // 開始日と終了日を"2020-5-10"みたいな形にする
  const convStartDate = convDate(dates[0]);
  const convEndDate = convDate(dates[1]);

  for (let d = convStartDate; d <= convEndDate; d.setDate(d.getDate() + 1)) {

    const schedule = bodyAry[2].substring(0).replace('予定　　：', '');
    const startDate = combDate(d, times[0]);
    const endDate = combDate(d, times[1]);

    regiEvent(schedule, startDate, endDate);
  }

  // スプシに登録するために必要
  const rStartDate = everydaySplitDate(dates[0], times[0]);
  const rEndDate = everydaySplitDate(dates[1], times[1]);

  const schedule = bodyAry[2].substring(0).replace('予定　　：', '');

  return [schedule, rStartDate, rEndDate];

}

function combDate(date, time) {

  const hour = time.slice(0, 2);
  const minutes = time.slice(3, 5);

  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), hour, minutes);
}

function everydaySplitDate(getDate, getTime) {

  const year = getDate.slice(0, 4);

  const month = getDate.slice(5, 7);

  const day = getDate.slice(8, 10);

  const hour = getTime.slice(0, 2);

  const minutes = getTime.slice(3, 5);

  return new Date(year, month - 1, day, hour, minutes);

}

function aday(bodyAry, dateText) {
  // その日のみの時
  // 開始日時・時刻と終了日時・時刻を「 ～ 」で分ける
  const dates = dateText.split(" ～ ");

  // 開始日時をフォーマット
  const startDate = adaySplitDate(dates[0].substring(0));
  Logger.log(startDate);

  // 終了時刻をフォーマット
  const endDate = adaySplitDate(dates[1].substring(0));
  Logger.log(endDate);

  // 予定の行を取得し「予定　　：」を取り除いて文字を抽出
  const schedule = bodyAry[1].substring(0).replace('予定　　：', '');

  // カレンダーに登録
  regiEvent(schedule, startDate, endDate);

  return [schedule, startDate, endDate];

}

function adaySplitDate(getDate) {
  const year = getDate.slice(0, 4);

  const month = getDate.slice(5, 7);

  const day = getDate.slice(8, 10);

  const hour = getDate.slice(15, 17);

  const minutes = getDate.slice(18, 20);

  return new Date(year, month - 1, day, hour, minutes);
}

function regiEvent(getSche, get_sdate, get_edate) {
  // 連携するアカウント
  const gAccount = "hayakawat@istyle.co.jp";

  // googleカレンダーの取得
  const calender = CalendarApp.getCalendarById(gAccount);

  // 予定を作成
  calender.createEvent(
    getSche,
    get_sdate,
    get_edate,
  );
}