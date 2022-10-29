// ScriptPropertiesのキー定数
const PROPERTY_KEY_SYNC_TOKEN = 'SYNC_TOKEN';
const PROPERTY_KEY_LINE_TOKEN = 'LINE_NOTIFY_ACCESS_TOKEN';

// LINE Notify APIエンドポイント定数
const ENDPOINT_LINE_NOTIFY_API = 'https://notify-api.line.me/api/notify';

// ScriptProperties
const properties = PropertiesService.getScriptProperties();

// Google Driveに保存されている予定全件リストのファイルID定数
const FILE_ID_EVENTS = properties.getProperty('SPREADSHEET_FILE_ID');

let nextSyncToken = '';

/**
 * メイン関数
 * カレンダー変更にトリガーされて起動する
 *
 * @param event カレンダー変更イベント情報
 */
function calendarUpdated(event) {
    console.time('----- calendarUpdated -----');

    try {
        // 変更されたカレンダーのIDを取得する
        const calendarId = event.calendarId;

        // syncTokenを用いて予定の差分リストを取得する
        const options = {
            syncToken: getSyncToken(calendarId)
        };
        const recentlyUpdatedEvents = getRecentlyUpdatedEvents(calendarId, options);

        // 予定の差分リストからLINE通知用のメッセージを生成する
        const message = generateMessage(recentlyUpdatedEvents);

        // LINEへ通知する
        notifyLINE(message);

        // スプレッドシートに保存した予定の全件リストを最新化する
        refleshStoredEvents(calendarId);

        // 次回起動用にsyncTokenを保持する
        properties.setProperty(PROPERTY_KEY_SYNC_TOKEN, nextSyncToken);
    } catch (e) {
        console.error(e);

        // エラー情報をLINEへ通知する
        const error = '\nカレンダー変更イベントの処理中にエラーが発生しました。\n\nError: ' + e.message;
        notifyLINE(error);
    }
    console.timeEnd('----- calendarUpdated -----');
}

/**
 * syncToken取得関数
 *
 * @param calendarId 対象カレンダーのID
 */
function getSyncToken(calendarId) {
    console.time('----- getSyncToken -----');

    // 前回起動時に保持したsyncTokenを取得する
    let token = properties.getProperty(PROPERTY_KEY_SYNC_TOKEN);

    // 前回起動時に保持したsyncTokenがない場合、予定リストから取得する
    if (!token) {
        token = Calendar.Events.list(calendarId, { 'timeMin': (new Date()).toISOString() }).nextSyncToken;
    }

    console.timeEnd('----- getSyncToken -----');
    return token;
}

/**
 * 予定の差分リスト取得関数
 *
 * @param calendarId 対象カレンダーのID
 * @param options    予定リスト取得リクエストのオプション
 */
function getRecentlyUpdatedEvents(calendarId, options) {
    console.time('----- getRecentlyUpdatedEvents -----');

    // 予定の差分リストを取得する
    const events = Calendar.Events.list(calendarId, options);

    // 次回起動用にsyncTokenを取得する
    nextSyncToken = events.nextSyncToken;

    console.timeEnd('----- getRecentlyUpdatedEvents -----');
    return events.items;
}

/**
 * LINE通知用メッセージ生成関数
 *
 * @param events 予定の差分リスト
 */
function generateMessage(events) {
    console.time('----- generateNotifyMessages -----');

    let message = '';
    const messages = [];
    for (let i = 0; i < events.length; i++) {
        // 予定差分のステータスを取得する
        const status = events[i].status;

        // ファイルに保存した予定の全件リストをIDで検索する
        const storedEvent = searchStoredEventById(events[i].id);

        let commonMsg = '';

        if (status == 'cancelled') { // 予定が削除された場合
            if (storedEvent) { // 予定の全件リストにIDが一致する予定が存在した場合
                const strStart = dateToString(storedEvent.start);
                const strEnd = dateToString(storedEvent.end);
                if (strStart.split(' ')[1]) {
                    commonMsg = (`\n\n${storedEvent.summary}\n${strStart}~\n${strEnd}`);
                } else {
                    commonMsg = (`\n\n${storedEvent.summary}\n${strStart} 終日`);
                    const strEndComp = dateComparison(strStart, strEnd);
                    if (strEndComp) {
                        commonMsg += `~\n${strEndComp} 終日`;
                    }
                }
                messages.push(`\nGoogleカレンダーの予定が削除されました。${commonMsg}`);
            } else { // 予定の全件リストにIDが一致する予定が存在しない場合
                messages.push('\nGoogleカレンダーの予定が削除されました。');
            }
        } else { // 予定が登録or更新された場合
            // 予定のdateもしくはdateTimeから予定の開始日時／終了日時を取得する
            const start = (events[i].start.dateTime) ? events[i].start.dateTime : events[i].start.date;
            const end = (events[i].end.dateTime) ? events[i].end.dateTime : events[i].end.date;
            const strStart = dateToString(start);
            const strEnd = dateToString(end);
            if (strStart.split(' ')[1]) {
                commonMsg = `\n\n${events[i].summary}\n${strStart}~\n${strEnd}`;
            } else {
                commonMsg = `\n\n${events[i].summary}\n${strStart} 終日`;
                const strEndComp = dateComparison(strStart, strEnd);
                if (strEndComp) {
                    commonMsg += `~\n${strEndComp} 終日`;
                }
            }
            if (storedEvent) { // 予定が更新された場合
                messages.push(`\nGoogleカレンダーの予定が更新されました。${commonMsg}`);
            } else { // 予定が登録された場合
                messages.push(`\nGoogleカレンダーに予定が登録されました。${commonMsg}`);
            }
        }
    }

    // メッセージ配列を結合してひとつにする
    message = messages.join('\n----------\n');

    console.timeEnd('----- generateNotifyMessages -----');
    return message;
}

/**
 * ファイルに保存した予定検索関数（主キー検索用）
 *
 * @param id ID
 */
function searchStoredEventById(id) {
    console.time('----- searchStoredEventById -----');

    // ファイルに保存した予定を検索する
    const event = searchStoredEvents('id = ' + id)[0];

    console.timeEnd('----- searchStoredEventById -----');
    return event;
}

/**
 * ファイルに保存した予定検索関数（複数結果取得用）
 *
 * @param filter 検索条件
 */
function searchStoredEvents(filter) {
    console.time('----- searchStoredEvents -----');

    // ファイルに保存した予定を検索する
    const events = [];
    const result = SpreadSheetsSQL.open(FILE_ID_EVENTS, 'DATA').select(['id', 'summary', 'start', 'end']).filter(filter).result();
    for (let i = 0; i < result.length; i++) {
        // 検索結果からStoredEventインスタンスを生成し、配列に格納する
        events.push(new StoredEvent(result[i].id, result[i].summary, result[i].start, result[i].end));
    }

    console.timeEnd('----- searchStoredEvents -----');
    return events;
}

/**
 * 日時->文字列変換関数
 *
 * @param source 日時
 */
function dateToString(source) {
    console.time('----- dateToString -----');

    const strSource = String(source);
    let stringFormat = '';
    const hhmm = strSource.split('T')[1];
    // 時刻が存在する場合のみ変換処理を行う
    if (hhmm) {
        const yyyyMMdd = strSource.split('T')[0];
        const yyyy = String(yyyyMMdd).split('-')[0];
        const MM = String(yyyyMMdd).split('-')[1];
        const dd = String(yyyyMMdd).split('-')[2];
        const hh = String(hhmm).split(':')[0];
        const mm = String(hhmm).split(':')[1];
        stringFormat = yyyy + '-' + MM + '-' + dd + ' ' + hh + ':' + mm;
    } else {
        // 引数をそのまま代入する
        stringFormat = source;
    }

    console.timeEnd('----- dateToString -----');
    return stringFormat;    
}

/**
 * 日付比較関数
 * 
 * @param start 開始日
 * @param end 終了日
 */
function dateComparison(start='2022-12-31', end='2023-01-01') {
    console.time('----- dateComparison -----');

    let result = '';

    const dtStart = new Date(start);
    const dtEnd = new Date(end);

    // 終了日を1日巻き戻す
    dtEnd.setDate(dtEnd.getDate() - 1);

    const dtStartYear = dtStart.getFullYear();
    const dtStartMonth = dtStart.getMonth() + 1;
    const dtStartDay = dtStart.getDate();
    const dtEndYear = dtEnd.getFullYear();
    const dtEndMonth = dtEnd.getMonth() + 1;
    const dtEndDay = dtEnd.getDate();

    if (dtStartYear !== dtEndYear || dtStartMonth !== dtEndMonth || dtStartDay !== dtEndDay) {
        result = `${dtEndYear}-${dtEndMonth.toString().padStart(2, '0')}-${dtEndDay.toString().padStart(2, '0')}`;
    }

    console.timeEnd('----- dateComparison -----');
    return result;
}

/**
 * LINE通知関数
 *
 * @param message 通知メッセージ
 */
function notifyLINE(message) {
    console.time('----- notifyLINE -----');

    // ScriptPropertiesからLINEのアクセストークンを取得する
    const token = properties.getProperty(PROPERTY_KEY_LINE_TOKEN);

    // LINE Notify APIを呼び出し、通知を行なう
    const options = {
        method: 'post',
        payload: 'message=' + message,
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
    };
    UrlFetchApp.fetch(ENDPOINT_LINE_NOTIFY_API, options);

    console.timeEnd('----- notifyLINE -----');
}

/**
 * スプレッドシートに保存した予定の全件リスト最新化関数
 *
 * @param calendarId 対象カレンダーのID
 */
function refleshStoredEvents(calendarId) {
    console.time('----- refleshStoredEvents -----');

    // 既存の全件リストを削除する
    SpreadSheetsSQL.open(FILE_ID_EVENTS, 'DATA').deleteRows();

    // カレンダーから最新の予定の全件リストを取得する
    const events = Calendar.Events.list(calendarId, { 'timeMin': (new Date()).toISOString() }).items;

    const storedEvents = [];
    for (let i = 0; i < events.length; i++) {
        // 予定のdateもしくはdateTimeから予定の開始日時／終了日時を取得する
        const start = (events[i].start.dateTime) ? events[i].start.dateTime : events[i].start.date;
        const end = (events[i].end.dateTime) ? events[i].end.dateTime : events[i].end.date;

        // 検索結果からStoredEventインスタンスを生成し、配列に格納する
        storedEvents.push(new StoredEvent(events[i].id, events[i].summary, start, end));
    }

    // 最新の予定の全件リストをスプレッドシートに保存する
    SpreadSheetsSQL.open(FILE_ID_EVENTS, 'DATA').insertRows(storedEvents);

    // スプレッドシートの全セルの書式設定を「書式なし」に設定する
    SpreadsheetApp.openById(FILE_ID_EVENTS).getSheetByName('DATA').getDataRange().setNumberFormat('@');

    console.timeEnd('----- refleshStoredEvents -----');
}

/**
 * スプレッドシートに保存した予定を表すエンティティクラス
 */
class StoredEvent {
    constructor(id, summary, start, end) {
        this.id = id;
        this.summary = summary;
        this.start = start;
        this.end = end;
    }
}