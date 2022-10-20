# google-calendar-line-notification
Googleカレンダーの変更をLINEに通知するGASアプリ  
カレンダーを複数人で共有している時に便利です

## 使用サービス・ライブラリ
- LINE Notify
- Google Calendar API
- Google Sheets API
- SpreadSheetsSQL

## 仕様
1. Googleカレンダーに予定を登録・変更・削除すると、GASイベントが発火
2. GoogleスプレッドシートをDB代わりに、差分をチェック
3. 差分のある予定を、LINE Notifyで通知

## ライセンス
[MIT](LICENSE)
