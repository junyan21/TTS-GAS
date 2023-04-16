# スプレッドシートに記載した英語テキストをGCP Text-To-Speechでmp3にするGAS

## OAuth2認証
下記ブログに助けていただいた。ありがとうございます。
- https://officeforest.org/wp/2023/01/14/google-apps-script%E3%81%8B%E3%82%89twitter-api%E3%82%92oauth2-0%E8%AA%8D%E8%A8%BC%E3%81%A7%E4%BD%BF%E3%81%86/#OAuth20

## メモ
- 変換量が増えた場合は、処理の分解を検討
- 指定したセルのみ変換する処理もあったほうが便利。直近はまるごとで乗り切る。