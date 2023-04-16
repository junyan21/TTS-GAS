// Compiled using tts-gas 1.0.0 (TypeScript 4.9.5)
"use strict";

// Thanks for a following post
// https://officeforest.org/wp/2018/10/28/google-apps-script%E3%81%A8cloud-speech-api%E3%81%A7%E6%96%87%E5%AD%97%E8%B5%B7%E3%81%93%E3%81%97/

// 認証用各種変数
const tokenUrl = "https://accounts.google.com/o/oauth2/token"
const jsonFileURL = "Your Google Cloud Service Account Key JSON File URL"

//OAuth2認証を実行する
function startOAuth() {
    //UIを取得する
    const ui = SpreadsheetApp.getUi();

    //認証を実行する
    const service = checkOAuth();
    ui.alert("認証が完了し、Access Tokenを取得しました。")
}

//Google DriveにあるサービスアカウントキーのJSONファイルを取得する
function getServiceAccKey() {
    // jsonFileURLからFileIDを取得する。ファイルがある体で!指定する
    const fileId = jsonFileURL.match(/[\w-]{25,}/)![0];

    //JSONファイルの中身を取得する
    const content = DriveApp.getFileById(fileId).getAs("application/json").getDataAsString();
    return JSON.parse(content);
}

//OAuth2.0認証を実行する
function checkOAuth() {
    //JSONファイルの中身を取得する
    var privateKeys = getServiceAccKey();

    // OAuth2はGASエディタからライブラリを追加する
    // "1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF"
    return OAuth2.createService('GoogleCloud:' + Session.getActiveUser().getEmail())
        //アクセストークンの取得用URLをセット
        .setTokenUrl(tokenUrl)

        //プライベートキーとクライアントIDをセットする
        .setPrivateKey(privateKeys['private_key'])
        .setIssuer(privateKeys['client_email'])

        //Access Tokenをスクリプトプロパティにセットする
        .setPropertyStore(PropertiesService.getScriptProperties())

        //スコープを設定する
        .setScope('https://www.googleapis.com/auth/cloud-platform');
}

//Access Tokenを取得する
function getOAuthToken() {
    //DriveAppを参照させておく
    //https://issuetracker.google.com/issues/131443766
    const id = DriveApp.getRootFolder().getId();

    //アクセストークンを取得する
    return ScriptApp.getOAuthToken();
}

const service = checkOAuth();
if (!service.hasAccess()) {
    console.log("This script does't have access!!!")
}

const accessToken = service.getAccessToken()
const TTS_API_KEY = 'Google Cloud Text-to-Speech API Key';
const SOUNDS_FOLDER_ID = 'Goodle Drive Folder ID';

function convertTextToSpeech() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    // Specify your target range
    const data = sheet.getRange("B1:B").getValues();
    const folder = Drive.Files!.get(SOUNDS_FOLDER_ID);

    for (let i = 0; i < data.length; i++) {
        const text = data[i][0];
        if (text === '') {
            continue;
        }
        const fileName = `${text.slice(0, 15)}.mp3`;
        const payload = {
            input: { text: text },
            voice: { languageCode: 'en-US', 'name': 'en-US-Neural2-A' },
            audioConfig: { audioEncoding: 'MP3' }
        };
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            method: 'post',
            contentType: 'application/json',
            headers: { 'Authorization': `Bearer ${accessToken}` },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };
        const response = UrlFetchApp.fetch('https://texttospeech.googleapis.com/v1/text:synthesize', options);
        const jsonResponse = JSON.parse(response.getContentText());
        const audioContent = Utilities.base64Decode(jsonResponse.audioContent);
        const audioBlob = Utilities.newBlob(audioContent, 'audio/mpeg', fileName);

        // Create the file using Drive API
        Drive.Files!.insert({ title: fileName, mimeType: 'audio/mpeg', parents: [{ id: SOUNDS_FOLDER_ID }] }, audioBlob);
    }
}
