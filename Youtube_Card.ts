import RPA from 'ts-rpa';
import { WebDriver, By, FileDetector, Key } from 'selenium-webdriver';
import { DH_NOT_SUITABLE_GENERATOR } from 'constants';
var fs = require('fs');

// 読み込みする スプレッドシートID と シート名 の記載
//const SSID = process.env.Youtube_Card_Setting_SheetID_test;
//const SSName1 = process.env.Youtube_Card_Setting_SheetName_test;
const SSID = process.env.Youtube_Card_Setting_SheetID;
const SSName1 = process.env.Youtube_Card_Setting_SheetName;

// 画像アップロードするGoogleDrive ID
const DriveID1 = process.env.Youtube_Card_Setting_Drive1;
const DriveID2 = process.env.Youtube_Card_Setting_Drive2;
const DriveID3 = process.env.Youtube_Card_Setting_Drive3;
const DriveID4 = process.env.Youtube_Card_Setting_Drive4;
const DriveID5 = process.env.Youtube_Card_Setting_Drive5;
const DriveID6 = process.env.Youtube_Card_Setting_Drive6;
const DriveIDs = [DriveID1, DriveID2, DriveID3, DriveID4, DriveID5, DriveID6];
// ダウンロードフォルダ
const DownloadDir = __dirname + '/Download/';

async function Start() {
  // スクリプト実行前にダウンロードフォルダを全て削除
  await RPA.File.rimraf({ dirPath: `${DownloadDir}` });
  // デバッグログを最小限にする
  RPA.Logger.level = 'INFO';
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
  });
  // 設定するアカウントデータ
  const FirstLoginAccount = ['アカウント名', 'ID', 'PW', 'URL'];
  const Account_left1 = ['アカウント名', 'ID', 'PW', 'URL'];
  const Account_left2 = ['アカウント名', 'ID', 'PW', 'URL'];
  const Account_left3 = ['アカウント名', 'ID', 'PW', 'URL'];
  const Account_right1 = ['アカウント名', 'ID', 'PW', 'URL'];
  const Account_right2 = ['アカウント名', 'ID', 'PW', 'URL'];
  const Account_right3 = ['アカウント名', 'ID', 'PW', 'URL'];
  const Accounts = [
    Account_left1,
    Account_left2,
    Account_left3,
    Account_right1,
    Account_right2,
    Account_right3
  ];
  await AccountSelect(
    FirstLoginAccount,
    Account_left1,
    Account_left2,
    Account_left3,
    Account_right1,
    Account_right2,
    Account_right3
  );
  //そもそも設定するかどうかのフラッグ・URLか再生リストかのフラッグ・何秒設定・URL・フレーズ・ティーザー・検索タイトル/カードタイトル・設定タイトル/詳細・固定コメント
  const SheetDatas = [
    [
      'false',
      ['URLor再生リスト'],
      ['秒数'],
      ['URL'],
      ['フレーズ文言'],
      ['ティーザー文言'],
      ['検索タイトル'],
      ['動画タイトル/詳細'],
      ['固定コメント'],
      ['カードタイトル']
    ],
    [
      'false',
      ['URLor再生リスト'],
      ['秒数'],
      ['URL'],
      ['フレーズ文言'],
      ['ティーザー文言'],
      ['検索タイトル'],
      ['動画タイトル/詳細'],
      ['固定コメント'],
      ['カードタイトル']
    ],
    [
      'false',
      ['URLor再生リスト'],
      ['秒数'],
      ['URL'],
      ['フレーズ文言'],
      ['ティーザー文言'],
      ['検索タイトル'],
      ['動画タイトル/詳細'],
      ['固定コメント'],
      ['カードタイトル']
    ],
    [
      'false',
      ['URLor再生リスト'],
      ['秒数'],
      ['URL'],
      ['フレーズ文言'],
      ['ティーザー文言'],
      ['検索タイトル'],
      ['動画タイトル/詳細'],
      ['固定コメント'],
      ['カードタイトル']
    ],
    [
      'false',
      ['URLor再生リスト'],
      ['秒数'],
      ['URL'],
      ['フレーズ文言'],
      ['ティーザー文言'],
      ['検索タイトル'],
      ['動画タイトル/詳細'],
      ['固定コメント'],
      ['カードタイトル']
    ],
    [
      'false',
      ['URLor再生リスト'],
      ['秒数'],
      ['URL'],
      ['フレーズ文言'],
      ['ティーザー文言'],
      ['検索タイトル'],
      ['動画タイトル/詳細'],
      ['固定コメント'],
      ['カードタイトル']
    ]
  ];
  // スプレッドシートからデータを取得する
  await ReadSheet1(SheetDatas);
  RPA.Logger.info(Accounts);
  try {
    // 格闘のアカウントで最初にログインする
    await YoutubeFirstLogin(FirstLoginAccount);
    await RPA.sleep(2000);
    for (let i in SheetDatas) {
      if (SheetDatas[i][0] == 'true') {
        if (Accounts[i][0] == '公式') {
          // NewsのアカウントはRPAではブロックされてログインできない
          continue;
        }
      }
      if (SheetDatas[i][0] == 'true') {
        await Youtube2ndLogin(Accounts[i]);
        await TitleSearch(SheetDatas[i][6][0]);
        await RPA.sleep(3000);
        await CardCreate(SheetDatas[i], i);
        await TitleSyousaiChange(SheetDatas[i]);
        if (SheetDatas[i][8][0] != undefined) {
          await CommentPost(SheetDatas[i]);
        }
      }
    }
  } catch {
    await RPA.WebBrowser.takeScreenshot();
    await RPA.Logger.info(
      'エラー出現.スクリーンショット撮ってブラウザ終了します'
    );
    await RPA.WebBrowser.quit();
  }
}

Start();

async function AccountSelect(
  FirstLoginAccount,
  Account_left1,
  Account_left2,
  Account_left3,
  Account_right1,
  Account_right2,
  Account_right3
) {
  // 設定するアカウントを取得する
  const firstAccountData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!B2:S48`
  });
  // 格闘のID・PW・URLを最初に取得する
  FirstLoginAccount[1] = firstAccountData[7][15];
  FirstLoginAccount[2] = firstAccountData[7][16];
  FirstLoginAccount[3] = firstAccountData[7][17];
  // スプシB2・B25・B48・J2・J25・J48　のプルダウンをアカウントとして保持する
  Account_left1[0] = firstAccountData[0][0];
  Account_left2[0] = firstAccountData[23][0];
  Account_left3[0] = firstAccountData[46][0];
  Account_right1[0] = firstAccountData[0][8];
  Account_right2[0] = firstAccountData[23][8];
  Account_right3[0] = firstAccountData[46][8];
  for (let i in firstAccountData) {
    // B2・B25・B48・J2・J25・J48　のプルダウンで選択されたアカウントのID・PW・URLを取得する
    if (Account_left1[0] == firstAccountData[i][14]) {
      Account_left1[1] = firstAccountData[i][15];
      Account_left1[2] = firstAccountData[i][16];
      Account_left1[3] = firstAccountData[i][17];
    }
    if (Account_left2[0] == firstAccountData[i][14]) {
      Account_left2[1] = firstAccountData[i][15];
      Account_left2[2] = firstAccountData[i][16];
      Account_left2[3] = firstAccountData[i][17];
    }
    if (Account_left3[0] == firstAccountData[i][14]) {
      Account_left3[1] = firstAccountData[i][15];
      Account_left3[2] = firstAccountData[i][16];
      Account_left3[3] = firstAccountData[i][17];
    }
    if (Account_right1[0] == firstAccountData[i][14]) {
      Account_right1[1] = firstAccountData[i][15];
      Account_right1[2] = firstAccountData[i][16];
      Account_right1[3] = firstAccountData[i][17];
    }
    if (Account_right2[0] == firstAccountData[i][14]) {
      Account_right2[1] = firstAccountData[i][15];
      Account_right2[2] = firstAccountData[i][16];
      Account_right2[3] = firstAccountData[i][17];
    }
    if (Account_right3[0] == firstAccountData[i][14]) {
      Account_right3[1] = firstAccountData[i][15];
      Account_right3[2] = firstAccountData[i][16];
      Account_right3[3] = firstAccountData[i][17];
    }
  }
  RPA.Logger.info('アカウント 左 1 ' + Account_left1);
  RPA.Logger.info('アカウント 左 2 ' + Account_left2);
  RPA.Logger.info('アカウント 左 3 ' + Account_left3);
  RPA.Logger.info('アカウント 右 1 ' + Account_right1);
  RPA.Logger.info('アカウント 右 2 ' + Account_right2);
  RPA.Logger.info('アカウント 右 3 ' + Account_right3);
}

async function YoutubeFirstLogin(Account) {
  await RPA.WebBrowser.get(Account[3]);
  await RPA.sleep(1000);
  // ヘッドレスモード 用
  const GoogleLoginID = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'Email' }),
    8000
  );
  await RPA.WebBrowser.sendKeys(GoogleLoginID, [`${Account[1]}`]);
  const NextButton = await RPA.WebBrowser.findElementById('next');
  await RPA.WebBrowser.mouseClick(NextButton);
  await RPA.sleep(1000);
  const GoogleLoginPW = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'Passwd' }),
    5000
  );
  await RPA.WebBrowser.sendKeys(GoogleLoginPW, [`${Account[2]}`]);
  const NextButton2 = await RPA.WebBrowser.findElementById('signIn');
  await RPA.WebBrowser.mouseClick(NextButton2);
  await RPA.sleep(4000);
  try {
    const ChannelSelectOK_Button = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementsLocated({
        className: `content-container yt-valign-container`
      }),
      5000
    );
    await RPA.WebBrowser.mouseClick(ChannelSelectOK_Button[0]);
    await RPA.sleep(5000);
  } catch {
    RPA.Logger.info('アカウント選択画面出ませんでしたので、スキップします');
  }
  // 画面確認用ログインver
  /*
    const GoogleLoginID = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:`//*[@id="identifierId"]`}),8000);
    await RPA.WebBrowser.sendKeys(GoogleLoginID,[`${Account[1]}`]);
    const NextButton = await RPA.WebBrowser.findElementByXPath('//*[@id="identifierNext"]');
    await RPA.WebBrowser.mouseClick(NextButton);
    await RPA.sleep(1000);
    const GoogleLoginPW = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'//*[@id="password"]/div[1]/div/div[1]/input'}),5000);
    await RPA.WebBrowser.sendKeys(GoogleLoginPW,[`${Account[2]}`]);
    const NextButton2 = await RPA.WebBrowser.findElementByXPath('//*[@id="passwordNext"]');
    await RPA.WebBrowser.mouseClick(NextButton2);
    await RPA.sleep(4000);
    try{
        const ChannelSelect = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({className:`specialized-identity-prompt-account-item page main`}),5000);
        await RPA.WebBrowser.mouseClick(ChannelSelect);
        await RPA.sleep(5000);
    }
    catch{
        RPA.Logger.info("アカウント選択画面出ませんでしたので、スキップします");
    }
    */
}

async function Youtube2ndLogin(Account) {
  // 別のアカウントに切り替えるためにログアウトする
  await RPA.WebBrowser.get('https://studio.youtube.com/channel/');
  await RPA.sleep(1500);
  const AccountButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'avatar-btn' }),
    8000
  );
  await RPA.WebBrowser.mouseClick(AccountButton);
  await RPA.sleep(500);
  const LogOutButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementsLocated({
      className: 'style-scope ytd-compact-link-renderer'
    }),
    5000
  );
  for (let i in LogOutButton) {
    const LogOutText = await LogOutButton[i].getText();
    if (LogOutText == 'ログアウト') {
      await RPA.WebBrowser.mouseClick(LogOutButton[i]);
      await RPA.sleep(3000);
      break;
    }
  }
  // 別のアカウントでログインする
  const LoginButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementsLocated({
      className: 'style-scope ytd-button-renderer style-suggestive size-small'
    }),
    5000
  );
  for (let i in LoginButton) {
    const LoginButtonText = await LoginButton[i].getText();
    if (LoginButtonText == 'ログイン') {
      await RPA.WebBrowser.mouseClick(LoginButton[i]);
      break;
    }
  }
  await RPA.sleep(1500);
  // headlessモード用
  const AccountAdd = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'account-chooser-add-account' }),
    5000
  );
  await RPA.WebBrowser.mouseClick(AccountAdd);
  await RPA.sleep(1000);
  const LoginID = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'Email' }),
    8000
  );
  await RPA.WebBrowser.sendKeys(LoginID, [Account[1]]);
  const NextButton1 = await RPA.WebBrowser.findElementById('next');
  await RPA.WebBrowser.mouseClick(NextButton1);
  await RPA.sleep(1000);
  const LoginPW = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'Passwd' }),
    5000
  );
  await RPA.WebBrowser.sendKeys(LoginPW, [Account[2]]);
  const NextButton2 = await RPA.WebBrowser.findElementById('signIn');
  await RPA.WebBrowser.mouseClick(NextButton2);
  await RPA.sleep(4000);
  try {
    const ChannelSelectOK_Button = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementsLocated({
        className: `content-container yt-valign-container`
      }),
      8000
    );
    await RPA.WebBrowser.mouseClick(ChannelSelectOK_Button[0]);
    await RPA.sleep(5000);
  } catch {
    RPA.Logger.info('アカウント選択画面出ませんでしたので、スキップします');
  }
  await RPA.WebBrowser.get(Account[3]);
  await RPA.sleep(3000);
  // アカウントを本家に切り替える
  const AvatorButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'avatar-btn' }),
    8000
  );
  await RPA.WebBrowser.mouseClick(AvatorButton);
  const AccountChangeButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementsLocated({
      className: 'style-scope ytd-compact-link-renderer'
    }),
    5000
  );
  for (let i in AccountChangeButton) {
    const LogOutText = await AccountChangeButton[i].getText();
    if (LogOutText == 'アカウントを切り替える') {
      await RPA.WebBrowser.mouseClick(AccountChangeButton[i]);
      await RPA.sleep(3000);
      break;
    }
  }
  const AccountName = await RPA.WebBrowser.findElementsByClassName(
    'style-scope ytd-account-item-renderer'
  );
  for (let i in AccountName) {
    const NameText = await AccountName[i].getText();
    if (Account[0] == 'アニメ') {
      if (NameText.indexOf('AbemaTV アニメ【公式】') >= 0) {
        await RPA.WebBrowser.mouseClick(AccountName[i]);
        RPA.Logger.info('AbemaTV アニメ【公式】 に切り替えました');
        break;
      }
    }
    if (Account[0] == '恋リア') {
      if (NameText.indexOf('AbemaTV 恋愛リアリティーショー【公式】') >= 0) {
        await RPA.WebBrowser.mouseClick(AccountName[i]);
        RPA.Logger.info(
          'AbemaTV 恋愛リアリティーショー【公式】 に切り替えました'
        );
        break;
      }
    }
    if (Account[0] == 'バラエティ') {
      if (NameText.indexOf('AbemaTV バラエティ【公式】') >= 0) {
        await RPA.WebBrowser.mouseClick(AccountName[i]);
        RPA.Logger.info('AbemaTV バラエティ【公式】に切り替えました');
        break;
      }
    }
    if (Account[0] == 'Mリーグ') {
      if (NameText.indexOf('M.LEAGUE [プロ麻雀リーグ]') >= 0) {
        await RPA.WebBrowser.mouseClick(AccountName[i]);
        RPA.Logger.info('M.LEAGUE [プロ麻雀リーグ] に切り替えました');
        break;
      }
    }
    if (Account[0] == '格闘') {
      if (NameText.indexOf('AbemaTV 格闘CH【公式】') >= 0) {
        await RPA.WebBrowser.mouseClick(AccountName[i]);
        RPA.Logger.info('AbemaTV 格闘CH【公式】 に切り替えました');
        break;
      }
    }
    if (Account[0] == 'ニュース') {
      if (NameText.indexOf('AbemaTV ニュース') >= 0) {
        await RPA.WebBrowser.mouseClick(AccountName[i]);
        RPA.Logger.info('AbemaTV ニュース に切り替えました');
        break;
      }
    }
  }
  await RPA.sleep(5000);

  // 画面確認用ログインver
  /*
    const AccountAdd = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementsLocated({className:`lCoei YZVTmd SmR8`}),5000);
    for(let i in AccountAdd){
        const Text = await AccountAdd[i].getAttribute("innerText");
        if(Text == "別のアカウントを使用"){
            await RPA.WebBrowser.mouseClick(AccountAdd[i]);
            break
        }
    }
    await RPA.sleep(1000);
    const LoginID = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'//*[@id="identifierId"]'}),8000);
    await RPA.WebBrowser.sendKeys(LoginID,[Account[1]]);
    const NextButton1 = await RPA.WebBrowser.findElementByXPath(`//*[@id="identifierNext"]`);
    await RPA.WebBrowser.mouseClick(NextButton1);
    await RPA.sleep(1000);
    const LoginPW = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'//*[@id="password"]/div[1]/div/div[1]/input'}),5000);
    await RPA.WebBrowser.sendKeys(LoginPW,[Account[2]]);
    const NextButton2 = await RPA.WebBrowser.findElementByXPath('//*[@id="passwordNext"]');
    await RPA.WebBrowser.mouseClick(NextButton2);
    await RPA.sleep(4000);
    try{
        const ChannelSelect = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({className:`specialized-identity-prompt-account-item page main`}),5000);
        await RPA.WebBrowser.mouseClick(ChannelSelect);
        await RPA.sleep(5000);
    }
    catch{
        RPA.Logger.info("アカウント選択画面出ませんでしたので、スキップします");
    }
    await RPA.WebBrowser.get(Account[3]);
    await RPA.sleep(3000);
    */
}

async function TitleSearch(TargetTitle) {
  await RPA.sleep(2000);
  const SearchInput = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'query-input' }),
    9000
  );
  await RPA.sleep(1000);
  await RPA.WebBrowser.sendKeys(SearchInput, [TargetTitle]);
  await RPA.sleep(5000);
  const Titles = await RPA.WebBrowser.findElementsByClassName(
    'style-scope ytcp-video-list-cell-video remove-default-style'
  );
  const TitlesImage = await RPA.WebBrowser.findElementsByClassName(
    'style-scope ytcp-img-with-fallback'
  );
  const TitleSearchJudge = ['false'];
  for (let i in Titles) {
    const TitleText = await Titles[i].getText();
    if (TitleText == TargetTitle) {
      RPA.Logger.info('検索にてタイトル一致しました');
      RPA.Logger.info(TitleText);
      await RPA.WebBrowser.mouseClick(TitlesImage[i]);
      await RPA.sleep(3000);
      TitleSearchJudge[0] = 'true';
      break;
    }
  }
  if (TitleSearchJudge[0] == 'false') {
    RPA.Logger.info('検索にてタイトルが一致しませんでした');
    await RPA.WebBrowser.quit();
  }
  await RPA.sleep(1000);
  const CardButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementsLocated({
      className: 'dropdown-trigger-text style-scope ytcp-text-dropdown-trigger'
    }),
    5000
  );
  for (let i in CardButton) {
    const CardButtonText = await CardButton[i].getText();
    if (CardButtonText.includes('カード') == true) {
      await RPA.WebBrowser.mouseClick(CardButton[i]);
      await RPA.sleep(2000);
      break;
    }
  }
  // 新しいタブができたらそちらに移る
  const windows = await RPA.WebBrowser.getAllWindowHandles();
  RPA.Logger.info(windows.length);
  if (windows.length != 1) {
    await RPA.WebBrowser.switchToWindow(windows[1]);
  }
}

async function ReadSheet1(SheetDatas) {
  const Data1 = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!B3:F23`
  });
  const Data2 = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!B26:F46`
  });
  const Data3 = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!B49:F69`
  });
  const Data4 = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!J3:N23`
  });
  const Data5 = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!J26:N46`
  });
  const Data6 = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!J49:N69`
  });
  const AllData = [];
  AllData[0] = Data1;
  AllData[1] = Data2;
  AllData[2] = Data3;
  AllData[3] = Data4;
  AllData[4] = Data5;
  AllData[5] = Data6;
  for (let i in AllData) {
    if (AllData[i][0][0] == 'RPA実行OK') {
      SheetDatas[i][0] = 'true';
      try {
        // 秒数が入っているブロックだけデータ取得する
        if (AllData[i][2][3].length > 0) {
          // リンクか再生リスト
          SheetDatas[i][1][0] = AllData[i][1][0];
          // 秒数
          //SheetDatas[i][2][0] = AllData[i][2][3];
          // URL
          SheetDatas[i][3][0] = AllData[i][2][2];
          // フレーズ文言
          SheetDatas[i][4][0] = AllData[i][3][1];
          // ティーザー文言
          SheetDatas[i][5][0] = AllData[i][4][1];
          // 検索するタイトル
          SheetDatas[i][6][0] = AllData[i][2][4];
          // カードのタイトル
          SheetDatas[i][9][0] = AllData[i][2][1];
          // カードのタイトルが 50文字以上なら、短くする
          if (String(SheetDatas[i][9][0]).length > 50) {
            RPA.Logger.info('タイトル文字数オーバー');
            const splitdata = SheetDatas[i][9][0].split('|');
            SheetDatas[i][9][0] = splitdata[0];
          }
          // 動画タイトル
          SheetDatas[i][7][0] = AllData[i][4][4];
          // 動画詳細
          SheetDatas[i][7][1] = AllData[i][6][4];
          // 固定コメント
          SheetDatas[i][8][0] = AllData[i][8][4];
        }
        if (AllData[i][6][3].length > 0) {
          SheetDatas[i][1][1] = AllData[i][5][0];
          SheetDatas[i][2][0] = AllData[i][6][3];
          SheetDatas[i][3][1] = AllData[i][6][2];
          SheetDatas[i][4][1] = AllData[i][7][1];
          SheetDatas[i][5][1] = AllData[i][8][1];
          SheetDatas[i][9][1] = AllData[i][6][1];
        }
        if (AllData[i][10][3].length > 0) {
          SheetDatas[i][1][2] = AllData[i][9][0];
          SheetDatas[i][2][1] = AllData[i][10][3];
          SheetDatas[i][3][2] = AllData[i][10][2];
          SheetDatas[i][4][2] = AllData[i][11][1];
          SheetDatas[i][5][2] = AllData[i][12][1];
          SheetDatas[i][9][2] = AllData[i][10][1];
        }
        if (AllData[i][14][3].length > 0) {
          SheetDatas[i][1][3] = AllData[i][13][0];
          SheetDatas[i][2][2] = AllData[i][14][3];
          SheetDatas[i][3][3] = AllData[i][14][2];
          SheetDatas[i][4][3] = AllData[i][15][1];
          SheetDatas[i][5][3] = AllData[i][16][1];
          SheetDatas[i][9][3] = AllData[i][14][1];
        }
        if (AllData[i][18][3].length > 0) {
          SheetDatas[i][1][4] = AllData[i][17][0];
          SheetDatas[i][2][3] = AllData[i][18][3];
          SheetDatas[i][3][4] = AllData[i][18][2];
          SheetDatas[i][4][4] = AllData[i][19][1];
          SheetDatas[i][5][4] = AllData[i][20][1];
          SheetDatas[i][9][4] = AllData[i][18][1];
        }
      } catch {}
    }
  }
  RPA.Logger.info(SheetDatas);
}

async function CardCreate(CardData, let_i) {
  await RPA.sleep(200);
  const MoviePlayer = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ id: 'movie_player' }),
    5000
  );
  const MoviePlayButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className: 'ytp-play-button ytp-button'
    }),
    3000
  );
  await RPA.WebBrowser.mouseMove(MoviePlayButton);
  // 開始3秒で一度止めてカードを設定する
  while (0 == 0) {
    const TimeSecond_3 = await RPA.WebBrowser.findElementByClassName(
      'ytp-time-current'
    );
    const TimeText = await TimeSecond_3.getText();
    if (String(TimeText) == '0:03') {
      await RPA.WebBrowser.mouseClick(MoviePlayer);
      break;
    }
    await RPA.sleep(200);
  }
  RPA.Logger.info(CardData);
  // カードを追加するボタンをおす
  const CardAppendButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      xpath:
        '/html/body/div[2]/div[4]/div/div[5]/div/div/div[1]/div[2]/div[2]/div[1]/span/button'
    }),
    5000
  );
  const CardAppendButtonText = await CardAppendButton.getText();
  if (String(CardAppendButtonText) == 'カードを追加') {
    RPA.Logger.info('カードを追加 ボタン押しました');
    await RPA.WebBrowser.mouseClick(CardAppendButton);
    await RPA.sleep(500);
  }
  // リンク作成 ボタンをおす
  const Texts = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementsLocated({
      className: 'annotator-clickcard-title'
    }),
    5000
  );
  const SakuseiButton = await RPA.WebBrowser.findElementByCSSSelector(
    `#yt-uix-clickcard-card4 > div.yt-uix-clickcard-card-border > div.yt-uix-clickcard-card-body > div > div:nth-child(4) > div.annotator-clickcard-right > button.yt-uix-button.yt-uix-button-size-default.yt-uix-button-default.annotator-create-button.yt-uix-clickcard-close`
  );
  for (let i in Texts) {
    const CardCreateButtonText = await Texts[i].getText();
    if (CardCreateButtonText == 'リンク') {
      await RPA.WebBrowser.mouseClick(SakuseiButton);
      RPA.Logger.info('リンクの作成 ボタン押しました');
      await RPA.sleep(300);
      break;
    }
  }
  // リンクのURL　の文字が出るまで待機
  while (0 == 0) {
    try {
      const LinkElement = await RPA.WebBrowser.findElementByClassName(
        'annotator-overlay-input-label'
      );
      const LinkURLText = await LinkElement.getText();
      if (LinkURLText == 'リンクの URL') {
        break;
      }
    } catch {}
  }
  await RPA.sleep(1000);
  // リンクURLに入力する
  const LinkURLInput = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className:
        'yt-uix-form-input-text annotator-prs-url annotator-initial-focus annotator-overlay-text-input'
    }),
    5000
  );
  await RPA.WebBrowser.sendKeys(LinkURLInput, [CardData[3][0]]);
  await RPA.sleep(300);
  const CardNextButton = await RPA.WebBrowser.findElementByClassName(
    'yt-uix-button yt-uix-button-size-default yt-uix-button-primary annotator-next'
  );
  await RPA.WebBrowser.mouseClick(CardNextButton);
  // 関連ウェブサイト　の文字が出てくるまで待機
  while (0 == 0) {
    await RPA.sleep(500);
    try {
      const Title = await RPA.WebBrowser.findElementByClassName(
        'yt annotator-prs-card-details create-only'
      );
      const TitleText = await Title.getText();
      if (TitleText == '関連ウェブサイト') {
        RPA.Logger.info('関連ウェブサイトの文字ありました');
        break;
      }
    } catch {}
  }
  // スプレッドシートにカードのタイトル記載があれば、入力する
  const CardTitleInput = await RPA.WebBrowser.findElementById(
    'annotator-prs-title-associated'
  );
  try {
    if (String(CardData[9][0]).length > 1) {
      RPA.Logger.info('タイトル記載されています');
      await CardTitleInput.clear();
      await RPA.sleep(200);
      await RPA.WebBrowser.sendKeys(CardTitleInput, [CardData[9][0]]);
      await RPA.sleep(300);
    }
    // スプレッドシートにカードタイトルの記載がなければ自動で入力された文字数を判定する
    if (String(CardData[9][0]).length < 1) {
      RPA.Logger.info('タイトル記載されていません');
      const Text1 = await RPA.WebBrowser.findElementById(
        'annotator-prs-title-associated'
      );
      const Text = await Text1.getAttribute('value');
      await RPA.sleep(1000);
      RPA.Logger.info(Text);
      if (String(Text).length > 50) {
        RPA.Logger.info('カードタイトル50文字超えたので調整します');
        const SplitText = await Text.split('|');
        await CardTitleInput.clear();
        await RPA.sleep(200);
        RPA.Logger.info(SplitText);
        await RPA.WebBrowser.sendKeys(CardTitleInput, [SplitText[0]]);
      }
    }
  } catch {
    RPA.Logger.info('3秒のカードタイトルがおかしいです');
  }
  // フレーズを入力する
  const PhraseInput = await RPA.WebBrowser.findElementById(
    'annotator-prs-action-associated'
  );
  if (CardData[4][0] != undefined) {
    RPA.Logger.info('3秒のフレーズ入力します');
    await PhraseInput.clear();
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(PhraseInput, [CardData[4][0]]);
  }
  // ティーザーを入力する
  const TeaserInput = await RPA.WebBrowser.findElementById(
    'annotator-teaser-text-associated'
  );
  if (CardData[5][0] != undefined) {
    RPA.Logger.info('3秒のティーザー入力します');
    await TeaserInput.clear();
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(TeaserInput, [CardData[5][0]]);
  }
  // カードを作成のボタンをおす
  const CardCreateApplyButton = await RPA.WebBrowser.findElementsByClassName(
    'yt-uix-button yt-uix-button-size-default yt-uix-button-primary annotator-save create-only'
  );
  for (let i in CardCreateApplyButton) {
    RPA.Logger.info(await CardCreateApplyButton[i].getText());
    if (String(await CardCreateApplyButton[i].getText()) == 'カードを作成') {
      await RPA.WebBrowser.mouseClick(CardCreateApplyButton[i]);
      await RPA.sleep(5000);
    }
  }
  // 画像が設定されていない時は、画像設定を行う
  try {
    const ImageErrorAlert = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        className: 'yt-uix-form-error-message'
      }),
      3000
    );
    const AlertText = await ImageErrorAlert.getText();
    RPA.Logger.info(AlertText);
    if (AlertText == '画像ファイルを選択してください') {
      RPA.Logger.info('画像が設定されていないため、画像設定します');
      const ImageInput = await RPA.WebBrowser.findElementByCSSSelector(
        '#annotations-initial-link-overlay-content > div > form > div.annotator-overlay-content > div.annotator-card-details-container.clearfix > div.annotator-card-image-container.annotator-prs-card-details > label > span > input'
      );
      const FileIDList = await RPA.Google.Drive.listFiles({
        parents: [`${DriveIDs[let_i]}`]
      });
      await RPA.Logger.info(`Drive内のファイル個数 → ${0 + FileIDList.length}`);
      if (FileIDList.length == 1) {
        await RPA.Google.Drive.download({ fileId: `${FileIDList[0].id}` });
        const FileList = await RPA.File.listFiles();
        const FileName = [];
        for (let i in FileList) {
          if (FileList[i].indexOf('.jpg') > 0) {
            FileName[0] = FileList[i];
            RPA.Logger.info(`この画像を設定します → ${FileName[0]}`);
            break;
          }
          if (FileList[i].indexOf('.png') > 0) {
            FileName[0] = FileList[i];
            RPA.Logger.info(`この画像を設定します → ${FileName[0]}`);
            break;
          }
        }
        RPA.Logger.info(DownloadDir + FileName[0]);
        await RPA.WebBrowser.sendKeys(ImageInput, [
          `${DownloadDir}${FileName[0]}`
        ]);
      }
      // GoogleDrive のファイルが2つ以上だったりダウンロードできなければデフォルトのあべま君画像を使う
      if (FileIDList.length != 1) {
        await RPA.Google.Drive.download({
          fileId: '1XN31XIdpb8HX9YrbZqErIEYl3GOKw9gY'
        });
        const FileList = await RPA.File.listFiles();
        const FileName = [];
        for (let i in FileList) {
          if (FileList[i].indexOf('.jpg') > 0) {
            FileName[0] = FileList[i];
            RPA.Logger.info(`この画像を設定します → ${FileName[0]}`);
            break;
          }
          if (FileList[i].indexOf('.png') > 0) {
            FileName[0] = FileList[i];
            RPA.Logger.info(`この画像を設定します → ${FileName[0]}`);
            break;
          }
        }
        RPA.Logger.info(DownloadDir + FileName[0]);
        await RPA.WebBrowser.sendKeys(ImageInput, [
          `${DownloadDir}${FileName[0]}`
        ]);
        await RPA.Logger.info('デフォルトの画像を設定します');
      }
      await RPA.sleep(2000);
      const ApplyButton = await RPA.WebBrowser.findElementByCSSSelector(
        '#annotations-initial-link-overlay-content > div > form > div.yt-uix-overlay-actions > button.yt-uix-button.yt-uix-button-size-default.yt-uix-button-primary.annotator-save.create-only'
      );
      await RPA.WebBrowser.mouseClick(ApplyButton);
      await RPA.sleep(5000);
    }
  } catch {}
  await RPA.WebBrowser.sendKeys(MoviePlayer, [RPA.WebBrowser.Key.chord('j')]);

  // リンク設定が2つ以上あるか判定する
  if (CardData[2].length == 1) {
    RPA.Logger.info('3秒以外ないのでリターン');
    return;
  }
  // 2つ目以降は、指定した秒数で設定する
  const OldTime = [];
  OldTime[0] = 0;
  for (let i in CardData[2]) {
    const Timedata = String(CardData[2][i]).split(':');
    RPA.Logger.info('時間の分割個数を調べます');
    RPA.Logger.info(Timedata.length);
    if (Timedata.length == 2) {
      const minutedata = Number(Timedata[0]) * 6;
      const seconddata = Number(Timedata[1]) / 10;
      RPA.Logger.info(minutedata + seconddata + '0秒に設定');
      const SettingTime = minutedata + seconddata - OldTime[0];
      for (let v = 0; v < SettingTime; v++) {
        await RPA.WebBrowser.sendKeys(MoviePlayer, [
          RPA.WebBrowser.Key.chord('l')
        ]);
        await RPA.sleep(150);
      }
      await CardCreates(CardData, Number(i) + 1, let_i);
      await RPA.sleep(5000);
      OldTime[0] = minutedata + seconddata;
      continue;
    }
    if (Timedata.length == 3) {
      const hourdata = Number(Timedata[0]) * 360;
      const minutedata = Number(Timedata[1]) * 6;
      const seconddata = Number(Timedata[2]) / 10;
      RPA.Logger.info(hourdata + minutedata + seconddata + '0秒に設定');
      const SettingTime = hourdata + minutedata + seconddata - OldTime[0];
      for (let v = 0; v < SettingTime; v++) {
        await RPA.WebBrowser.sendKeys(MoviePlayer, [
          RPA.WebBrowser.Key.chord('l')
        ]);
        await RPA.sleep(150);
      }
      await CardCreates(CardData, Number(i) + 1, let_i);
      await RPA.sleep(5000);
      OldTime[0] = hourdata + minutedata + seconddata;
    }
  }
}

async function CardCreates(CardData, i, let_i) {
  // カードを追加するボタンをおす
  const CardAppendButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      xpath:
        '/html/body/div[2]/div[4]/div/div[5]/div/div/div[1]/div[2]/div[2]/div[1]/span/button'
    }),
    5000
  );
  const CardAppendButtonText = await CardAppendButton.getText();
  if (String(CardAppendButtonText) == 'カードを追加') {
    await RPA.WebBrowser.mouseClick(CardAppendButton);
    await RPA.sleep(300);
  }
  const Link_or_ListFlag = [];
  Link_or_ListFlag[0] = 'None';
  // リンク作成か再生リスト作成か判定して処理を変える
  if (CardData[1][i] == 'リンク') {
    Link_or_ListFlag[0] = 'リンク';
  }
  if (CardData[1][i] == '再生リスト') {
    Link_or_ListFlag[0] = '再生リスト';
  }
  const Texts = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementsLocated({
      className: 'annotator-clickcard-title'
    }),
    5000
  );
  // フラグが再生リストなら、再生リスト作成を行う
  if (Link_or_ListFlag[0] == '再生リスト') {
    const CardCreateButton = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementsLocated({
        css: `#yt-uix-clickcard-card4 > div.yt-uix-clickcard-card-border > div.yt-uix-clickcard-card-body > div > div:nth-child(1) > div.annotator-clickcard-right > button.yt-uix-button.yt-uix-button-size-default.yt-uix-button-default.annotator-create-button.yt-uix-clickcard-close`
      }),
      5000
    );
    const CardCreateButtonText = await CardCreateButton[0].getText();
    if (CardCreateButtonText == '作成') {
      await RPA.WebBrowser.mouseClick(CardCreateButton[0]);
      RPA.Logger.info('再生リストの作成ボタン押しました');
      await RPA.sleep(300);
    }
    const PlayListInput = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        className: 'yt-uix-form-input-text yt-video-picker-url'
      }),
      5000
    );
    await RPA.WebBrowser.sendKeys(PlayListInput, [CardData[3][i]]);
    const TextCreateButton = await RPA.WebBrowser.findElementByXPath(
      '//*[@id="annotations-edit-overlay-content-video"]/div/form/div[1]/details/summary'
    );
    await RPA.WebBrowser.mouseClick(TextCreateButton);
    const ListPhraseInput = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        id: 'annotator-custom-text-video'
      }),
      5000
    );
    await RPA.WebBrowser.sendKeys(ListPhraseInput, [CardData[4][i]]);
    const ListTeaserInput = await RPA.WebBrowser.findElementById(
      'annotator-teaser-text-video'
    );
    await RPA.WebBrowser.sendKeys(ListTeaserInput, [CardData[5][i]]);
    await RPA.sleep(300);
    const CardCreateApplyButton = await RPA.WebBrowser.findElementsByClassName(
      'yt-uix-button yt-uix-button-size-default yt-uix-button-primary annotator-save create-only'
    );
    for (let i in CardCreateApplyButton) {
      const Text = await CardCreateApplyButton[i].getText();
      if (String(Text) == 'カードを作成') {
        // カード作成のボタンをおす
        await RPA.WebBrowser.mouseClick(CardCreateApplyButton[i]);
        await RPA.sleep(3000);
      }
    }
  }
  // フラッグがリンクなら、リンク作成を行う
  if (Link_or_ListFlag[0] == 'リンク') {
    const CardCreateButton = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        css: `#yt-uix-clickcard-card4 > div.yt-uix-clickcard-card-border > div.yt-uix-clickcard-card-body > div > div:nth-child(4) > div.annotator-clickcard-right > button.yt-uix-button.yt-uix-button-size-default.yt-uix-button-default.annotator-create-button.yt-uix-clickcard-close`
      }),
      5000
    );
    const CardCreateButtonText = await CardCreateButton.getText();
    if (CardCreateButtonText == '作成') {
      await RPA.WebBrowser.mouseClick(CardCreateButton);
      await RPA.sleep(300);
    }
    // リンクのURL　の文字が出るまで待機
    while (0 == 0) {
      try {
        const LinkElement = await RPA.WebBrowser.findElementByClassName(
          'annotator-overlay-input-label'
        );
        const LinkURLText = await LinkElement.getText();
        if (LinkURLText == 'リンクの URL') {
          RPA.Logger.info('リンクのURLの文字出ました');
          break;
        }
      } catch {}
    }
    await RPA.sleep(2000);
    // リンクURLに入力する
    //const LinkURLInput = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({id:'annotator-prs-url-link'}),8000);
    //await RPA.WebBrowser.sendKeys(LinkURLInput,[CardData[3][i]]);
    await RPA.WebBrowser.driver.executeScript(
      `document.querySelectorAll("#annotator-prs-url-link")[0].value = "${CardData[3][i]}"`
    );
    RPA.Logger.info(`リンクのURL入力しました → ${CardData[3][i]}`);
    await RPA.sleep(300);
    const CardNextButton = await RPA.WebBrowser.findElementByClassName(
      'yt-uix-button yt-uix-button-size-default yt-uix-button-primary annotator-next'
    );
    await RPA.WebBrowser.mouseClick(CardNextButton);
    // 関連ウェブサイト　の文字が出てくるまで待機
    while (0 == 0) {
      await RPA.sleep(500);
      try {
        const Title = await RPA.WebBrowser.findElementByClassName(
          'yt annotator-prs-card-details create-only'
        );
        const TitleText = await Title.getText();
        if (TitleText == '関連ウェブサイト') {
          RPA.Logger.info('関連ウェブサイトの文字ありました');
          break;
        }
      } catch {}
    }
    await RPA.sleep(500);
    // スプレッドシートにカードのタイトル記載があれば、入力する
    const CardTitleInput = await RPA.WebBrowser.findElementById(
      'annotator-prs-title-associated'
    );
    try {
      if (String(CardData[9][i]).length > 1) {
        RPA.Logger.info('タイトル記載されています');
        await CardTitleInput.clear();
        await RPA.sleep(200);
        await RPA.WebBrowser.sendKeys(CardTitleInput, [CardData[9][i]]);
        await RPA.sleep(300);
      }
      // スプレッドシートにカードタイトルの記載がなければ自動で入力された文字数を判定する
      if (String(CardData[9][i]).length < 1) {
        RPA.Logger.info('タイトル記載されていません');
        const Text1 = await RPA.WebBrowser.findElementById(
          'annotator-prs-title-associated'
        );
        const Text = await Text1.getAttribute('value');
        await RPA.sleep(1000);
        RPA.Logger.info(Text);
        if (String(Text).length > 50) {
          RPA.Logger.info('カードタイトル50文字超えたので調整します');
          const SplitText = await Text.split('|');
          await CardTitleInput.clear();
          await RPA.sleep(200);
          RPA.Logger.info(SplitText);
          await RPA.WebBrowser.sendKeys(CardTitleInput, [SplitText[0]]);
        }
      }
    } catch {
      RPA.Logger.info('カードタイトルがおかしいです');
    }
    // フレーズを入力する
    const PhraseInput = await RPA.WebBrowser.findElementById(
      'annotator-prs-action-associated'
    );
    if (String(CardData[4][i]).length > 1) {
      await PhraseInput.clear();
      await RPA.sleep(100);
      await RPA.WebBrowser.sendKeys(PhraseInput, [CardData[4][i]]);
    }
    // ティーザーを入力する
    const TeaserInput = await RPA.WebBrowser.findElementById(
      'annotator-teaser-text-associated'
    );
    if (String(CardData[5][i]).length > 1) {
      await TeaserInput.clear();
      await RPA.sleep(100);
      await RPA.WebBrowser.sendKeys(TeaserInput, [CardData[5][i]]);
    }
    // カードを作成のボタンをおす
    const CardCreateApplyButton = await RPA.WebBrowser.findElementsByClassName(
      'yt-uix-button yt-uix-button-size-default yt-uix-button-primary annotator-save create-only'
    );
    for (let v in CardCreateApplyButton) {
      RPA.Logger.info(await CardCreateApplyButton[v].getText());
      if (String(await CardCreateApplyButton[v].getText()) == 'カードを作成') {
        await RPA.WebBrowser.mouseClick(CardCreateApplyButton[v]);
        await RPA.sleep(4000);
      }
    }
    // 画像が設定されていない時は、画像設定を行う
    try {
      const ImageErrorAlert = await RPA.WebBrowser.wait(
        RPA.WebBrowser.Until.elementLocated({
          className: 'yt-uix-form-error-message'
        }),
        3000
      );
      const AlertText = await ImageErrorAlert.getText();
      RPA.Logger.info(AlertText);
      if (AlertText == '画像ファイルを選択してください') {
        RPA.Logger.info('画像が設定されていないため、画像設定します');
        const ImageInput = await RPA.WebBrowser.findElementByCSSSelector(
          '#annotations-initial-link-overlay-content > div > form > div.annotator-overlay-content > div.annotator-card-details-container.clearfix > div.annotator-card-image-container.annotator-prs-card-details > label > span > input'
        );
        const FileIDList = await RPA.Google.Drive.listFiles({
          parents: [`${DriveIDs[let_i]}`]
        });
        RPA.Logger.info(`Drive内のファイル個数 → ${0 + FileIDList.length}`);
        if (FileIDList.length == 1) {
          await RPA.Google.Drive.download({ fileId: `${FileIDList[0].id}` });
          const FileList = await RPA.File.listFiles();
          const FileName = [];
          for (let i in FileList) {
            if (FileList[i].indexOf('.jpg') > 0) {
              FileName[0] = FileList[i];
              RPA.Logger.info(`この画像を設定します → ${FileName[0]}`);
              break;
            }
            if (FileList[i].indexOf('.png') > 0) {
              FileName[0] = FileList[i];
              RPA.Logger.info(`この画像を設定します → ${FileName[0]}`);
              break;
            }
          }
          await RPA.WebBrowser.sendKeys(ImageInput, [
            `${DownloadDir}${FileName[0]}`
          ]);
        }
        // GoogleDrive のファイルが2つ以上だったりダウンロードできなければデフォルトのあべま君画像を使う
        if (FileIDList.length != 1) {
          await RPA.Google.Drive.download({
            fileId: '1XN31XIdpb8HX9YrbZqErIEYl3GOKw9gY'
          });
          const FileList = await RPA.File.listFiles();
          const FileName = [];
          for (let i in FileList) {
            if (FileList[i].indexOf('.jpg') > 0) {
              FileName[0] = FileList[i];
              RPA.Logger.info(`この画像を設定します → ${FileName[0]}`);
              break;
            }
            if (FileList[i].indexOf('.png') > 0) {
              FileName[0] = FileList[i];
              RPA.Logger.info(`この画像を設定します → ${FileName[0]}`);
              break;
            }
          }
          RPA.Logger.info(DownloadDir + FileName[0]);
          await RPA.WebBrowser.sendKeys(ImageInput, [
            `${DownloadDir}${FileName[0]}`
          ]);
          await RPA.Logger.info('デフォルトの画像を設定します');
        }
        await RPA.sleep(3000);
        const ApplyButton = await RPA.WebBrowser.findElementByCSSSelector(
          '#annotations-initial-link-overlay-content > div > form > div.yt-uix-overlay-actions > button.yt-uix-button.yt-uix-button-size-default.yt-uix-button-primary.annotator-save.create-only'
        );
        await RPA.WebBrowser.mouseClick(ApplyButton);
        await RPA.sleep(5000);
      }
    } catch {}
  }
}

// 動画のタイトルと詳細を記入する
async function TitleSyousaiChange(CardData) {
  // タブが２つ以上あれば1以外消す
  const TabClose = await RPA.WebBrowser.getAllWindowHandles();
  RPA.Logger.info(TabClose.length);
  if (TabClose.length > 1) {
    RPA.Logger.info('タブ閉じました');
    await RPA.WebBrowser.closeWindow(TabClose[1]);
    await RPA.WebBrowser.switchToWindow(TabClose[0]);
    await RPA.sleep(500);
  }
  //const titleinput = await RPA.WebBrowser.wait(RPA.WebBrowser.Until.elementLocated({xpath:'/html/body/ytcp-app/ytcp-entity-page/div/div/main/div/ytcp-animatable[8]/ytcp-video-metadata-editor-section/ytcp-video-metadata-editor/div/ytcp-animatable/ytcp-video-metadata-basics/div/div[1]/div[1]/ytcp-mention-textbox/ytcp-form-input-container/div[1]/div[2]/ytcp-mention-input/div'}),8000);
  const input = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementsLocated({ css: '#textbox' }),
    5000
  );
  // 動画タイトルを一度エンコードして jsで入力する
  if (CardData[7][0] != undefined) {
    await input[0].clear();
    await RPA.sleep(300);
    const Text1 = await encodeURI(CardData[7][0]);
    await RPA.WebBrowser.driver.executeScript(
      `document.querySelectorAll("#textbox")[0].innerText = decodeURI("${Text1}")`
    );
    await RPA.WebBrowser.sendKeys(input[0], [` `]);
    await RPA.Logger.info('タイトル入力したよ');
  }
  // 詳細を一度エンコードして jsで入力する
  if (CardData[7][1] != undefined) {
    await input[1].clear();
    await RPA.sleep(300);
    const Text2 = await encodeURI(CardData[7][1]);
    await RPA.WebBrowser.driver.executeScript(
      `document.querySelectorAll("#textbox")[1].innerText = decodeURI("${Text2}")`
    );
    await RPA.WebBrowser.sendKeys(input[1], [` `]);
    await RPA.Logger.info('詳細入力したよ');
  }
  //const SaveButton = await RPA.WebBrowser.findElementById('save');
  const SaveButton = await RPA.WebBrowser.findElementsByClassName(
    'label style-scope ytcp-button'
  );
  for (let i in SaveButton) {
    const SaveButtonText = await SaveButton[i].getText();
    if (SaveButtonText == '保存') {
      await RPA.Logger.info('保存ボタン 押します');
      await RPA.WebBrowser.takeScreenshot();
      // 保存ボタンをおす
      await RPA.WebBrowser.mouseClick(SaveButton[i]);
    }
  }
  await RPA.sleep(2300);
}

// 公式Youtubeページに遷移し、固定用コメントを投稿する
async function CommentPost(CardData) {
  //await RPA.WebBrowser.get("https://studio.youtube.com/video/jEmlZ0qZaqI/edit?utm_campaign=upgrade&utm_medium=redirect&utm_source=%2Fmy_videos");
  //await RPA.sleep(2000);
  const YoutubeLink = await RPA.WebBrowser.findElementById(
    'overlay-link-to-youtube'
  );
  const YoutubeURL = await YoutubeLink.getAttribute('href');
  await RPA.WebBrowser.driver.executeScript(
    "window.open(arguments[0], 'newtab')"
  );
  await RPA.sleep(500);
  // 新しいタブを開き、そこでコメント投稿する
  const Tabs = await RPA.WebBrowser.getAllWindowHandles();
  if (Tabs.length > 1) {
    await RPA.WebBrowser.switchToWindow(Tabs[1]);
    await RPA.sleep(200);
  }
  await RPA.WebBrowser.get(YoutubeURL);
  await RPA.sleep(2000);
  await RPA.WebBrowser.scrollTo({
    xpath: `//*[@id="description"]/yt-formatted-string`
  });
  await RPA.sleep(1000);
  const CommentInputButton = await RPA.WebBrowser.findElementById(
    'placeholder-area'
  );
  await RPA.WebBrowser.mouseClick(CommentInputButton);
  await RPA.sleep(500);
  const CommentInput = await RPA.WebBrowser.findElementById(
    'contenteditable-root'
  );
  // コメントを一度エンコードして jsで入力する
  const Text = await encodeURI(CardData[8][0]);
  await RPA.WebBrowser.driver.executeScript(
    `document.getElementById("contenteditable-root").innerText = decodeURI("${Text}")`
  );
  await RPA.sleep(100);
  // コメントの最後に半角スペースを入れると投稿できる
  await RPA.WebBrowser.sendKeys(CommentInput, [' ']);
  await RPA.sleep(200);
  const CommentButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementsLocated({
      className: 'style-scope ytd-button-renderer style-primary size-default'
    }),
    5000
  );
  for (let i in CommentButton) {
    const CommentButtonText = await CommentButton[i].getText();
    if (CommentButtonText == 'コメント') {
      await RPA.WebBrowser.mouseClick(CommentButton[i]);
      break;
    }
  }
  await RPA.sleep(2000);
  const CommentsCount = await RPA.WebBrowser.findElementsByCSSSelector(
    '#action-buttons'
  );
  RPA.Logger.info('コメント数 → ' + CommentsCount.length);
  // 一番上のコメントを固定する
  const Comment1 = await RPA.WebBrowser.findElementByXPath(
    '/html/body/ytd-app/div/ytd-page-manager/ytd-watch-flexy/div[4]/div[1]/div/ytd-comments/ytd-item-section-renderer/div[3]/ytd-comment-thread-renderer/ytd-comment-renderer/div[2]/div[3]/ytd-menu-renderer/yt-icon-button/button'
  );
  await RPA.WebBrowser.mouseClick(Comment1);
  const KoteiButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementsLocated({
      className: 'style-scope ytd-menu-navigation-item-renderer'
    }),
    5000
  );
  for (let i in KoteiButton) {
    const KoteiText = await KoteiButton[i].getText();
    if (KoteiText == '固定') {
      await RPA.WebBrowser.mouseClick(KoteiButton[i]);
      break;
    }
  }
  // 固定するボタンをおす
  const KoteiApplyButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementsLocated({
      className: 'style-scope yt-button-renderer style-primary size-default'
    }),
    5000
  );
  await RPA.WebBrowser.mouseClick(KoteiApplyButton[0]);
  await RPA.WebBrowser.takeScreenshot();
  await RPA.sleep(2000);
  // タブが２つ以上あれば1以外消す
  const TabClose = await RPA.WebBrowser.getAllWindowHandles();
  RPA.Logger.info(TabClose.length);
  if (TabClose.length > 1) {
    await RPA.WebBrowser.closeWindow(TabClose[1]);
    await RPA.WebBrowser.switchToWindow(TabClose[0]);
    await RPA.sleep(500);
  }
}
