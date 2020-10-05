import RPA from 'ts-rpa';
const moment = require('moment');

var Youtube = require('youtube-node');
var youtube = new Youtube();
youtube.setKey(process.env.YoutubeNode_API_ABEMA);

// 流用時の変更箇所
// バラエティ のコピー
const ChannelID = 'UC9OvLVXb-okaVtoK9V8Biwg'; // 取得したいチャンネルIDを入れる
let SheetID = process.env.RPA_Senden_SheetID; //RPA管理シートID
const SheetName = 'バラエティ';
//

var keyword = ''; // キーワードは無しにするとそのチャンネル全て取得できる
let pagetoken = ''; //50件以上検索するためのページトークン

const SheetDay = moment().format('D'); // スプシに貼り付ける日付を取得
let YoutubeDay = moment().subtract(2, 'days').format('YYYY-MM-DD'); // 昨日を取得
console.log(SheetDay, '日');
const TodayDate = moment().subtract(1, 'days').format('YYYY/MM/DD'); // 昨日の日付を取得

// 貼り付ける列
const SheetColmun = {
  '1': 'S',
  '2': 'T',
  '3': 'U',
  '4': 'V',
  '5': 'W',
  '6': 'X',
  '7': 'Y',
  '8': 'Z',
  '9': 'AA',
  '10': 'AB',
  '11': 'AC',
  '12': 'AD',
  '13': 'AE',
  '14': 'AF',
  '15': 'AG',
  '16': 'AH',
  '17': 'AI',
  '18': 'AJ',
  '19': 'AK',
  '20': 'AL',
  '21': 'AM',
  '22': 'AN',
  '23': 'AO',
  '24': 'AP',
  '25': 'AQ',
  '26': 'AR',
  '27': 'AS',
  '28': 'AT',
  '29': 'AU',
  '30': 'AV',
  '31': 'AW',
};

var ViewData;
const BangumiName = [];
const RowData = [];
let YoutubeData = [];

console.log(RowData);

async function Start() {
  var limit = 100;
  var items;
  var item;
  youtube.addParam('order', 'date');
  youtube.addParam('type', 'video');
  youtube.addParam('regionCode', 'JP');
  youtube.addParam('publishedAfter', `${YoutubeDay}T00:00:00Z`); // 今月のみの動画情報を取得する
  youtube.addParam('channelId', ChannelID);
  youtube.addParam('pageToken', pagetoken); // ページトークンを追加することで50件以上前のデータも取得できる

  youtube.search(keyword, limit, async function (err, result) {
    if (err) {
      console.log(err);
      return;
    }
    if (pagetoken == '') {
      pagetoken = result['nextPageToken']; // 50件以上検索するためページトークンを代入
    }
    items = result['items'];
    for (var i in items) {
      item = items[i];
      const Data = {
        viewCount: null,
        title: item['snippet']['title'],
        id: item['id']['videoId'],
        publishedAt: null,
        url: `https://www.youtube.com/watch?v=${item['id']['videoId']}`,
      };
      let UploadDay = String(items[i]['snippet']['publishedAt']);
      UploadDay = UploadDay.split('T')[0]; // Tの文字列で区切る
      UploadDay = UploadDay.replace(/-/g, '/'); // ハイフンを全てスラッシュに変更
      Data['publishedAt'] = UploadDay;
      YoutubeData.push(Data);
    }
  });
}

//Start();

async function VideoData(id) {
  youtube.getById(id, async function (err, result) {
    if (err) {
      console.log(err);
      return;
    }
    ViewData = result['items'][0]['statistics'];
  });
}

async function test() {
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10),
  });
  // RPA管理スプシから集計用シートIDを取得
  const sheetdata = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: SheetID,
    range: `宣伝(YoutubeTwiiter)RPA管理シート!P3:P3`,
  });
  SheetID = sheetdata[0][0];

  let SheetData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: SheetID,
    range: `${SheetName}!B4:B20`,
  });
  // 番組名を取得
  for (let i in SheetData) {
    if (
      SheetData[i] == [] ||
      SheetData[i][0] == undefined ||
      SheetData[i][0] == ''
    ) {
      continue;
    }
    if (SheetData[i][0] == '対応者') {
      break;
    }
    if (SheetData[i][0] != '') {
      BangumiName.push(SheetData[i][0]); // B列の番組名を取得する
    }
  }
  console.log(BangumiName);
  SheetData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: SheetID,
    range: `${SheetName}!L1:R300`,
  });

  // 番組名ごとに、何行目から貼り付けるか判定する
  let j = 0;
  let Row = { Start: null, End: null, 貼り付け開始: null };
  for (let i in SheetData) {
    if (SheetData[i] == []) {
      continue;
    }
    if (SheetData[i][5] == 'タイトル') {
      continue;
    }
    if (SheetData[i][0] == BangumiName[j]) {
      Row['Start'] = Number(i) + 3;
      j += 1;
    }
    if (SheetData[i][2] == '1') {
      Row['End'] = Number(i) + 1;
    }
    if (
      Row['Start'] != null &&
      Row['貼り付け開始'] == null &&
      SheetData[i][5] != undefined
    ) {
      Row['貼り付け開始'] = Number(i);
    }
    if (Row['貼り付け開始'] == null) {
      Row['貼り付け開始'] = Row['End'];
    }
    if (Row['Start'] != null && Row['End'] != null) {
      RowData.push(Row);
      Row = { Start: null, End: null, 貼り付け開始: null };
    }
  }
  //console.log(BangumiName);
  console.log(RowData);
  //let Day = moment().format('YYYY-MM'); // Youtubeの今月検索用
  let Day = moment().subtract(1, 'month').format('YYYY-MM');
  Day = Day + String('-01');
  YoutubeDay = Day;

  await Start();
  await RPA.sleep(2000);
  await Start(); // 50件以上検索するためもう一度
  await RPA.sleep(2000);
  await Start();
  await RPA.sleep(2000);
  console.log(YoutubeData);
  console.log('今月データ個数', YoutubeData.length);
  let TodayYoutubeData = [];
  let youtubeURL = []; // 被りをなくすためにURLを格納させる
  for (let i in YoutubeData) {
    if (
      YoutubeData[i]['publishedAt'] == TodayDate &&
      youtubeURL.indexOf(YoutubeData[i]['url']) == -1
    ) {
      TodayYoutubeData.push(YoutubeData[i]);
      youtubeURL.push(YoutubeData[i]['url']); // 配列検索してURLがかぶっていないものだけ格納
    }
  }
  console.log('今日のデータ個数', TodayYoutubeData.length);
  TodayYoutubeData = TodayYoutubeData.reverse(); // 配列のデータを反転させる

  // それぞれの番組ごとに貼り付ける行数をだす
  for (let i in BangumiName) {
    let row = 0;
    for (let j in TodayYoutubeData) {
      if (TodayYoutubeData[j]['title'].includes(BangumiName[i]) == true) {
        TodayYoutubeData[j]['row'] = RowData[i]['貼り付け開始'] - row;
        row += 1;
      }
    }
  }
  console.log('全体データ:', YoutubeData);
  console.log('昨日データ:', TodayYoutubeData);
  //　スプシにタイトルとアップロード日を貼り付け
  for (let i in TodayYoutubeData) {
    let ReplaceTitle = TodayYoutubeData[i]['title']
      .replace(/&quot;/g, `"`)
      .replace(/&amp;/g, `&`);
    console.log('リプレイスタイトル:', ReplaceTitle);
    await TitleSetValues_function(
      [ReplaceTitle, TodayYoutubeData[i]['publishedAt']],
      TodayYoutubeData[i]['row']
    );
  }

  // スプシからタイトルを取得
  const SheetTitleData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: SheetID,
    range: `${SheetName}!Q1:Q300`,
  });
  for (let i in SheetTitleData) {
    if (
      SheetTitleData[i] == [] ||
      SheetTitleData[i][0] == undefined ||
      SheetTitleData[i][0] == '' ||
      SheetTitleData[i][0] == 'タイトル'
    ) {
      continue;
    }
    for (let j in YoutubeData) {
      var strHead = YoutubeData[j]['title']
        .replace(/&quot;/g, '')
        .replace(/&amp;/g, '');
      strHead = strHead.slice(0, 20);
      let replaceText = SheetTitleData[i][0]
        .replace(/"/g, '')
        .replace(/&/g, '');
      // 先頭10文字がスプシと同じなら 再生数を取得する
      if (replaceText.includes(strHead) == true) {
        console.log('先頭一致');
        await VideoData(YoutubeData[j]['id']); //再生数を取得
        await RPA.sleep(5000);
        YoutubeData[j]['viewCount'] = ViewData['viewCount'];
        console.log('行数:', Number(i) + 1, YoutubeData[j]);
        await CountViewSetValues_function(
          YoutubeData[j]['viewCount'],
          Number(i) + 1
        );
        break;
      }
    }
  }
}

test();

async function TitleSetValues_function(Text, Row) {
  try {
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: SheetID,
      range: `${SheetName}!Q${Row}:R${Row}`,
      parseValues: true,
      values: [Text],
    });
  } catch (Error) {
    RPA.Logger.info(Error);
    RPA.Logger.info('貼り付けエラー');
  }
}

async function CountViewSetValues_function(Text, Row) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: SheetID,
    range: `${SheetName}!${SheetColmun[SheetDay]}${Row}:${SheetColmun[SheetDay]}${Row}`,
    parseValues: true,
    values: [[`${Text}`]],
  });
}
