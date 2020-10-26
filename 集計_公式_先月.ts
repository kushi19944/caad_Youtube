import RPA from 'ts-rpa';
const moment = require('moment');

var Youtube = require('youtube-node');
var youtube = new Youtube();
youtube.setKey(process.env.YoutubeNode_API_CAAD);

// 流用時の変更箇所
// 公式 先月
const ChannelID = 'UCLsdm7nCJCVTWSid7G_f0Pg'; // 取得したいチャンネルIDを入れる
let SheetID = process.env.RPA_Senden_SheetID; //RPA管理シートID
const SheetName = '公式';
//

var keyword = ''; // キーワードは無しにするとそのチャンネル全て取得できる
let pagetoken = ''; //50件以上検索するためのページトークン

const SheetDay = moment().format('D'); // スプシに貼り付ける日付を取得
let YoutubeDay = moment().subtract(2, 'days').format('YYYY-MM-DD'); // 昨日を取得
console.log(SheetDay, '日');
const TodayDate = moment().subtract(1, 'days').format('YYYY/MM/DD'); // 昨日の日付を取得

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
    pagetoken = result['nextPageToken']; // 50件以上検索するためページトークンを代入
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
  let Day = moment().subtract(1, 'month').format('YYYY-MM');
  Day = Day + String('-01');
  YoutubeDay = Day;
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10),
  });
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10),
  });
  // RPA管理スプシから集計用シートIDを取得
  const sheetdata = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: SheetID,
    range: `宣伝(YoutubeTwiiter)RPA管理シート!Q3:Q3`,
  });
  SheetID = sheetdata[0][0];
  await Start();
  await RPA.sleep(2000);
  await Start(); // 50件以上検索するためもう一度
  await RPA.sleep(2000);
  await Start(); // 100件以上検索するためもう一度
  await RPA.sleep(2000);
  await Start(); // 100件以上検索するためもう一度
  await RPA.sleep(2000);
  for (let i in YoutubeData){
    console.log(YoutubeData[i])
  }
  //console.log(YoutubeData);
  await JudgeColumn(); // 貼り付ける列を検索する

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

async function CountViewSetValues_function(Text, Row) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: SheetID,
    range: `${SheetName}!${Column}${Row}:${Column}${Row}`,
    parseValues: true,
    values: [[`${Text}`]],
  });
}

let Column = 0;
async function JudgeColumn() {
  const ColumnData = {
    '-2': 'AU',
    '-1': 'AV',
    '0': 'AW',
    '1': 'AX',
    '2': 'AY',
    '3': 'AZ',
    '4': 'BA',
    '5': 'BB',
    '6': 'BC',
    '7': 'BD',
    '8': 'BE',
    '9': 'BF',
    '10': 'BG',
    '11': 'BH',
    '12': 'BI',
    '13': 'BJ',
    '14': 'BK',
    '15': 'BL',
    '16': 'BM',
    '17': 'BN',
    '18': 'BO',
    '19': 'BP',
    '20': 'BQ',
    '21': 'BR',
    '22': 'BS',
    '23': 'BT',
    '24': 'BU',
    '25': 'BV',
    '26': 'BW',
    '27': 'BX',
    '28': 'BY',
    '29': 'BZ',
    '30': 'CA',
    '31': 'CB',
  };
  const data = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: SheetID,
    range: `${SheetName}!AX3:AX3`,
  });
  const JudgeColumn = data[0][0].split('/')[1];
  // AX列が1日ならそのままAX列に貼り付けるように。(月ごとに貼り付ける列が変化するため)
  console.log(JudgeColumn);
  if (JudgeColumn == '2') {
    Column = -1;
  }
  if (JudgeColumn == '3') {
    Column = -2;
  }
  if (JudgeColumn == '4') {
    Column = -3;
  }
  const A = Number(SheetDay) + Number(Column);
  Column = ColumnData[A];
  RPA.Logger.info('貼り付け列:', Column);
}
