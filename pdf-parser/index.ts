import xlsx, { CellObject } from 'xlsx';

const DATA_PATH = `${__dirname}/../test/20230529_weekly.xlsx`;

type DayOfWeekKeys = 'MON' | 'TUE' | 'WED' | 'THU' | 'FRI' | 'SAT' | 'SUN';
const dayOfWeekKeys = [
  'MON',
  'TUE',
  'WED',
  'THU',
  'FRI',
  'SAT',
  'SUN',
] satisfies DayOfWeekKeys[];

const dayOfWeekTexts = {
  MON: '月曜日',
  TUE: '火曜日',
  WED: '水曜日',
  THU: '木曜日',
  FRI: '金曜日',
  SAT: '土曜日',
  SUN: '日曜日',
} satisfies Record<DayOfWeekKeys, string>;

type Category = { name: string; defaultPrice?: string };
const categories = [
  {
    name: '電大唐揚げ専門店',
  },
  {
    name: 'バラエティー',
  },
  {
    name: 'カレー',
  },
  {
    name: 'うどん/そば',
    defaultPrice: '￥350 大盛¥400',
  },
  {
    name: 'ラーメン',
    defaultPrice: '￥350 大盛¥400',
  },
  {
    name: 'パスタ',
    defaultPrice: '￥400',
  },
  {
    name: '日替わり定食',
    defaultPrice: '¥500',
  },
] as Category[];

type Menu = {
  caterogyName: string;
  name: string;
  price: string;
  isSpecialMenu: boolean;
};

// WorkSheet を読み込み
const book = xlsx.readFile(DATA_PATH, {
  cellStyles: true,
});
const sheet = book.Sheets[book.SheetNames[0]];

// シートオブジェクトのキーの内、中身の入ったセルデータのキーのみ抽出
const cellKeys = Object.keys(sheet).filter(
  (n) => !n.startsWith('!') && sheet[n].t !== 'z'
);
const dayofweekStartCells = {
  MON: null,
  TUE: null,
  WED: null,
  THU: null,
  FRI: null,
  SAT: null,
  SUN: null,
} as Record<DayOfWeekKeys, { row: number; column: string } | null>;

// 先ほど抽出したセルの中から曜日文字列の含まれたセルを検索
// 同時に全ての列の中で空欄でない一番下のセルを取得する
let maxRowNum = 1;
for (const key of cellKeys) {
  const column = key.substring(0, 1);
  const row = parseInt(key.substring(1));
  const value = sheet[key] as CellObject;
  if (maxRowNum < row) maxRowNum = row;
  for (const dow of dayOfWeekKeys) {
    if (`${value.v}`.match(dow) != null) {
      dayofweekStartCells[dow] = {
        row,
        column,
      };
    }
  }
}

// メニュー探索
const merges = sheet['!merges'] ?? [];
const dayofweekDataAry: Record<
  DayOfWeekKeys,
  { value: string; isPaintedCell: boolean }[] | null
> = {
  MON: null,
  TUE: null,
  WED: null,
  THU: null,
  FRI: null,
  SAT: null,
  SUN: null,
};
for (const dow of dayOfWeekKeys) {
  const cell = dayofweekStartCells[dow];
  if (cell === null) {
    console.error(`Failed get ${dow}`);
    continue;
  }
  let row = cell.row;
  let column = cell.column;
  let values: { value: string; isPaintedCell: boolean }[] = [];
  // 一番下のセルは電大ソフトとかのセルなのでその手前で止める
  const smaxColumnNum = maxRowNum - 2;
  for (let srow = row + 1; srow < smaxColumnNum; srow++) {
    // sclm => ASCII Code
    for (let sclm = column.charCodeAt(0); sclm > 67; sclm--) {
      const value = sheet[`${String.fromCharCode(sclm)}${srow}`] as
        | CellObject
        | undefined;
      // console.log(`👀 ${String.fromCharCode(sclm)}${srow}`);
      // 中身が入っているセルがあればそれをメニュー名として確定する
      if (value !== undefined && value.t !== 'z') {
        // console.log(`Found: ${value.v}`);
        values.push({
          value: `${value.v}`,
          isPaintedCell: value.s?.patternType !== 'none',
        });
        break;
      }
      // カレーなどのメニューは複数日連続している場合は結合セルに入っている場合がある
      // なので、今見たセルが横方向の結合セルに含まれている場合は左に進んで検索を継続する
      // 含まれていない場合はその日の当該カテゴリメニューは提供されていない可能性がある (ex. 土曜日M2階など)
      let clmNumZStart = sclm - 65; // A => 0
      let rowNumZStart = srow - 1;
      let isMerged = false;
      for (const { s: sm, e: em } of merges) {
        if (
          sm.c <= clmNumZStart &&
          clmNumZStart <= em.c &&
          sm.r === rowNumZStart &&
          rowNumZStart === em.r
        ) {
          isMerged = true;
          break;
        }
      }
      // どの結合セルにも含まれていなければメニューが入っているセルではないと判断して下へ進む
      if (!isMerged) break;
    }
  }
  dayofweekDataAry[dow] = values;
}

// メニューオブジェクトに変換
const cnmax = categories.length;
const dayOfWeekMenus: Record<DayOfWeekKeys, Menu[]> = {
  MON: [],
  TUE: [],
  WED: [],
  THU: [],
  FRI: [],
  SAT: [],
  SUN: [],
};
for (const dow of dayOfWeekKeys) {
  const ary = dayofweekDataAry[dow];
  if (ary === null) continue;
  const len = ary.length;
  let cn = 0;
  for (let n = 0; n < len; n++) {
    const category = categories[cn];
    let price = '';
    // 個別に金額が設定されている場合はメニュー名に混ざっていることがあるのでそれを使用する
    // ついでにメニュー名の改行を除去する
    // thx https://stackoverflow.com/questions/10805125/how-to-remove-all-line-breaks-from-a-string
    let name = ary[n].value.replace(/(\r\n|\n|\r)/gm, '');
    if (name.includes('¥')) {
      const pos = name.indexOf('¥');
      price = name.substring(pos);
      name = name.substring(0, pos);
    }
    // メニュー名のセルに色がついている場合は特別メニューとして取り扱う
    // (値段のセルには常に色がついているので注意)
    const isSpecialMenu = ary[n].isPaintedCell;
    // メニュー名に金額が含まれていない場合はデフォルト価格か価格セルの内容を使用
    price =
      price === ''
        ? category.defaultPrice === undefined
          ? ary[++n].value
          : category.defaultPrice
        : price;
    // 複数のスペースを除去
    price = price.replace(/ +/g, ' ');
    // 円マークがついていなければ追加
    if (/^[0-9]+$/.test(price)) {
      price = `¥${price}`;
    }
    // 円マークの表記を統一
    price = price.replace(/￥/g, '¥');
    dayOfWeekMenus[dow].push({
      caterogyName: category.name,
      name,
      price,
      isSpecialMenu,
    });
    if (cnmax == ++cn) break;
  }
}

// メニュー出力
for (const dow of dayOfWeekKeys) {
  console.log(`=== ${dayOfWeekTexts[dow]} のメニュー ===`);
  if (dayOfWeekMenus[dow].length === 0) {
    console.log('この日は営業していません');
  } else {
    for (const menu of dayOfWeekMenus[dow]) {
      console.log(menu.caterogyName);
      console.log(`\t${menu.isSpecialMenu ? '✨ ' : ''}${menu.name}`);
      console.log(`\t${menu.price}`);
    }
  }
  console.log('');
}
