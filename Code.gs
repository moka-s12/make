const LIST_SHEET = "依頼一覧";
const DATA_SHEET = "回答データ";
const DETAIL_SHEET = "制作用詳細";

const ROOT_FOLDER_NAME = "one｜イラスト依頼";

const STATUS_LIST = [
  "未着手",
  "ヒアリング待ち",
  "制作待ち",
  "ラフ制作中",
  "ラフ確認待ち",
  "修正対応中",
  "清書中",
  "最終確認",
  "納品済み",
  "保留",
  "キャンセル"
];


/* =========================================================
   フォーム受信
========================================================= */

function doPost(e) {
  try {
    if (
      !e ||
      !e.postData ||
      !e.postData.contents
    ) {
      throw new Error(
        "送信データがありません"
      );
    }

    const order =
      JSON.parse(
        e.postData.contents
      );

    const data =
      order.data || {};
    validateEmoteOrder_(data);

    const ss =
      SpreadsheetApp
        .getActiveSpreadsheet();

    if (!ss) {
      throw new Error(
        "スプレッドシートを取得できませんでした"
      );
    }


    /* ---------- シート ---------- */

    const listSheet =
      getOrCreateSheet_(
        ss,
        LIST_SHEET
      );

    const dataSheet =
      getOrCreateSheet_(
        ss,
        DATA_SHEET
      );

    const detailSheet =
      getOrCreateSheet_(
        ss,
        DETAIL_SHEET
      );


    setupListSheet_(
      listSheet
    );

    setupDataSheet_(
      dataSheet
    );

    setupDetailSheet_(
      detailSheet
    );


    /* ---------- Drive ---------- */

    const driveData =
      saveUploadedImages_(
        order
      );


    /* ---------- 回答データ ---------- */

    const record =
      buildRecord_(
        order,
        driveData
      );

    const headers =
      getDataHeaders_();

    const row =
      headers.map(
        header =>
          record[header] ?? ""
      );

    dataSheet.appendRow(
      row
    );


    /* ---------- 依頼一覧 ---------- */

    const receipt =
      order.receipt || "";

    const listRow = [
      receipt,
      formatDate_(
        order.createdAt
      ),
      data.requester?.name || "",
      join_(
        data.request?.types
      ),
      "未着手",
      "",
      "",
      "詳細を見る"
    ];

    listSheet.appendRow(
      listRow
    );


    const rowNumber =
      listSheet.getLastRow();


    applyStatusValidation_(
      listSheet.getRange(
        rowNumber,
        5
      )
    );


    /* ---------- 詳細リンク ---------- */

    const detailUrl =
      ss.getUrl() +
      "#gid=" +
      detailSheet.getSheetId();


    const richText =
      SpreadsheetApp
        .newRichTextValue()
        .setText(
          "詳細を見る"
        )
        .setLinkUrl(
          detailUrl
        )
        .build();


    listSheet
      .getRange(
        rowNumber,
        8
      )
      .setRichTextValue(
        richText
      );


    /* ---------- 受付番号プルダウン ---------- */

    updateReceiptDropdown_(
      ss
    );


    return jsonResponse_({
      ok: true,
      receipt:
        receipt,
      folderUrl:
        driveData.folderUrl || ""
    });

  } catch (error) {

    console.error(
      error &&
      error.stack
        ? error.stack
        : error
    );

    return jsonResponse_({
      ok: false,
      error:
        String(
          error.message ||
          error
        )
    });
  }
}


/* =========================================================
   回答データ作成
========================================================= */

function buildRecord_(
  order,
  driveData
) {

  const data =
    order.data || {};

  const categories =
    driveData.categories || {};


  return {

    ...emoteRecord_(data, categories),
    "受付番号":
      order.receipt || "",

    "受付日時":
      formatDate_(
        order.createdAt
      ),

    "受付ステータス":
      order.status ||
      "受付済み",

    "依頼区分":
      data.route || "",


    /* ご依頼者 */

    "活動名":
      data.requester?.name || "",

    "Discord名":
      data.requester?.discord || "",

    "制作前ヒアリング":
      data.requester?.hearing || "",


    /* 制作内容 */

    "制作内容":
      join_(
        data.request?.types
      ),

    "イベント種類":
      join_(
        data.request?.eventTypes
      ),

    "イベント・企画名":
      data.request?.eventName || "",

    "その他の依頼内容":
      data.request?.other || "",


    /* キャラクター */

    "キャラクター種類":
      data.character?.type || "",

    "その他キャラクター":
      data.character?.otherType || "",

    "動物モチーフ":
      join_(
        data.character?.animalMotif
      ),

    "獣人度":
      data.character?.beastLevel || "",

    "性別":
      data.character?.gender || "",

    "見た目年齢":
      data.character?.visualAge || "",

    "身長指定":
      data.character?.heightType || "",

    "身長(cm)":
      data.character?.heightCm || "",

    "体型":
      data.character?.bodyType || "",

    "胸サイズ":
      data.character?.bustSize || "",


    /* 既存キャラ */

    "既存キャラ参考画像":
      imageLinksToText_(
        categories[
          "キャラクター参考"
        ]
      ),

    "既存キャラ補足":
      data.existingCharacter
        ?.notes || "",


    /* 髪 */

    "髪の長さ":
      data.hair?.length || "",

    "髪型系統":
      data.hair?.category || "",

    "髪型詳細":
      data.hair?.detail || "",

    "前髪系統":
      data.hair?.bangCategory || "",

    "前髪詳細":
      data.hair?.bangDetail || "",

    "髪色":
      data.hair?.color || "",

    "追加カラー種類":
      data.hair?.extraType || "",

    "追加カラー":
      data.hair?.extraColor || "",

    "髪参考画像":
      imageLinksToText_(
        categories[
          "髪型参考"
        ]
      ),


    /* 目 */

    "目の色":
      data.eyes?.color || "",

    "目の特徴":
      join_(
        data.eyes?.features
      ),

    "左目カラー":
      data.eyes?.leftColor || "",

    "右目カラー":
      data.eyes?.rightColor || "",

    "顔まわり補足":
      data.eyes?.faceNotes || "",


    /* デザイン */

    "雰囲気":
      join_(
        data.design?.mood
      ),

    "デザイン方向性":
      join_(
        data.design?.taste
      ),

    "メインカラー":
      join_(
        data.design?.mainColors
      ),

    "特徴・モチーフ":
      join_(
        data.design?.features
      ),

    "特徴補足":
      data.design?.featureNotes || "",


    /* 衣装 */

    "衣装系統":
      data.outfit?.category || "",

    "衣装詳細":
      data.outfit?.detail || "",

    "衣装参考画像":
      imageLinksToText_(
        categories[
          "衣装参考"
        ]
      ),

    "衣装補足":
      data.outfit?.notes || "",


    /* 既存衣装 */

    "既存衣装カテゴリ":
      data.existingOutfit
        ?.category || "",

    "既存衣装詳細":
      data.existingOutfit
        ?.detail || "",

    "既存衣装参考画像":
      imageLinksToText_(
        categories[
          "既存衣装参考"
        ]
      ),


    /* イラスト */

    "描画範囲":
      data.illustration?.range || "",

    "表情":
      join_(
        data.illustration
          ?.expressions
      ),

    "ポーズ":
      data.illustration?.pose || "",

    "絶対入れてほしいもの":
      data.illustration
        ?.mustHave || "",

    "NG要素":
      data.illustration?.ng || "",

    "イラスト参考画像":
      imageLinksToText_(
        categories[
          "イラスト参考"
        ]
      ),

    "その他・補足":
      data.illustration?.notes || "",


    /* Drive */

    "依頼フォルダ":
      driveData.folderUrl || ""
  };
}


/* =========================================================
   回答データ ヘッダー
========================================================= */

function getDataHeaders_() {

  return [

    "受付番号",
    "受付日時",
    "受付ステータス",
    "依頼区分",

    "活動名",
    "Discord名",
    "制作前ヒアリング",

    "制作内容",
    "イベント種類",
    "イベント・企画名",
    "その他の依頼内容",

    "キャラクター種類",
    "その他キャラクター",
    "動物モチーフ",
    "獣人度",
    "性別",
    "見た目年齢",
    "身長指定",
    "身長(cm)",
    "体型",
    "胸サイズ",

    "既存キャラ参考画像",
    "既存キャラ補足",

    "髪の長さ",
    "髪型系統",
    "髪型詳細",
    "前髪系統",
    "前髪詳細",
    "髪色",
    "追加カラー種類",
    "追加カラー",
    "髪参考画像",

    "目の色",
    "目の特徴",
    "左目カラー",
    "右目カラー",
    "顔まわり補足",

    "雰囲気",
    "デザイン方向性",
    "メインカラー",
    "特徴・モチーフ",
    "特徴補足",

    "衣装系統",
    "衣装詳細",
    "衣装参考画像",
    "衣装補足",

    "既存衣装カテゴリ",
    "既存衣装詳細",
    "既存衣装参考画像",

    "描画範囲",
    "表情",
    "ポーズ",
    "絶対入れてほしいもの",
    "NG要素",
    "イラスト参考画像",
    "その他・補足",

    "依頼フォルダ",
    ...EMOTE_HEADERS
  ];
}


/* =========================================================
   依頼一覧
========================================================= */

function setupListSheet_(
  sheet
) {

  const headers = [
    "受付番号",
    "受付日",
    "活動名",
    "制作内容",
    "ステータス",
    "担当",
    "納期",
    "詳細"
  ];


  if (
    sheet.getLastRow() === 0
  ) {

    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .setValues([
        headers
      ]);
  }


  sheet.setFrozenRows(1);


  sheet.setColumnWidth(
    1,
    180
  );

  sheet.setColumnWidth(
    2,
    140
  );

  sheet.setColumnWidth(
    3,
    180
  );

  sheet.setColumnWidth(
    4,
    220
  );

  sheet.setColumnWidth(
    5,
    160
  );

  sheet.setColumnWidth(
    6,
    120
  );

  sheet.setColumnWidth(
    7,
    130
  );

  sheet.setColumnWidth(
    8,
    120
  );


  sheet
    .getRange(
      "C:D"
    )
    .setWrap(true);


  if (
    sheet.getLastRow() >= 2
  ) {

    applyStatusValidation_(
      sheet.getRange(
        2,
        5,
        sheet.getLastRow() - 1,
        1
      )
    );
  }
}


/* =========================================================
   ステータス
========================================================= */

function applyStatusValidation_(
  range
) {

  const rule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        STATUS_LIST,
        true
      )
      .setAllowInvalid(false)
      .build();


  range.setDataValidation(
    rule
  );
}


/* =========================================================
   回答データシート
========================================================= */

function setupDataSheet_(
  sheet
) {

  const headers =
    getDataHeaders_();

  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }


  if (
    sheet.getLastRow() === 0
  ) {

    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .setValues([
        headers
      ]);
  }


  // 既存列の順番を変えず、追加列のみ末尾へ移行する。
  const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const legacy = headers.slice(0, headers.length - EMOTE_HEADERS.length);
  if (legacy.some((header, i) => existing[i] !== header)) throw new Error("回答データの既存列が想定と異なります。列順を確認してください。");
  if (existing.length > headers.length || existing.some((header, i) => header !== headers[i])) throw new Error("回答データの追加列が想定と異なります。");
  if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}


/* =========================================================
   制作用詳細

   ※ ここは毎回消さない！
========================================================= */

function setupDetailSheet_(
  sheet
) {

  const currentTitle =
    sheet
      .getRange("A1")
      .getValue();


  /*
    初回だけ作成
  */

  if (
    currentTitle !==
    "🎨 制作用詳細"
  ) {

    sheet.clear();


    /*
      念のため結合解除
    */

    try {
      sheet
        .getRange(
          "A1:B1"
        )
        .breakApart();
    } catch (error) {}


    sheet
      .getRange(
        "A1:B1"
      )
      .merge();


    sheet
      .getRange(
        "A1"
      )
      .setValue(
        "🎨 制作用詳細"
      );


    sheet
      .getRange(
        "A2"
      )
      .setValue(
        "受付番号"
      );
  }


  sheet.setColumnWidth(
    1,
    190
  );

  sheet.setColumnWidth(
    2,
    500
  );


  sheet
    .getRange(
      "A:B"
    )
    .setVerticalAlignment(
      "top"
    );


  sheet
    .getRange(
      "B:B"
    )
    .setWrap(true);
}


/* =========================================================
   受付番号プルダウン
========================================================= */

function updateReceiptDropdown_(
  ss
) {

  const dataSheet =
    ss.getSheetByName(
      DATA_SHEET
    );

  const detailSheet =
    ss.getSheetByName(
      DETAIL_SHEET
    );


  if (
    !dataSheet ||
    !detailSheet
  ) {
    return;
  }


  const lastRow =
    dataSheet.getLastRow();


  if (
    lastRow < 2
  ) {
    return;
  }


  const receiptRange =
    dataSheet.getRange(
      2,
      1,
      lastRow - 1,
      1
    );


  const rule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInRange(
        receiptRange,
        true
      )
      .setAllowInvalid(false)
      .build();


  detailSheet
    .getRange(
      "B2"
    )
    .setDataValidation(
      rule
    );
}


/* =========================================================
   制作用詳細 選択
========================================================= */

function onEdit(e) {

  if (
    !e ||
    !e.range
  ) {
    return;
  }


  const sheet =
    e.range.getSheet();


  if (
    sheet.getName() !==
    DETAIL_SHEET
  ) {
    return;
  }


  if (
    e.range.getA1Notation() !==
    "B2"
  ) {
    return;
  }


  const receipt =
    String(
      e.range.getValue() || ""
    ).trim();


  renderDetail_(
    receipt
  );
}


/* =========================================================
   制作用詳細 描画
========================================================= */

function renderDetail_(
  receipt
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const dataSheet =
    ss.getSheetByName(
      DATA_SHEET
    );


  const detailSheet =
    ss.getSheetByName(
      DETAIL_SHEET
    );


  if (
    !dataSheet ||
    !detailSheet
  ) {
    return;
  }


  clearDetailBody_(
    detailSheet
  );


  if (!receipt) {
    return;
  }


  const values =
    dataSheet
      .getDataRange()
      .getValues();


  if (
    values.length < 2
  ) {
    return;
  }


  const headers =
    values[0];


  const row =
    values
      .slice(1)
      .find(
        item =>
          String(
            item[0]
          ) ===
          String(
            receipt
          )
      );


  if (!row) {

    detailSheet
      .getRange(
        "A4"
      )
      .setValue(
        "該当する依頼が見つかりませんでした"
      );

    return;
  }


  const record = {};


  headers.forEach(
    (header, index) => {

      record[
        String(header)
      ] =
        row[index];
    }
  );


  let nextRow = 4;
  nextRow = addSection_(detailSheet, nextRow, "有償エモート制作", EMOTE_HEADERS.map(header => [header, record[header]]));


  nextRow =
    addSection_(
      detailSheet,
      nextRow,
      "基本情報",
      [
        [
          "受付番号",
          record["受付番号"]
        ],
        [
          "受付日時",
          record["受付日時"]
        ],
        [
          "依頼区分",
          record["依頼区分"]
        ],
        [
          "活動名",
          record["活動名"]
        ],
        [
          "Discord名",
          record["Discord名"]
        ],
        [
          "制作前ヒアリング",
          record[
            "制作前ヒアリング"
          ]
        ],
        [
          "依頼フォルダ",
          record["依頼フォルダ"]
        ]
      ]
    );


  nextRow =
    addSection_(
      detailSheet,
      nextRow,
      "依頼内容",
      [
        [
          "制作内容",
          record["制作内容"]
        ],
        [
          "イベント種類",
          record["イベント種類"]
        ],
        [
          "イベント・企画名",
          record[
            "イベント・企画名"
          ]
        ],
        [
          "その他",
          record[
            "その他の依頼内容"
          ]
        ]
      ]
    );


  nextRow =
    addSection_(
      detailSheet,
      nextRow,
      "キャラクター",
      [
        [
          "キャラクター種類",
          record[
            "キャラクター種類"
          ]
        ],
        [
          "その他キャラクター",
          record[
            "その他キャラクター"
          ]
        ],
        [
          "動物モチーフ",
          record[
            "動物モチーフ"
          ]
        ],
        [
          "獣人度",
          record["獣人度"]
        ],
        [
          "性別",
          record["性別"]
        ],
        [
          "見た目年齢",
          record["見た目年齢"]
        ],
        [
          "身長指定",
          record["身長指定"]
        ],
        [
          "身長(cm)",
          record["身長(cm)"]
        ],
        [
          "体型",
          record["体型"]
        ],
        [
          "胸サイズ",
          record["胸サイズ"]
        ],
        [
          "既存キャラ参考画像",
          record[
            "既存キャラ参考画像"
          ]
        ],
        [
          "既存キャラ補足",
          record[
            "既存キャラ補足"
          ]
        ]
      ]
    );


  nextRow =
    addSection_(
      detailSheet,
      nextRow,
      "髪・顔",
      [
        [
          "髪の長さ",
          record["髪の長さ"]
        ],
        [
          "髪型系統",
          record["髪型系統"]
        ],
        [
          "髪型詳細",
          record["髪型詳細"]
        ],
        [
          "前髪系統",
          record["前髪系統"]
        ],
        [
          "前髪詳細",
          record["前髪詳細"]
        ],
        [
          "髪色",
          record["髪色"]
        ],
        [
          "追加カラー種類",
          record[
            "追加カラー種類"
          ]
        ],
        [
          "追加カラー",
          record["追加カラー"]
        ],
        [
          "髪参考画像",
          record["髪参考画像"]
        ],
        [
          "目の色",
          record["目の色"]
        ],
        [
          "目の特徴",
          record["目の特徴"]
        ],
        [
          "左目カラー",
          record["左目カラー"]
        ],
        [
          "右目カラー",
          record["右目カラー"]
        ],
        [
          "顔まわり補足",
          record["顔まわり補足"]
        ]
      ]
    );


  nextRow =
    addSection_(
      detailSheet,
      nextRow,
      "デザイン",
      [
        [
          "雰囲気",
          record["雰囲気"]
        ],
        [
          "デザイン方向性",
          record[
            "デザイン方向性"
          ]
        ],
        [
          "メインカラー",
          record[
            "メインカラー"
          ]
        ],
        [
          "特徴・モチーフ",
          record[
            "特徴・モチーフ"
          ]
        ],
        [
          "特徴補足",
          record["特徴補足"]
        ]
      ]
    );


  nextRow =
    addSection_(
      detailSheet,
      nextRow,
      "衣装",
      [
        [
          "衣装系統",
          record["衣装系統"]
        ],
        [
          "衣装詳細",
          record["衣装詳細"]
        ],
        [
          "衣装参考画像",
          record[
            "衣装参考画像"
          ]
        ],
        [
          "衣装補足",
          record["衣装補足"]
        ],
        [
          "既存衣装カテゴリ",
          record[
            "既存衣装カテゴリ"
          ]
        ],
        [
          "既存衣装詳細",
          record[
            "既存衣装詳細"
          ]
        ],
        [
          "既存衣装参考画像",
          record[
            "既存衣装参考画像"
          ]
        ]
      ]
    );


  addSection_(
    detailSheet,
    nextRow,
    "イラスト",
    [
      [
        "描画範囲",
        record["描画範囲"]
      ],
      [
        "表情",
        record["表情"]
      ],
      [
        "ポーズ",
        record["ポーズ"]
      ],
      [
        "絶対入れてほしいもの",
        record[
          "絶対入れてほしいもの"
        ]
      ],
      [
        "NG要素",
        record["NG要素"]
      ],
      [
        "イラスト参考画像",
        record[
          "イラスト参考画像"
        ]
      ],
      [
        "その他・補足",
        record[
          "その他・補足"
        ]
      ]
    ]
  );
}


/* =========================================================
   詳細画面クリア
========================================================= */

function clearDetailBody_(
  sheet
) {

  const maxRows =
    Math.max(
      sheet.getMaxRows() - 3,
      1
    );


  const range =
    sheet.getRange(
      4,
      1,
      maxRows,
      2
    );


  try {
    range.breakApart();
  } catch (error) {}


  range.clearContent();
}


/* =========================================================
   詳細 セクション追加

   空欄は表示しない
   全部空ならセクション自体出さない
========================================================= */

function addSection_(
  sheet,
  startRow,
  title,
  fields
) {

  const visibleFields =
    fields.filter(
      ([label, value]) =>
        hasValue_(value)
    );


  if (
    !visibleFields.length
  ) {
    return startRow;
  }


  sheet
    .getRange(
      startRow,
      1,
      1,
      2
    )
    .merge();


  sheet
    .getRange(
      startRow,
      1
    )
    .setValue(
      `【${title}】`
    )
    .setFontWeight(
      "bold"
    );


  let row =
    startRow + 1;


  visibleFields.forEach(
    ([label, value]) => {

      sheet
        .getRange(
          row,
          1
        )
        .setValue(
          label
        )
        .setFontWeight(
          "bold"
        );


      const valueCell =
        sheet.getRange(
          row,
          2
        );


      valueCell
        .setValue(
          value
        )
        .setWrap(true);


      /*
        URLだけの項目なら
        クリック可能にする
      */

      if (
        typeof value ===
          "string" &&
        /^https?:\/\/\S+$/
          .test(
            value.trim()
          )
      ) {

        const richText =
          SpreadsheetApp
            .newRichTextValue()
            .setText(
              label ===
                "依頼フォルダ"
                ? "📁 依頼フォルダを開く"
                : "画像を開く"
            )
            .setLinkUrl(
              value.trim()
            )
            .build();


        valueCell
          .setRichTextValue(
            richText
          );
      }


      row++;
    }
  );


  return row + 1;
}


/* =========================================================
   DRIVE 保存
========================================================= */

function saveUploadedImages_(
  order
) {

  const uploads =
    Array.isArray(
      order.uploads
    )
      ? order.uploads
      : [];


  /*
    画像なしなら
    フォルダを作らない
  */

  if (
    !uploads.length
  ) {

    return {
      folderUrl: "",
      categories: {}
    };
  }


  const rootFolder =
    getOrCreateRootFolder_();


  const activityName =
    sanitizeFolderName_(
      order.data
        ?.requester
        ?.name ||
      "名称未設定"
    );


  const receipt =
    sanitizeFolderName_(
      order.receipt ||
      "受付番号なし"
    );


  const requestFolder =
    rootFolder
      .createFolder(
        `${receipt}｜${activityName}`
      );


  const categoryFolders = {};

  const categoryUrls = {};


  uploads.forEach(
    upload => {

      if (
        !upload ||
        !(upload.data || upload.base64) ||
        !upload.name
      ) {
        return;
      }


      const category =
        sanitizeFolderName_(
          upload.category ||
          "その他参考"
        );


      if (
        !categoryFolders[
          category
        ]
      ) {

        const folder =
          requestFolder
            .createFolder(
              category
            );


        categoryFolders[
          category
        ] =
          folder;


        categoryUrls[
          category
        ] =
          [];
      }


      const bytes =
        Utilities
          .base64Decode(
            upload.data || upload.base64
          );


      const blob =
        Utilities
          .newBlob(
            bytes,
            upload.type ||
              "application/octet-stream",
            upload.name
          );


      const file =
        categoryFolders[
          category
        ]
          .createFile(
            blob
          );


      categoryUrls[
        category
      ].push({
        name:
          file.getName(),

        url:
          file.getUrl()
      });
    }
  );


  return {
    folderUrl:
      requestFolder
        .getUrl(),

    categories:
      categoryUrls
  };
}


/* =========================================================
   Drive 親フォルダ
========================================================= */

function getOrCreateRootFolder_() {

  const folders =
    DriveApp
      .getFoldersByName(
        ROOT_FOLDER_NAME
      );


  if (
    folders.hasNext()
  ) {
    return folders.next();
  }


  return DriveApp
    .createFolder(
      ROOT_FOLDER_NAME
    );
}


/* =========================================================
   フォルダ名 安全化
========================================================= */

function sanitizeFolderName_(
  value
) {

  return String(
    value || ""
  )
    .replace(
      /[\\/:*?"<>|]/g,
      "＿"
    )
    .trim();
}


/* =========================================================
   画像リンク → セル用テキスト
========================================================= */

function imageLinksToText_(
  files
) {

  if (
    !Array.isArray(files) ||
    !files.length
  ) {
    return "";
  }


  return files
    .map(
      file =>
        `${file.name}\n${file.url}`
    )
    .join(
      "\n\n"
    );
}


/* =========================================================
   シート取得
========================================================= */

function getOrCreateSheet_(
  ss,
  name
) {

  return (
    ss.getSheetByName(
      name
    ) ||
    ss.insertSheet(
      name
    )
  );
}


/* =========================================================
   値あり？
========================================================= */

function hasValue_(
  value
) {

  if (
    value === null ||
    value === undefined
  ) {
    return false;
  }


  if (
    Array.isArray(value)
  ) {
    return (
      value.length > 0
    );
  }


  return (
    String(value)
      .trim() !== ""
  );
}


/* =========================================================
   配列 → 文字
========================================================= */

function join_(
  value
) {

  if (
    Array.isArray(value)
  ) {

    return value
      .filter(Boolean)
      .join("、");
  }


  return value || "";
}


/* =========================================================
   日付
========================================================= */

function formatDate_(
  value
) {

  if (!value) {
    return "";
  }


  const date =
    new Date(value);


  if (
    Number.isNaN(
      date.getTime()
    )
  ) {
    return String(value);
  }


  return Utilities
    .formatDate(
      date,
      "Asia/Tokyo",
      "yyyy/MM/dd HH:mm"
    );
}


/* =========================================================
   JSON返却
========================================================= */

function jsonResponse_(
  data
) {

  return ContentService
    .createTextOutput(
      JSON.stringify(
        data
      )
    )
    .setMimeType(
      ContentService
        .MimeType
        .JSON
    );
}
const EMOTE_PRICES = {5: 1000, 10: 1800, 15: 2500};
const EMOTE_STYLES = ["① 3D（リアル寄り・立体感）", "② 2D（アニメ塗り・高精細）", "③ 平面（シンプル・デフォルメ）", "④ 立体的（ふんわり・ぷっくり）"];
const EMOTE_HEADERS = ["エモート個数", "エモート料金(円)", "エモート絵柄", "エモート希望", "エモート参考画像", "納品形式", "支払方法", "支払時期", "支払条件確認"];
function validateEmoteOrder_(data) {
  if (data.route !== "有償エモート制作") return;
  const e = data.emote || {};
  const count = Number(e.count);
  if (!EMOTE_PRICES[count]) throw new Error("エモート個数が不正です");
  if (!Array.isArray(e.styles) || !e.styles.length || e.styles.some(style => !EMOTE_STYLES.includes(style))) throw new Error("絵柄を確認してください");
  if (!String(e.wishes || "").trim()) throw new Error("エモートの希望を入力してください");
  if (e.paymentAgreed !== true) throw new Error("支払条件の確認が必要です");
  if (!String(data.requester?.name || "").trim() || !String(data.requester?.discord || "").trim() || !["希望する", "希望しない"].includes(data.requester?.hearing)) throw new Error("依頼者情報を確認してください");
  e.count = count;
  e.price = EMOTE_PRICES[count];
  e.styles = [...new Set(e.styles)];
  e.format = "背景透過PNG";
  e.paymentMethod = "要相談";
  e.paymentTiming = "前払い";
  data.request = {types: ["有償エモート制作"]};
}
function emoteRecord_(data, categories) {
  if (data.route !== "有償エモート制作") return {};
  const e = data.emote;
  return {"エモート個数": e.count, "エモート料金(円)": EMOTE_PRICES[e.count], "エモート絵柄": join_(e.styles), "エモート希望": e.wishes, "エモート参考画像": imageLinksToText_(categories["エモート参考"]), "納品形式": e.format, "支払方法": e.paymentMethod, "支払時期": e.paymentTiming, "支払条件確認": "確認済み"};
}
