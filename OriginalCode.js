/*
 ***************
 * update1 2024/08/07 見積区分, 材料証明書,熱処理,表面処理,寸法記録削除(フィールドは現存)*
 ***************
 */

(() => {
  "use strict";

  let hot; // handsontable用変数

  // kintoneのレコード更新・追加時は、レコード番号などアップデートできないフィールドがあるので、除外するためのメソッド
  const setParams = (record) => {
    const result = {};
    for (const prop in record) {
      if (
        [
          "レコード番号",
          "作成日時",
          "更新日時",
          "作成者",
          "更新者",
          "ステータス",
          "作業者",
        ].indexOf(prop) === -1
      ) {
        result[prop] = record[prop];
      }
    }
    return result;
  };

  // kintoneのレコード取得用メソッド
  const getRecords = (callback, errorCallback) => {
    kintone.api(
      "/k/v1/records",
      "GET",
      { app: kintone.app.getId(), query: '案件管理レコード番号 = "5580"' },
      (resp) => {
        callback(resp);
      },
      (resp) => {
        errorCallback(resp);
      }
    );
  };

  // kintoneのレコード更新、追加用メソッド
  const saveRecords = (records, changedDatas, callback, errorCallback) => {
    const requests = [];
    const updateRecords = [];
    const insertRecords = [];
    let changedRows = [];

    // 変更されたセルの配列から、変更があった行だけ抜き出す
    for (let i = 0; i < changedDatas.length; i++) {
      changedRows.push(changedDatas[i][0]);
    }
    // 変更があった行番号の重複を排除
    changedRows = changedRows.filter((x, i, self) => {
      return self.indexOf(x) === i;
    });

    // 変更があった行から、レコード追加か変更かを判断し、クエリをつくる
    for (let i = 0; i < changedRows.length; i++) {
      if (
        !records[changedRows[i]]["レコード番号"] ||
        !records[changedRows[i]]["レコード番号"].value
      ) {
        insertRecords.push(setParams(records[changedRows[i]]));
      } else {
        updateRecords.push({
          id: records[changedRows[i]]["レコード番号"].value,
          record: setParams(records[changedRows[i]]),
        });
      }
    }

    if (updateRecords.length > 0) {
      // 更新用bulkRequest
      requests.push({
        method: "PUT",
        api: "/k/v1/records.json",
        payload: {
          app: kintone.app.getId(),
          records: updateRecords,
        },
      });
    }

    if (insertRecords.length > 0) {
      // 追加用bulkRequest
      requests.push({
        method: "POST",
        api: "/k/v1/records.json",
        payload: {
          app: kintone.app.getId(),
          records: insertRecords,
        },
      });
    }

    if (requests.length > 0) {
      // bulkrequestで一括で追加、更新。
      // 失敗した場合はロールバックされる。
      kintone.api(
        "/k/v1/bulkRequest",
        "POST",
        { requests: requests },
        (resp) => {
          console.dir(requests);
          console.dir(resp);
          callback(resp);
        },
        (resp) => {
          errorCallback(resp);
        }
      );
    }
  };

  // kintoneのレコード削除用メソッド
  const deleteRecords = (records, index, amount, callback, errorCallback) => {
    let i;
    const ids = [];
    for (i = index; i < index + amount; i++) {
      if (records[i] && records[i]["レコード番号"] && records[i]["レコード番号"].value) {
        ids.push(records[i]["レコード番号"].value);
      }
    }
    if (ids.length > 0) {
      kintone.api(
        "/k/v1/records",
        "DELETE",
        { app: kintone.app.getId(), ids: ids },
        (resp) => {
          callback(resp);
        },
        (resp) => {
          errorCallback(resp);
        }
      );
    } else {
      callback({ message: "No records to delete" });
    }
  };

  // 一覧ビュー表示用のイベントハンドラ
  kintone.events.on(["app.record.index.show"], (event) => {
    if (event.viewId !== 6449627) return;

    const container = document.getElementById("sheet");

    // handsontable初期化
    hot = new Handsontable(container, {
      // この時点ではdataは入力せず、あとから読み込ませるようにする。（データ更新時も再読み込みさせたいため）
      data: [],

      // 空白行
      minSpareRows: 10,

      // 表示したいカラム
      colHeaders: [
        "発注日",
        "工事No",
        "変更工事No",
        "納品",
        "部品名",
        "図番",
        "材質",
        "個数",
        " ",
        "担当",
        "発注先",
        "納期",
        "変更納期",
        "納品チェック",
        "研修チェック",
        "見積_単価",
        "見積_金額",
        "備考",
        "依頼先",
        "原価_単価",
        "原価_総額",
        "状態",
        "ご納品日",
        "発注書送付",
        "加工品リスト共有",
        "見積提出",
        "請求書提出",
        "状態",
        "材料証明書",
        "熱処理",
        "表面処理",
        "寸法記録",
        "案件管理レコード番号",
        "レコード番号",
      ],

      // コンテキストメニュー（右クリックメニュー）を指定。今回は削除用メニューのみ。
      contextMenu: ["remove_row"],

      licenseKey: "non-commercial-and-evaluation",

      // 必要に応じてreadOnlyの指定ができます。
      columns: [
        { data: "発注日.value", type: "date", dateFormat: "YYYY-MM-DD" },
        { data: "工事No.value" },
        { data: "変更工事No.value" },
        { data: "納品.value" },
        { data: "部品名.value" },
        { data: "図番.value" },
        { data: "材質.value" },
        { data: "個数.value" },
        { data: "空白１.value" },
        { data: "担当.value" },
        { data: "発注先.value" },
        { data: "納期.value", type: "date", dateFormat: "YYYY-MM-DD" },
        { data: "変更納期.value", type: "date", dateFormat: "YYYY-MM-DD" },
        { data: "納品チェック.value", type: "date", dateFormat: "YYYY-MM-DD" },
        { data: "研修チェック.value", type: "date", dateFormat: "YYYY-MM-DD" },
        { data: "見積_単価.value" },
        { data: "見積_金額.value" },
        { data: "備考.value" },
        { data: "依頼先.value" },
        { data: "原価_単価.value" },
        { data: "原価_総額.value" },
        {
          data: "状態_ドロップダウン.value",
          type: "dropdown",
          source: ["内示済み", "失注", " "],
        },
        { data: "ご納品日.value", type: "date", dateFormat: "YYYY-MM-DD" },
        { data: "発注書送付.value", type: "dropdown", source: ["済", ""] },
        { data: "加工品リスト共有.value", type: "dropdown", source: ["済", ""] },
        { data: "見積提出.value", type: "dropdown", source: ["済", ""] },
        { data: "請求書提出.value", type: "date", dateFormat: "YYYY-MM-DD" },
        { data: "状態.value" },
        { data: "材料証明書.value", type: "dropdown", source: ["不要", "必要"] },
        { data: "熱処理.value", type: "dropdown", source: ["不要", "必要", "必要（記録が必要）"] },
        {
          data: "表面処理.value",
          type: "dropdown",
          source: ["不要", "必要", "必要（記録が必要）"],
        },
        {
          data: "寸法記録.value",
          type: "dropdown",
          source: [
            "不要",
            "必要",
            "必要（レ点チェック）",
            "必要（実測値記入）",
            "必要（レ点チェック、実測値記入）",
          ],
        },
        { data: "案件管理レコード番号.value" },
        { data: "レコード番号.value" },
      ],

      // スプレットシート上のレコードを削除したときに呼び出されるイベント
      // 引数indexは削除する行
      // 引数amountは削除する行数
      beforeRemoveRow: (index, amount) => {
        // kintoneのレコードを削除する
        deleteRecords(
          hot.getSourceData(),
          index,
          amount,
          (deleteRecordsResp) => {
            console.dir(deleteRecordsResp);
            getRecords((getRecordsResp) => {
              // 削除後、データを再読み込み
              hot.loadData(getRecordsResp.records);
            });
          },
          (resp) => {
            console.dir(resp);
          }
        );
      },

      // スプレットシート上のレコードが編集されたときに呼び出されるイベント
      afterChange: (change, source) => {
        console.log(source);

        // データ読み込み時はイベントを終了
        if (source === "loadData") {
          return;
        }

        // kintoneのレコードを更新、追加する
        saveRecords(
          hot.getSourceData(),
          change,
          (resp) => {
            console.dir(resp);
            getRecords(
              (saveRecordsResp) => {
                // 更新後、データを再読み込み
                hot.loadData(saveRecordsResp.records);
              },
              (getRecordResp) => {
                // レコード取得失敗時に呼び出される
                console.dir(getRecordResp);
              }
            );
          },
          (resp) => {
            // 更新・追加時に呼び出される
            console.dir(resp);
          }
        );

        // 変更が見積_単価や数量にあった場合、見積_金額を再計算
        change.forEach(([row, prop, oldVal, newVal]) => {
          if (prop === "見積_単価.value" || prop === "個数.value") {
            const unitPrice = parseFloat(hot.getDataAtRowProp(row, "見積_単価.value")) || 0;
            const quantity = parseFloat(hot.getDataAtRowProp(row, "個数.value")) || 0;
            const amount = unitPrice * quantity;
            hot.setDataAtRowProp(row, "見積_金額.value", amount);
          }

          // 変更が原価_単価や数量にあった場合、原価_総額を再計算
          if (prop === "原価_単価.value" || prop === "個数.value") {
            const costUnitPrice = parseFloat(hot.getDataAtRowProp(row, "原価_単価.value")) || 0;
            const costQuantity = parseFloat(hot.getDataAtRowProp(row, "個数.value")) || 0;
            const costAmount = costUnitPrice * costQuantity;
            hot.setDataAtRowProp(row, "原価_総額.value", costAmount);

            // 原価_単価の90%を10の位で切り捨てて見積_単価に入れる
            const estimatedUnitPrice = Math.floor(costUnitPrice / 0.9 / 10) * 10;
            hot.setDataAtRowProp(row, "見積_単価.value", estimatedUnitPrice);
          }
        });
      },
    });

    // レコードを取得してhandsontableに反映
    getRecords((resp) => {
      hot.loadData(resp.records);
      // autoload(); // この行を削除して自動更新を停止
    });
  });
})();
