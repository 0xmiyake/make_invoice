
// 請求書出力
kintone.events.on('app.record.detail.show', (event) => {
    // メニューボタンが配置されている要素を取得
    const headerMenuSpace = kintone.app.record.getHeaderMenuSpaceElement();

    // 新しい帳票出力ボタンの作成
    const reportButton = document.createElement('button');
    reportButton.id = 'report_button';
    reportButton.innerText = '帳票出力';

    // メニューエリアにボタンを追加
    headerMenuSpace.appendChild(reportButton); // メニューボタンの隣に追加

    // 帳票出力ボタンを押したときのイベント
    reportButton.addEventListener('click', async function () {
        new kintone.Promise(function (resolve) {
            // GETリクエストパラメータの設定
            const body = {
                app: 250, // 帳票フォーマットアプリid（それぞれ違うので変えてください）
                query: 'レコード番号 = "1"', // フォーマットを保存しているレコード番号でデータをGET
            };
            resolve(body);
        })
            .then(getData) // 帳票フォーマットアプリからレコード番号1をget
            .then(dlFile) // getしたデータにあるfileKeyをつかって、ファイルダウンロードAPIを実行してレスポンスをもらう
            .then(writeXlsx.bind(null, event)) // excelに書き込む。レコードのデータeventを第1引数として渡す
            .then(saveLocal); // レスポンスをローカルにセーブする（自分のPCにダウンロードされる）
    });
});


function getData(requestParam) {
    return new kintone.Promise(function (resolve, reject) {
        kintone.api(
            kintone.api.url("/k/v1/records", true),
            "GET",
            requestParam,
            function (resp) {
                resolve(resp); // getしたデータを次のdlFile関数に渡します
            },
            function (err) {
                console.log(err);
            }
        );
    });
}

// getData関数からデータをrespで受け取ります。
function dlFile(resp) {
    return new kintone.Promise(function (resolve, reject) {
        // fileKeyを取り出します
        const filekey = resp.records[0].make_invoice_file.value[0].fileKey;
        // fileKeyをurlに設定します。
        const url = kintone.api.urlForGet("/k/v1/file", { fileKey: filekey }, true);
        // ファイルダウンロードAPI を実行します。
        const xhr = new XMLHttpRequest();
        xhr.open("GET", url);
        xhr.setRequestHeader("X-Requested-With", "XMLHttpRequest");
        xhr.responseType = "blob";

        xhr.onload = function () {
            if (xhr.status === 200) {
                // エクセルのフォーマットがレスポンスとして返却されます。
                console.log(xhr.response);
                resolve(xhr.response); // エクセルのフォーマットを次のwriteXlsx関数にわたす。
            }
        };
        xhr.send();
    });
}

// レコードのデータをeventで、エクセルのフォーマットをrespで受け取る
function writeXlsx(event, resp) {
    return new kintone.Promise(async function (resolve, reject) {
        try {
            const record = event.record;
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(resp);
            const worksheet = workbook.worksheets[0];

            // ヘッダー情報の入力
            worksheet.getCell("A5").value = record.company_name?.value || ''; // company_nameがundefinedの場合空に
            worksheet.getCell("D6").value = record.company_person?.value || '';
            worksheet.getCell("N2").value = record.invoice_date?.value || '';
            worksheet.getCell("M14").value = record.payment_deadline?.value || '';
            worksheet.getCell("D14").value = record.total_amount?.value || '';
            worksheet.getCell("C35").value = record.notes?.value || '';
            worksheet.getCell("L29").value = record.invoice_total?.value || '';
            worksheet.getCell("L30").value = record.invoice_tax?.value || '';
            worksheet.getCell("L31").value = record.total_amount?.value || '';

            // 複数の明細を `invoice_table` から取得し、各行に書き込む
            const invoiceItems = record.invoice_table?.value; // invoice_tableがundefinedの場合に備える
            if (invoiceItems && invoiceItems.length > 0) {
                let startRow = 17; // 開始行番号

                // `invoice_table` の各レコードをループ
                for (let i = 0; i < invoiceItems.length; i++) {
                    const invoiceRow = invoiceItems[i]?.value;

                    worksheet.getCell(`A${startRow}`).value = i + 1; // 行番号を1からスタートして記入
                    worksheet.getCell(`B${startRow}`).value = invoiceRow?.invoice_item?.value || ''; // 商品名
                    worksheet.getCell(`J${startRow}`).value = invoiceRow?.invoice_quantity?.value || ''; // 数量
                    worksheet.getCell(`K${startRow}`).value = invoiceRow?.invoice_unit?.value || ''; // 単位
                    worksheet.getCell(`L${startRow}`).value = invoiceRow?.invoice_unit_price?.value || ''; // 単価
                    // worksheet.getCell(`O${startRow}`).value = invoiceRow?.invoice_subtotal?.value || ''; // 小計

                    startRow++; // 次の行へ
                }
            } else {
                console.error("invoice_table が空です、または存在しません。");
            }

            // UInt8Arrayを生成
            const uint8Array = await workbook.xlsx.writeBuffer();
            resolve(uint8Array); // 次のsaveLocal関数に渡す
        } catch (error) {
            console.error("エラーが発生しました: ", error);
            reject(error);
        }
    });
}

function saveLocal(uint8Array) {
    return new kintone.Promise(function (resolve, reject) {
        // Blobオブジェクトにファイルを格納
        const blob = new Blob([uint8Array], {
            type: "application/octet-binary",
        });
        const url = window.URL || window.webkitURL;

        // BlobURLの取得
        const blobUrl = url.createObjectURL(blob);

        // リンクを作成し、そこにBlobオブジェクトを設定する
        const alink = document.createElement("a");
        alink.textContent = "ダウンロード";
        alink.download = "seikyuusyo.xlsx";
        alink.href = blobUrl;
        alink.target = "_blank";

        // マウスイベントを設定
        const e = new MouseEvent("click", {
            view: window,
            bubbles: true,
            cancelable: true,
        });

        // aタグのクリックイベントをディスパッチする
        alink.dispatchEvent(e);
    });
}

