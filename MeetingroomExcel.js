const { PDFDocument, rgb } = require("pdf-lib");
const fs = require("fs");
const Excel = require("exceljs");
const path = require("path");

async function extractFormFieldsFromPdfAndExportToExcel() {
  try {
    // 讀取本地端的 PDF 檔案
    const filePathFrom = process.cwd() + "/MeetingRoom.pdf";
    // const filePathFrom = "./outputPDF/MeetingRoom.pdf";
    const pdfBytes = fs.readFileSync(filePathFrom); // 創建 PDFDocument

    const pdfDoc = await PDFDocument.load(pdfBytes); // 取得填寫後的表單欄位內容

    const form = pdfDoc.getForm();

    // title 欄位名稱
    const title = [
      "申請單位",
      "身份別",
      "聯絡電話(手機)",
      "E-mail",
      "活動名稱",
      "活動人數",
      "活動內容",
      "活動日期",
      "申請時段",
      "已詳閱借用辦法",
      "終止借用，不得異議",
    ];
    // 申請單位
    const applicantUnitField = form.getTextField("applicantUnit").getText();

    // 身份別
    let identity = "";
    const checkBoxIdentity1 = form.getCheckBox("identity.one").isChecked();
    const checkBoxIdentity2 = form.getCheckBox("identity.two").isChecked();
    if (checkBoxIdentity1) {
      identity = "校友";
    } else if (checkBoxIdentity2) {
      identity = "校內單位";
    } else {
      identity = "無填寫";
    }
    // E-mail
    const emailField = form.getTextField("email").getText();
    // 聯絡電話
    const mobilePhoneField = form.getTextField("mobilePhone").getText();
    // 活動名稱
    const eventNameField = form.getTextField("eventName").getText();
    // 活動人數
    const eventAmountField = form.getTextField("eventAmount").getText();
    // 活動內容
    const eventContentField = form.getTextField("eventContent").getText();
    // 活動日期
    let eventDate = "";
    const applyStartYField = form.getTextField("applyStartY").getText();
    const applyStartMField = form.getTextField("applyStartM").getText();
    const applyStartDField = form.getTextField("applyStartD").getText();

    const applyEndYField = form.getTextField("applyEndY").getText();
    const applyEndMField = form.getTextField("applyEndM").getText();
    const applyEndDField = form.getTextField("applyEndD").getText();
    if (
      applyStartYField != undefined ||
      applyStartMField != undefined ||
      applyStartDField != undefined ||
      applyEndYField != undefined ||
      applyEndMField != undefined ||
      applyEndDField != undefined
    ) {
      eventDate = `${applyStartYField}年 ${applyStartMField}月 ${applyStartDField}日 ~ ${applyEndYField}年 ${applyEndMField}月 ${applyEndDField}日`;
    } else {
      eventDate = "無填寫";
    }
    // 申請時段
    let eventTime = "";
    const checkBoxT1 = form.getCheckBox("card.T1").isChecked();
    const checkBoxT2 = form.getCheckBox("card.T2").isChecked();
    const checkBoxT3 = form.getCheckBox("card.T3").isChecked();
    if (checkBoxT1) {
      eventTime = "上午09至12時";
    } else if (checkBoxT2) {
      eventTime = "中午12至14時";
    } else if (checkBoxT3) {
      eventTime = "下午14至17時";
    } else {
      eventTime = "無填寫";
    }
    // 已詳閱借用辦法
    let read = "X";
    const checkBoxRead = form.getCheckBox("read").isChecked();
    if (checkBoxRead) {
      read = "V";
    }
    // 終止借用，不得異議
    let noOther = "X";
    const checkBoxNo = form.getCheckBox("no").isChecked();
    if (checkBoxNo) {
      read = "V";
    }

    // 要檢查的 Excel 檔案路徑
    const filePath = process.cwd() + "/MeetingRoom.xlsx";
    // const filePath = "MeetingRoom.xlsx";
    fs.access(filePath, fs.constants.F_OK, (err) => {
      // 建立Excel work
      const workbook = new Excel.Workbook();
      if (err) {
        console.log("檔案不存在");

        // 第一個分頁
        const worksheet1 = workbook.addWorksheet("臨時卡"); // 將填寫後的文字內容寫入 Excel
        worksheet1.addRow(title);
        addToWorkBook(worksheet1);
        // 儲存更新後的 Excel 檔案
        return workbook.xlsx.writeFile(filePath);
      } else {
        console.log("檔案存在");
        workbook.xlsx
          .readFile(filePath)
          .then(() => {
            const worksheet1 = workbook.getWorksheet(1); // 取得第一個工作表

            // 將資料陣列加入新的 Excel 行
            addToWorkBook(worksheet1);

            // 儲存更新後的 Excel 檔案
            return workbook.xlsx.writeFile(filePath);
          })
          .then(() => {
            console.log(filePath);
            console.log("新資料已添加到 Excel 檔案");
          })
          .catch((error) => {
            console.error("發生錯誤：", error);
          });
      }
    });
    // 加入
    function addToWorkBook(worksheet1) {
      worksheet1.addRow([
        applicantUnitField,
        identity,
        mobilePhoneField,
        emailField,
        eventNameField,
        eventAmountField,
        eventContentField,
        eventDate,
        eventTime,
        read,
        noOther,
      ]);
    }
    // const excelFilePath = "output.xlsx";
    // await workbook.xlsx.writeFile(excelFilePath);

    console.log("填寫內容已匯出到 Excel 檔案！");
  } catch (err) {
    console.error("發生錯誤：", err);
  }
}

extractFormFieldsFromPdfAndExportToExcel();
