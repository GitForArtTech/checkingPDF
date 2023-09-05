const {
  PDFDocument,
  rgb,
  PDFField,
  PDFSignature,
  PDFForm,
} = require("pdf-lib");
const fs = require("fs");
const fetch = require("node-fetch");
const fontkit = require("pdf-fontkit");
const donatePDF = "./originalPDF/ActivityHall.pdf";

async function createPdfWithChineseText() {
  // load PDF
  const donatePDFBytes = fs.readFileSync(donatePDF);
  const pdfDoc = await PDFDocument.load(donatePDFBytes);
  pdfDoc.registerFontkit(fontkit);

  // 取得第一個頁面
  const pages = pdfDoc.getPages();
  const page = pages[0];

  // 選擇一個支持中文的字型（請確保字型文件存在並且指定正確的路徑）
  const fontBytes = fs.readFileSync("./jf.ttf");
  const font = await pdfDoc.embedFont(fontBytes);
  const processBorderSize = (x, y, w, h) => {
    let retObject = {
      x: x,
      y: y,
      width: w,
      height: h,
      //   textColor: rgb(0, 0, 0),
      borderColor: rgb(0, 0, 1),
      //   backgroundColor:rgb(1, 1, 1),
      borderWidth: 1,
      font: font,
      fontSize: 12,
      isPassword: false,
      isReadOnly: false,
      isVisible: true, // 設為可見狀態
    };
    return retObject;
  };

  // 建立表格
  const form = pdfDoc.getForm();
  // 申請日期
  const fillFormYearField = form.createTextField("fillFormYear");
  fillFormYearField.addToPage(page, processBorderSize(410, 770, 38, 15));
  const fillFormMonthField = form.createTextField("fillFormMonth");
  fillFormMonthField.addToPage(page, processBorderSize(465, 770, 22, 15));
  const fillFormDayField = form.createTextField("fillFormDay");
  fillFormDayField.addToPage(page, processBorderSize(502, 770, 23, 15));

  // 申請單位
  const applicantUnitField = form.createTextField("applicantUnit");
  applicantUnitField.addToPage(page, processBorderSize(137, 735, 165, 15));
  // 申請人
  const applicantPersonField = form.createTextField("applicantPerson");
  applicantPersonField.addToPage(page, processBorderSize(378, 735, 180, 15));
  // 聯絡電話/手機
  const telPhoneField = form.createTextField("telPhone");
  telPhoneField.addToPage(page, processBorderSize(137, 693, 165, 15));
  // email
  const emailField = form.createTextField("email");
  emailField.addToPage(page, processBorderSize(378, 693, 180, 15));
  // 活動名稱
  const eventNameField = form.createTextField("eventName");
  eventNameField.addToPage(page, processBorderSize(137, 651, 165, 15));
  // 活動人數
  const eventAmountField = form.createTextField("eventAmount");
  eventAmountField.addToPage(page, processBorderSize(378, 651, 60, 15));
  // 借用場地
  // 27號
  const checkBox27 = form.createCheckBox("site.27");
  checkBox27.addToPage(page, {
    x: 140,
    y: 614,
    width: 10,
    height: 10,
  });
  const checkBox271 = form.createCheckBox("site.271");
  checkBox271.addToPage(page, {
    x: 190,
    y: 614,
    width: 10,
    height: 10,
  });
  const checkBox272 = form.createCheckBox("site.272");
  checkBox272.addToPage(page, {
    x: 345,
    y: 614,
    width: 10,
    height: 10,
  });
  // 28號
  const checkBox28 = form.createCheckBox("site.28");
  checkBox28.addToPage(page, {
    x: 140,
    y: 593,
    width: 10,
    height: 10,
  });
  const checkBox281 = form.createCheckBox("site.281");
  checkBox281.addToPage(page, {
    x: 190,
    y: 593,
    width: 10,
    height: 10,
  });
  const checkBox282 = form.createCheckBox("site.282");
  checkBox282.addToPage(page, {
    x: 345,
    y: 593,
    width: 10,
    height: 10,
  });
  // 29號
  const checkBox29 = form.createCheckBox("site.29");
  checkBox29.addToPage(page, {
    x: 140,
    y: 575,
    width: 10,
    height: 10,
  });
  const checkBox291 = form.createCheckBox("site.291");
  checkBox291.addToPage(page, {
    x: 238,
    y: 575,
    width: 10,
    height: 10,
  });
  const checkBox292 = form.createCheckBox("site.292");
  checkBox292.addToPage(page, {
    x: 345,
    y: 575,
    width: 10,
    height: 10,
  });
  // 借用時間
  const applyStartYField = form.createTextField("applyStartY");
  applyStartYField.addToPage(page, processBorderSize(140, 540, 38, 15));
  const applyStartMField = form.createTextField("applyStartM");
  applyStartMField.addToPage(page, processBorderSize(190, 540, 40, 15));
  const applyStartDField = form.createTextField("applyStartD");
  applyStartDField.addToPage(page, processBorderSize(243, 540, 40, 15));
  // 週 與 時
  const applyStartWeekField = form.createTextField("applyStartW");
  applyStartWeekField.addToPage(page, processBorderSize(320, 540, 23, 15));
  const applyStartTimeField = form.createTextField("applyStartT");
  applyStartTimeField.addToPage(page, processBorderSize(355, 540, 38, 15));

  const applyEndYField = form.createTextField("applyEndY");
  applyEndYField.addToPage(page, processBorderSize(140, 500, 38, 15));
  const applyEndMField = form.createTextField("applyEndM");
  applyEndMField.addToPage(page, processBorderSize(190, 500, 40, 15));
  const applyEndDField = form.createTextField("applyEndD");
  applyEndDField.addToPage(page, processBorderSize(243, 500, 40, 15));

  // 週 與 時
  const applyEndWeekField = form.createTextField("applyEndW");
  applyEndWeekField.addToPage(page, processBorderSize(320, 500, 23, 15));
  const applyEndTimeField = form.createTextField("applyEndT");
  applyEndTimeField.addToPage(page, processBorderSize(355, 500, 38, 15));

  // 共計小時數
  const totalHourField = form.createTextField("totalHour");
  totalHourField.addToPage(page, processBorderSize(450, 516, 70, 15));

  // 注意事項
  // 詳閱
  const checkBoxRead = form.createCheckBox("read");
  checkBoxRead.addToPage(page, {
    x: 140,
    y: 473,
    width: 10,
    height: 10,
  });
  // 不得異議
  const checkBoxNo = form.createCheckBox("no");
  checkBoxNo.addToPage(page, {
    x: 140,
    y: 418,
    width: 10,
    height: 10,
  });
  // 活動是否含餐飲或外燴
  const checkBoxEatY = form.createCheckBox("EatY");
  checkBoxEatY.addToPage(page, {
    x: 140,
    y: 360,
    width: 10,
    height: 10,
  });
  const checkBoxEatN = form.createCheckBox("EatN");
  checkBoxEatN.addToPage(page, {
    x: 140,
    y: 340,
    width: 10,
    height: 10,
  });
  // 借用設備
  const checkBoxRent1 = form.createCheckBox("Rent1");
  checkBoxRent1.addToPage(page, {
    x: 334,
    y: 372,
    width: 10,
    height: 10,
  });
  const checkBoxRent2 = form.createCheckBox("Rent2");
  checkBoxRent2.addToPage(page, {
    x: 334,
    y: 350,
    width: 10,
    height: 10,
  });
  // 其他原因
  const reason1Field = form.createTextField("reason1");
  reason1Field.addToPage(page, processBorderSize(370, 350, 185, 15));
  const reason2Field = form.createTextField("reason2");
  reason2Field.addToPage(page, processBorderSize(335, 327, 220, 15));

  // 秘書處審核
  const checkBoxReviewY = form.createCheckBox("ReviewY");
  checkBoxReviewY.addToPage(page, {
    x: 120,
    y: 208,
    width: 12,
    height: 12,
  });
  const checkBoxReviewN = form.createCheckBox("ReviewN");
  checkBoxReviewN.addToPage(page, {
    x: 120,
    y: 195,
    width: 12,
    height: 12,
  });
  const checkBoxReviewN1 = form.createCheckBox("ReviewN1");
  checkBoxReviewN1.addToPage(page, {
    x: 123,
    y: 181,
    width: 11,
    height: 11,
  });
  const checkBoxReviewN2 = form.createCheckBox("ReviewN2");
  checkBoxReviewN2.addToPage(page, {
    x: 267,
    y: 181,
    width: 11,
    height: 11,
  });
  const checkBoxReviewN3 = form.createCheckBox("ReviewN3");
  checkBoxReviewN3.addToPage(page, {
    x: 398,
    y: 181,
    width: 11,
    height: 11,
  });

  // 保證金
  const marginField = form.createTextField("margin");
  marginField.addToPage(page, processBorderSize(145, 130, 60, 15));
  // 場地管理員
  const checkBoxManagerY = form.createCheckBox("ManagerY");
  checkBoxManagerY.addToPage(page, {
    x: 120,
    y: 110,
    width: 12,
    height: 12,
  });
  const checkBoxManagerN = form.createCheckBox("ManagerN");
  checkBoxManagerN.addToPage(page, {
    x: 204,
    y: 110,
    width: 12,
    height: 12,
  });

  // 將PDF文件保存到磁碟
  const pdfBytes = await pdfDoc.save();
  fs.writeFileSync("./outputPDF/ActivityHall.pdf", pdfBytes);
}

createPdfWithChineseText();
