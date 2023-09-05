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
const donatePDF = "./originalPDF/MeetingRoom.pdf";

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
      borderColor: rgb(1, 1, 1),
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
  // 申請單位
  const applicantUnitField = form.createTextField("applicantUnit");
  applicantUnitField.addToPage(page, processBorderSize(150, 685, 70, 15));
  // 身份別
  const checkBoxIdentity1 = form.createCheckBox("identity.one");
  checkBoxIdentity1.addToPage(page, {
    x: 345,
    y: 688,
    width: 10,
    height: 10,
  });
  const checkBoxIdentity2 = form.createCheckBox("identity.two");
  checkBoxIdentity2.addToPage(page, {
    x: 388,
    y: 688,
    width: 10,
    height: 10,
  });
  // Email
  const emailField = form.createTextField("email");
  emailField.addToPage(page, processBorderSize(300, 645, 195, 15));
  // 聯絡電話（手機）
  const mobilePhoneField = form.createTextField("mobilePhone");
  mobilePhoneField.addToPage(page, processBorderSize(150, 645, 70, 15));
  // 活動名稱
  const eventNameField = form.createTextField("eventName");
  eventNameField.addToPage(page, processBorderSize(150, 588, 70, 15));

  // 活動人數
  const eventAmountField = form.createTextField("eventAmount");
  eventAmountField.addToPage(page, processBorderSize(300, 588, 90, 15));

  // 活動內容
  const eventContentField = form.createTextField("eventContent");
  eventContentField.addToPage(page, processBorderSize(150, 525, 345, 15));

  // 活動日期
  const applyStartYField = form.createTextField("applyStartY");
  applyStartYField.addToPage(page, processBorderSize(152, 458, 20, 15)); // 135
  const applyStartMField = form.createTextField("applyStartM");
  applyStartMField.addToPage(page, processBorderSize(178, 458, 18, 15));
  const applyStartDField = form.createTextField("applyStartD");
  applyStartDField.addToPage(page, processBorderSize(202, 458, 18, 15));

  const applyEndYField = form.createTextField("applyEndY");
  applyEndYField.addToPage(page, processBorderSize(152, 438, 20, 15));
  const applyEndMField = form.createTextField("applyEndM");
  applyEndMField.addToPage(page, processBorderSize(178, 438, 18, 15));
  const applyEndDField = form.createTextField("applyEndD");
  applyEndDField.addToPage(page, processBorderSize(202, 438, 18, 15));

  // 申請時段
  const checkBoxT1 = form.createCheckBox("card.T1");
  checkBoxT1.addToPage(page, {
    x: 354,
    y: 470,
    width: 10,
    height: 10,
  });
  const checkBoxT2 = form.createCheckBox("card.T2");
  checkBoxT2.addToPage(page, {
    x: 354,
    y: 450,
    width: 10,
    height: 10,
  });
  const checkBoxT3 = form.createCheckBox("card.T3");
  checkBoxT3.addToPage(page, {
    x: 354,
    y: 430,
    width: 10,
    height: 10,
  });

  // 詳閱
  const checkBoxRead = form.createCheckBox("read");
  checkBoxRead.addToPage(page, {
    x: 88,
    y: 390,
    width: 10,
    height: 10,
  });
  // 不得異議
  const checkBoxNo = form.createCheckBox("no");
  checkBoxNo.addToPage(page, {
    x: 88,
    y: 330,
    width: 10,
    height: 10,
  });

  // 將PDF文件保存到磁碟
  const pdfBytes = await pdfDoc.save();
  fs.writeFileSync("./outputPDF/MeetingRoom.pdf", pdfBytes);
}

createPdfWithChineseText();
