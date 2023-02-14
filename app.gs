// variables
let dates = [];
let count = 0;
const safeCount = 40;
let clears = [];

function createMenu() {
  SpreadsheetApp.getUi().createMenu('Crawl').addItem('Dữ liệu chứng khoán', 'loadFormInput').addToUi();
}

function loadFormInput() {
  const width = 640;
  const height = 480;
  const html = HtmlService.createHtmlOutputFromFile('index');
  html.setWidth(width).setHeight(height);
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(html, "Nhập khoảng thời gian");
}

function onOpen() {
  createMenu();
}

function crawlData(startDate = '2022-02-03', endDate = '2023-02-08') {
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  if (startDate && endDate) {
    ws.appendRow([startDate, endDate, new Date().toTimeString().slice(0, 8)]);

    let startDateTime = new Date(startDate);
    let endDateTime = new Date(endDate);
    let flagDate = true;
    while (flagDate) {
      if (parseInt(startDateTime.toISOString().slice(0, 10).replaceAll('-', '')) > parseInt(endDateTime.toISOString().slice(0, 10).replaceAll('-', ''))) {
        flagDate = false;
      } else {
        const dateDay = endDateTime.getDate();
        const dateMonth = endDateTime.getMonth() + 1;
        const dateYear = endDateTime.getFullYear();
        endDateTime = new Date(endDateTime.setDate(endDateTime.getDate() - 1));
        dates.push(`${dateDay < 10 ? `0${dateDay}` : dateDay}.${dateMonth < 10 ? `0${dateMonth}` : dateMonth}.${dateYear}`);
      }
    }
    console.log('dates: ', dates.length);
    fetchData(dates);
  } else {
    ws.appendRow(["null", "null"]);
  }
}

function fetchData(dates) {
  try {
    if (dates?.length - count < safeCount) {
      let data = dates?.slice(count)?.map(date => getHoseData(date))?.filter(value => Object.keys(value).length > 0);
      count = dates?.length;
      writeData(data);
    } else {
      let data = dates?.slice(count, count + safeCount)?.map(date => getHoseData(date))?.filter(value => Object.keys(value).length > 0);
      count += safeCount;
      writeData(data);
    }
  } catch (error) {
    console.log(error.toString());
    Browser.msgBox("Timeout! Try again");
  }
}

function writeData(data) {
  console.log("data: ", data);
  const as = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = as.getSheetByName('template');
  let inputValues = {};
  data?.forEach(item => {
    Object.keys(item).forEach(function (key) {
      if (!inputValues[key]) {
        inputValues[key] = [];
      }
      inputValues[key].push(item[key]);
    });
  });
  console.log('length key: ', Object.keys(inputValues).length);

  Object.keys(inputValues).forEach(function (key, index) {
    if (as.getSheetByName(key)) {
      let currentSheet = as.getSheetByName(key);
      const isClear = clears.find(token => token === key);
      if (!isClear) {
        currentSheet.getRange(4, 1, 1000, 13).clear();
        clears.push(key);
      }
      currentSheet.getRange(currentSheet.getLastRow() + 1, 1, inputValues[key]?.length, 13).setValues(inputValues[key]);
    } else {
      as.insertSheet(key, { template: templateSheet });
      let currentSheet = as.getSheetByName(key);
      currentSheet.getRange('A1').setValue(key);
      currentSheet.getRange(4, 1, inputValues[key]?.length, 13).setValues(inputValues[key]);
    }

    if (index === Object.keys(inputValues).length - 1) {
      console.log("Lấy dữ liệu hoàn tất. Đang ghi dữ liệu");
      // Browser.msgBox("Lấy dữ liệu hoàn tất. Đang ghi dữ liệu");
    }
  });
  if (count !== dates.length) {
    fetchData(dates);
  }
}

function getHoseData(date = '08.02.2023') {
  // console.log(date);
  let response = UrlFetchApp.fetch(`https://www.hsx.vn/Modules/Rsde/Report/QuoteReport?pageFieldName1=Date&pageFieldValue1=${date}&pageFieldOperator1=eq&pageFieldName2=KeyWord&pageFieldValue2=&pageFieldOperator2=&pageFieldName3=IndexType&pageFieldValue3=188803177&pageFieldOperator3=&pageCriteriaLength=3&_search=false&nd=1675827987067&rows=2147483647&page=1&sidx=id&sord=desc`, {
    "headers": {
      "accept": "application/json, text/javascript, */*; q=0.01",
      "accept-language": "vi-VN,vi;q=0.9,fr-FR;q=0.8,fr;q=0.7,en-US;q=0.6,en;q=0.5",
      "sec-ch-ua": "\"Not_A Brand\";v=\"99\", \"Google Chrome\";v=\"109\", \"Chromium\";v=\"109\"",
      "sec-ch-ua-mobile": "?0",
      "sec-ch-ua-platform": "\"Windows\"",
      "sec-fetch-dest": "empty",
      "sec-fetch-mode": "cors",
      "sec-fetch-site": "same-origin",
      "x-requested-with": "XMLHttpRequest",
      "cookie": "_ga=GA1.2.121148903.1671943941; ASP.NET_SessionId=pyb0vxhqjjwaiwgnujjj4ct2; TS016df111=01343ddb6a6af446a1f528ceaac0c0f54a9dd3175f6c0a201da872be6c39aac2398908fb39d7dc11790fbc7adaf618b47d5335fbadaed27e93114b3140bfc7cd7fef8c0e49; _gid=GA1.2.2134534016.1675921649; _gat_gtag_UA_116051872_2=1; TS0d710d04027=085cef26a9ab2000aaaef5591b165ba793e4422191d0c6bae8e03414a06b5ff3dc996ae331ed10ba08f89eb0a61130009af035fe9fb0c47bead2e7009f0be1972adf2df97dd65fcac4bad525d508684f252dfed68841ce5398412a757945e5f1",
      "Referer": "https://www.hsx.vn/Modules/Rsde/Report/Index?fid=a78ae7acb08d49ecae4b628c9ba98d26",
      "Referrer-Policy": "strict-origin-when-cross-origin"
    },
    "method": "GET"
  });
  console.log('res: ', JSON.parse(response.getResponseCode()), ' -- date: ', date);
  let data = JSON.parse(response.getContentText())?.rows;
  let result = {};
  if (data?.length > 0) {
    data?.forEach(function (item) {
      result[item?.id] = [date?.replaceAll('.', '/'), ...item?.cell?.splice(4)];
    });
  }
  // console.log('res: ', result);
  Utilities.sleep(1600);
  return result;
}
