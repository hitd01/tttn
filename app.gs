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

function crawlData(startDate = '2022-09-01', endDate = new Date().toISOString().slice(0, 10)) {
  try {
    Browser.msgBox("Đang thực hiện lấy dữ liệu. Vui lòng không thực hiện lại. Nhấn OK để tiếp tục!");
    const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
    if (startDate && endDate) {
      ws.appendRow([startDate, endDate, new Date().toTimeString().slice(0, 8)]);

      let startDateTime = new Date(startDate);
      let endDateTime = new Date(endDate);
      let flagDate = true;
      let dates = [];
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

      var userProperties = PropertiesService.getUserProperties();
      userProperties.setProperty('DATES', JSON.stringify(dates));
      userProperties.setProperty('CLEARS', JSON.stringify([]));
      userProperties.setProperty('IS_WRITING', 'false');
      userProperties.setProperty('INDEX_WRITING', '0');
      if (!userProperties.getProperty('TICKERS')) {
        userProperties.setProperty('TICKERS', JSON.stringify([]));
      }
      if (!userProperties.getProperty('TICKER_COUNT')) {
        userProperties.setProperty('TICKER_COUNT', '0');
      }

      fetchData();
    }
  }
  catch (error) {
    console.log(error.message);
  }
}

function fetchData() {
  try {
    const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
    ws.appendRow(["run fetch", new Date().toTimeString().slice(0, 8)]);
    startOrResumeContinousExecutionInstance("fetchData");
    var userProperties = PropertiesService.getUserProperties();

    let isWriting = userProperties.getProperty('IS_WRITING');

    if (isWriting === 'true') {
      writeData();
    } else {
      if (getBatchKey("fetchData") === "") setBatchKey("fetchData", 0);

      var counter = Number(getBatchKey("fetchData"));
      var dates = JSON.parse(userProperties.getProperty('DATES'));
      console.log("dates fetch: ", dates.length);
      let data = JSON.parse(userProperties.getProperty('DATA'))?.length > 0 ? JSON.parse(userProperties.getProperty('DATA')) : [];
      for (let i = 0 + counter; i < dates?.length; i++) {
        if (isTimeRunningOut("fetchData")) {
          return;
        }

        const hoseData = getHoseData(dates[i]);
        if (hoseData) {
          if (Object.keys(hoseData).length > 0) {
            data = [...data, hoseData];
            userProperties.setProperty('DATA', JSON.stringify(data));
          }
          setBatchKey("fetchData", i);
        } else {
          i--;
        }
      }

      if (Number(getBatchKey("fetchData")) === dates?.length - 1) {
        writeData();
      }
    }
  } catch (error) {
    console.log(error.message);
    if (error.message?.includes('exceeded the property storage quota')) {
      writeData();
    }
  }
}

function writeData() {
  try {
    // console.log("data: ", data);
    var userProperties = PropertiesService.getUserProperties();
    let tickers = JSON.parse(userProperties.getProperty('TICKERS'));
    let tickerCount = Number(userProperties.getProperty('TICKER_COUNT'));
    const data = JSON.parse(userProperties.getProperty('DATA'));
    const as = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = as.getSheetByName('Template');

    console.log("data length: ", data?.length);
    console.log("tickers: ", tickers);
    console.log("ticker count: ", tickerCount);

    let inputValues = {};
    data?.forEach(item => {
      Object.keys(item).forEach(function (key) {
        if (!inputValues[key]) {
          inputValues[key] = [];
        }
        inputValues[key].push(item[key]);
      });
    });
    console.log('length key length: ', Object.keys(inputValues).length);

    let clears = JSON.parse(userProperties.getProperty('CLEARS'));
    let indexWriting = Number(userProperties.getProperty('INDEX_WRITING'));

    let dataWrite = Object.keys(inputValues).slice(indexWriting);
    console.log("data write: ", dataWrite);

    for (let i = 0; i < dataWrite.length; i++) {
      if (isTimeRunningOut("fetchData")) {
        userProperties.setProperty('IS_WRITING', 'true');
        userProperties.setProperty('INDEX_WRITING', `${indexWriting + i}`);
        console.log(indexWriting + i, indexWriting, i);
        return;
      }

      if (tickerCount > 0) {
        let ticker = tickers?.find(ticker => ticker[dataWrite[i]]);
        if (ticker) {
          if (as.getSheetByName(ticker[dataWrite[i]])) {
            let currentSheet = as.getSheetByName(ticker[dataWrite[i]]);
            const isClear = clears?.find(token => token === dataWrite[i]);
            if (!isClear) {
              currentSheet.getRange(4, 1, 1000, 13).clear();
              clears.push(dataWrite[i]);
              userProperties.setProperty('CLEARS', JSON.stringify(clears));
            }
            currentSheet.getRange(currentSheet.getLastRow() + 1, 1, inputValues[dataWrite[i]]?.length, 13).setValues(inputValues[dataWrite[i]]);
          } else {
            as.insertSheet(ticker[dataWrite[i]], { template: templateSheet });
            let currentSheet = as.getSheetByName(ticker[dataWrite[i]]);
            currentSheet.getRange('A1').setValue(dataWrite[i]);
            currentSheet.getRange(4, 1, inputValues[dataWrite[i]]?.length, 13).setValues(inputValues[dataWrite[i]]);

            clears.push(dataWrite[i]);
            userProperties.setProperty('CLEARS', JSON.stringify(clears));
          }
        } else {
          as.insertSheet(`${tickerCount + 1}`, { template: templateSheet });
          let currentSheet = as.getSheetByName(`${tickerCount + 1}`);
          currentSheet.getRange('A1').setValue(dataWrite[i]);
          currentSheet.getRange(4, 1, inputValues[dataWrite[i]]?.length, 13).setValues(inputValues[dataWrite[i]]);

          let newTicker = {};
          newTicker[dataWrite[i]] = `${tickerCount + 1}`;
          tickers?.push(newTicker);
          userProperties.setProperty('TICKERS', JSON.stringify(tickers));
          userProperties.setProperty('TICKER_COUNT', `${tickerCount + 1}`);
          tickerCount++;

          clears.push(dataWrite[i]);
          userProperties.setProperty('CLEARS', JSON.stringify(clears));
        }
      } else {
        as.insertSheet('1', { template: templateSheet });
        let currentSheet = as.getSheetByName('1');
        currentSheet.getRange('A1').setValue(dataWrite[i]);
        currentSheet.getRange(4, 1, inputValues[dataWrite[i]]?.length, 13).setValues(inputValues[dataWrite[i]]);

        let newTicker = {};
        newTicker[dataWrite[i]] = `1`;
        tickers?.push(newTicker);
        userProperties.setProperty('TICKERS', JSON.stringify(tickers));
        userProperties.setProperty('TICKER_COUNT', '1');
        tickerCount++;

        clears.push(dataWrite[i]);
        userProperties.setProperty('CLEARS', JSON.stringify(clears));
      }

      if (i === dataWrite.length - 1) {
        console.log("Lấy dữ liệu hoàn tất. Đang ghi dữ liệu");
        userProperties.setProperty('IS_WRITING', 'false');
        userProperties.setProperty('INDEX_WRITING', '0');
      }
    }

    let isWriting = userProperties.getProperty('IS_WRITING');
    if (isWriting === 'false') {
      var counter = Number(getBatchKey("fetchData"));
      var dates = JSON.parse(userProperties.getProperty('DATES'));
      if (counter === dates?.length - 1) {
        endContinuousExecutionInstance("fetchData");
      } else {
        userProperties.deleteProperty('DATA');
        fetchData();
      }
    }
  }
  catch (error) {
    console.log(error.message);
  }
}

function getHoseData(date = '08.02.2023') {
  try {
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
    // console.log('data fetch: ', data);
    let result = {};
    if (data?.length > 0) {
      data?.forEach(function (item) {
        result[item?.id] = [date?.replaceAll('.', '/'), ...item?.cell?.splice(4)];
      });
    }
    // console.log('result: ', result);
    // Utilities.sleep(1600);
    return result;
  }
  catch (error) {
    console.log("fetch error: ", error.message);
    console.log('res error -- date: ', date);
    if (error.message?.includes('Timeout')) {
      getHoseData(date);
    }
  }
}

/**
 *  ---  Blog: https://gist.github.com/patt0/8395003  ---
 */

function startOrResumeContinousExecutionInstance(fname = "fetchData") {
  console.log("start: ", fname);
  var userProperties = PropertiesService.getUserProperties();
  var start = userProperties.getProperty('GASCBL_' + fname + '_START_BATCH');
  if (start === "" || start === null) {
    start = new Date();
    userProperties.setProperty('GASCBL_' + fname + '_START_BATCH', start);
    userProperties.setProperty('GASCBL_' + fname + '_KEY', "");
  }

  userProperties.setProperty('GASCBL_' + fname + '_START_ITERATION', new Date());

  deleteCurrentTrigger_(fname);
  enableNextTrigger_(fname);
}

function setBatchKey(fname, key) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('GASCBL_' + fname + '_KEY', key);
}

function getBatchKey(fname) {
  var userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty('GASCBL_' + fname + '_KEY');
}

function endContinuousExecutionInstance(fname = "fetchData", emailRecipient = 'huanpanda972001@gmail.com', customTitle = 'DONE') {
  console.log("end: ", fname);
  var userProperties = PropertiesService.getUserProperties();

  var end = new Date();
  var start = userProperties.getProperty('GASCBL_' + fname + '_START_BATCH');
  var key = userProperties.getProperty('GASCBL_' + fname + '_KEY');
  var emailTitle = customTitle + " : Continuous Execution Script for " + fname;
  var body = "Started: " + start + "<br>" + "Ended: " + end + "<br>" + "LAST KEY : " + key;
  MailApp.sendEmail(emailRecipient, emailTitle, "", { htmlBody: body });

  deleteCurrentTrigger_(fname);
  userProperties.deleteProperty('GASCBL_' + fname + '_START_ITERATION');
  userProperties.deleteProperty('GASCBL_' + fname + '_START_BATCH');
  userProperties.deleteProperty('GASCBL_' + fname + '_KEY');
  userProperties.deleteProperty('GASCBL_' + fname);
  userProperties.deleteProperty('DATES');
  userProperties.deleteProperty('DATA');
  userProperties.deleteProperty('CLEARS');
}

function isTimeRunningOut(fname) {
  var userProperties = PropertiesService.getUserProperties();
  var start = new Date(userProperties.getProperty('GASCBL_' + fname + '_START_ITERATION'));
  var now = new Date();

  var timeElapsed = Math.floor((now.getTime() - start.getTime()) / 1000);
  return (timeElapsed > 300);
}

function enableNextTrigger_(fname) {
  var userProperties = PropertiesService.getUserProperties();
  var nextTrigger = ScriptApp.newTrigger(fname).timeBased().after(6 * 60 * 1000 + 3000).create();
  var triggerId = nextTrigger.getUniqueId();

  userProperties.setProperty('GASCBL_' + fname, triggerId);
}

function deleteCurrentTrigger_(fname) {
  var userProperties = PropertiesService.getUserProperties();
  var triggerId = userProperties.getProperty('GASCBL_' + fname);
  var triggers = ScriptApp.getProjectTriggers();
  for (var i in triggers) {
    if (triggers[i].getUniqueId() === triggerId)
      ScriptApp.deleteTrigger(triggers[i]);
  }
  userProperties.setProperty('GASCBL_' + fname, "");
}

function debugFunc() {
  try {
    const scriptProperties = PropertiesService.getUserProperties();

    // scriptProperties.deleteAllProperties();

    const data = scriptProperties.getProperties();
    for (const key in data) {
      console.log('Key: %s, Value: %s', key, data[key]);
    }

    // var userProperties = PropertiesService.getUserProperties();
    // var triggerId = userProperties.getProperty('GASCBL_fetchData');
    // console.log("triggerId: ", triggerId);
    // var triggers = ScriptApp.getProjectTriggers();
    // for (var i in triggers) {
    //   if (triggers[i].getUniqueId() === triggerId) {
    //     ScriptApp.deleteTrigger(triggers[i]);
    //     console.log('deleted');
    //   }
    // }
  } catch (err) {
    // TODO (developer) - Handle exception
    console.log('Failed with error %s', err.message);
  }
}
