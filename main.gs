const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheetResponses = ss.getSheetByName("Form Responses 1");
const sheetSettings = ss.getSheetByName("Settings");
const answersSettingsRange_F1 = sheetSettings.getRange("G2");


function dataManipulationAlgorithm() {
  let answersRange_F1 = answersSettingsRange_F1.getValue().toString().split('-');
  let lr = sheetResponses.getLastRow();
  let multiplicator = sheetSettings.getRange("J2").getValue();
  let answersScore_Arr = sheetSettings.getRange("K2:L").getValues().filter(item => item[0] != '');
  let amswersScore_List = new Map(answersScore_Arr);
  let numberOfAnswerToMultiplicate = sheetSettings.getRange("J3").getValue().toString().split(',');
  let answersResults_F1 = sheetResponses.getRange(getStringRange(answersSettingsRange_F1, lr)).getValues()[0];
  Logger.log(answersResults_F1.length)
  let sum_F1 = 0;
  let reg = /.*\(/gi;
  for (let i = 0; i < answersResults_F1.length; i++) {
    let answersResults_F1_SearchPhrase = answersResults_F1[i].toString().match(reg)[0];
    answersResults_F1_SearchPhrase = answersResults_F1_SearchPhrase.substring(0, answersResults_F1_SearchPhrase.length - 2);
    Logger.log(numberOfAnswerToMultiplicate)
    if (numberOfAnswerToMultiplicate.find(item => item == i + 1) != undefined) {
      Logger.log("Сработает мультипликатор!   " + numberOfAnswerToMultiplicate.find(item => item == i + 1))
      sum_F1 += amswersScore_List.get(answersResults_F1_SearchPhrase) * multiplicator;
    } else {
      sum_F1 += amswersScore_List.get(answersResults_F1_SearchPhrase);
    }
  Logger.log(sum_F1)
  }


}


/**
 * Функция возвращает строку с указание диапазона в котором ищем ответы 
 * @param {string} answersSettingsRange_F1 диапазон с строкой диапазона поиска ответов (Пример: "F-AG")
 * @param {number} lr номер последней строки из листа с ответами
 * @return {string} strRange строка готовая для вставки в .getRange(). (Пример: "F3:AG3")
 */
function getStringRange(answersSettingsRange_F1, lr) {
  let string = answersSettingsRange_F1.getValue();
  let [firstCol, lastCol] = getSeparatedArray(string);
  let strRange = firstCol+lr+":"+lastCol+lr;
  return strRange;
}

/**
 * Возвращает массив где каждая ячейка результат разделения строки сепаратором (сепараторы:",", "-"", ";"")
 * @param {string} string строка для сепарации
 * @return {Array} resArr массив, где значение каждой ячейки результат сепарации строки
 */
function getSeparatedArray(string) {
  let separator = string.match(/,|-|;/gi)[0];
  let resArr = string.split(separator);
  return resArr;
}

function sendMail() {
  let lr = sheet.getLastRow();
  const clientEmail = sheet.getRange("B" + lr).getValue();
  let jasper1 = UrlFetchApp
    .fetch("https://i.ibb.co/RPT2Ksk/appscript.png")
    .getBlob()
    .setName("jasper1URL");
  let message = "<!DOCTYPE html><html><body><h1>Привет! Єто тетсовое письмо!))</h1><p>Это тело письма в П</p><ul>" +
    "<li>Список 1</li>" +
    "<li>Список 2</li>" +
    "<li>Список 3</li>" +
    "<li>Список 4</li>" +
    '<li>Список 5</li></ul><img src="cid:jasper1URL" alt="asper 1" width="220"></body></html>'
  Logger.log(clientEmail)
  GmailApp.sendEmail(clientEmail, 'Attachment example', 'Тело письма. Показівается, если гаджет не поддерживает ХТМЛ', {
    htmlBody: message,
    inlineImages:
    {
      jasper1URL: jasper1
    }
  });
}
