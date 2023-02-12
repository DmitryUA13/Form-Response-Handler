const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheetResponses = ss.getSheetByName('Form Responses 1');
const sheetSettings = ss.getSheetByName('Settings');
const answersSettingsRange_F1 = sheetSettings.getRange('G2');
const answersSettingsRange_F2 = sheetSettings.getRange('G3');
const answersSettingsRange_F3 = sheetSettings.getRange('G4');


function dataManipulationAlgorithm() {

  let lr = sheetResponses.getLastRow();
  Logger.log("LAST ROW: " + lr)
  let multiplicator = sheetSettings.getRange('J2').getValue();
  let answersScore_Arr = sheetSettings.getRange('K2:L').getValues().filter(item => item[0] != '');
  let amswersScore_List = new Map(answersScore_Arr);
  let numberOfAnswerToMultiplicate = sheetSettings.getRange('J3').getValue().toString().split(',');
  let answersResults_F1 = sheetResponses.getRange(getStringRange(answersSettingsRange_F1, lr)).getValues()[0];
  let sum_F1 = 0;
  let reg = /.*\(/gi;

  for (let i = 0; i < answersResults_F1.length; i++) {
    let answersResults_F1_SearchPhrase = answersResults_F1[i].toString().match(reg)[0];
    answersResults_F1_SearchPhrase = answersResults_F1_SearchPhrase.substring(0, answersResults_F1_SearchPhrase.length - 2);
    if (numberOfAnswerToMultiplicate.find(item => item == i + 1) != undefined) {
      sum_F1 += amswersScore_List.get(answersResults_F1_SearchPhrase) * multiplicator;
    } else {
      sum_F1 += amswersScore_List.get(answersResults_F1_SearchPhrase);
    }
  }

  let firstPartTextRange = 'O1';
  let valuesRange_F1 = 'M2:O5';
  let responseTextToClientOn_F1 = getTetxMessageForSlelectedScoreResult(firstPartTextRange, valuesRange_F1, sum_F1);


  let answersResults_F2 = sheetResponses.getRange(getStringRange(answersSettingsRange_F2, lr)).getValues()[0]
  let avgOfAnswers_f2 = getAverage(answersResults_F2);
  let valueNumAnswToCalculateDiapason_F2 = sheetSettings.getRange('M6:O8').getValues();
  let responseTextToClientOn_F2 = getTextSelectionResult(valueNumAnswToCalculateDiapason_F2, avgOfAnswers_f2);
  
  // Форма 2 Главній ответ
  responseTextToClientOn_F2 = getReplacetStringText(responseTextToClientOn_F2, '{score}', avgOfAnswers_f2);

  let avgAwarenessOfDreamingScore = getArrOfAnswerToCalculate("J4", answersResults_F2);
  //Текст для отправки  Awareness Dreaming
  let awarenessOfDreamingScoreText = getReplacetStringText(sheetSettings.getRange("I4").getValue(),'{score}', avgAwarenessOfDreamingScore );
  let avgOfDayDreamingScore = getArrOfAnswerToCalculate("J5", answersResults_F2);
  //Текст для отправки  Day Dreaming
  let dayDreamingScoreText = getReplacetStringText(sheetSettings.getRange("I5").getValue(),'{score}', avgOfDayDreamingScore );
  let avgOfDreamSensationsScore = getArrOfAnswerToCalculate("J6", answersResults_F2);
  //Текст для отправки  Day Dreaming
  let dreamSensationsScoreText = getReplacetStringText(sheetSettings.getRange("I6").getValue(),'{score}', avgOfDreamSensationsScore);
  let avgOfDejaStatesScore = getArrOfAnswerToCalculate("J7", answersResults_F2);
  //Текст для отправки  Day Dreaming
  let dejaStatesScoreText = getReplacetStringText(sheetSettings.getRange("I7").getValue(),'{score}', avgOfDejaStatesScore);
  let avgOfComprehensibilityScore = getArrOfAnswerToCalculate("J8", answersResults_F2);
  //Текст для отправки  Day Dreaming
  let comprehensibilityScoreText = getReplacetStringText(sheetSettings.getRange("I8").getValue(),'{score}', avgOfComprehensibilityScore);
  let avgOfIntensityOfSensesScore = getArrOfAnswerToCalculate("J9", answersResults_F2);
  //Текст для отправки  Day Dreaming
  let intensityOfSensesScoreText = getReplacetStringText(sheetSettings.getRange("I9").getValue(),'{score}', avgOfIntensityOfSensesScore);


  let answersResults_F3 = sheetResponses.getRange(getStringRange(answersSettingsRange_F3, lr)).getValues()[0]
  let avgOfAnswers_f3 = getAverage(answersResults_F3);
  let valueNumAnswToCalculateDiapason_F3 = sheetSettings.getRange('M9:O11').getValues();
  let responseTextToClientOn_F3 = getTextSelectionResult(valueNumAnswToCalculateDiapason_F3, avgOfAnswers_f3);
  responseTextToClientOn_F3 = getReplacetStringText(responseTextToClientOn_F3, '{score}', avgOfAnswers_f3);

  //Отправляем письмо
  let arrMessage = [
    responseTextToClientOn_F1,
    responseTextToClientOn_F2,
    awarenessOfDreamingScoreText,
    dayDreamingScoreText,
    dreamSensationsScoreText,
    dejaStatesScoreText,
    comprehensibilityScoreText,
    intensityOfSensesScoreText,
    responseTextToClientOn_F3
  ]
  sendMail(arrMessage);

}


/**
 * Функция возвращает строку с указание диапазона в котором ищем ответы 
 * @param {string} answersSettingsRange_F1 диапазон с строкой диапазона поиска ответов (Пример: 'F-AG')
 * @param {number} lr номер последней строки из листа с ответами
 * @return {string} strRange строка готовая для вставки в .getRange(). (Пример: 'F3:AG3')
 */
function getStringRange(answersSettingsRange_F1, lr) {
  let string = answersSettingsRange_F1.getValue();
  let [firstCol, lastCol] = getSeparatedArray(string);
  let strRange = firstCol + lr + ':' + lastCol + lr;
  return strRange;
}

/**
 * Возвращает массив где каждая ячейка результат разделения строки сепаратором (сепараторы:',', '-', ';')
 * @param {string} string строка для сепарации
 * @return {Array} resArr массив, где значение каждой ячейки результат сепарации строки
 */
function getSeparatedArray(string) {
  let separator = string.match(/,|-|;/gi)[0];
  let resArr = string.split(separator);
  return resArr;
}

/**
 * Функция отправлет текст сообщения клиенту
 */
function sendMail(arrMessage) {
  let lr = sheetResponses.getLastRow();
  const clientEmail = sheetResponses.getRange('B' + lr).getValue();
  const clientName = sheetResponses.getRange('C' + lr).getValue();
  let [responseTextToClientOn_F1,
    responseTextToClientOn_F2,
    awarenessOfDreamingScoreText,
    dayDreamingScoreText,
    dreamSensationsScoreText,
    dejaStatesScoreText,
    comprehensibilityScoreText,
    intensityOfSensesScoreText,
    responseTextToClientOn_F3] = arrMessage;

  let jasper1 = UrlFetchApp
    .fetch('https://i.ibb.co/RPT2Ksk/appscript.png')
    .getBlob()
    .setName('jasper1URL');
  let message =
    `<!DOCTYPE html><html><body><h1>${clientName}, below is the detailed information that we received by processing the form</h1><p></p>` +
    '<H3>Sleep Quality Questionnaire (SQS)</H3>'+
    `<p>${responseTextToClientOn_F1}</p>`+
    '<H3>Dream Questionnaire (ICP Horton)</H3>'+
    `<p>${responseTextToClientOn_F2}</p>`+
    '<ul>'+
    `<li>${awarenessOfDreamingScoreText}</li>` +
    `<li>${dayDreamingScoreText}</li>` +
    `<li>${dreamSensationsScoreText}</li>` +
    `<li>${dejaStatesScoreText}</li>` +
    `<li>${comprehensibilityScoreText}</li>` +
    `<li>${intensityOfSensesScoreText}</li></ul>` +
    '<H3>Lucid Dreaming Questionnaire (LUSK)</H3>'+
    `<p>${responseTextToClientOn_F3}</p>`+
    '<img src="cid:jasper1URL" alt="asper 1" width="220">' +
    '</body></html>'
  Logger.log(clientEmail)
  GmailApp.sendEmail(
    clientEmail, 
    `${clientName}, the answer to the survey you just completed is in this email`,
    `${clientName}, below is the detailed information that we received by processing the form \n`+
    `${responseTextToClientOn_F1} \n ${responseTextToClientOn_F2} \n `+
    `${awarenessOfDreamingScoreText} \n ${dayDreamingScoreText} \n ${dreamSensationsScoreText} \n ${dejaStatesScoreText} \n ${comprehensibilityScoreText} \n ${intensityOfSensesScoreText} \n ${responseTextToClientOn_F3}`, 
    {
    htmlBody: message,
    inlineImages:
    {
      jasper1URL: jasper1
    }
  });
}

/**
 * Функция возвращает текст согласно набранным баллам
 */
function getTetxMessageForSlelectedScoreResult(firstPartTextRange, valuesRange, score) {
  let firstPartText = sheetSettings.getRange(firstPartTextRange).getValue().toString();
  let resulFirstPartTexWhithScore = getReplacetStringText(firstPartText, '{score}', score);
  let valuesScoreDiapasonCorrespondingResponseText = sheetSettings.getRange(valuesRange).getValues();
  let textSelectionResult = getTextSelectionResult(valuesScoreDiapasonCorrespondingResponseText, score);
  Logger.log(resulFirstPartTexWhithScore + " " + textSelectionResult)
  return resulFirstPartTexWhithScore + " " + textSelectionResult
}

/**
 * Функция возвращает строку с замененным шаблоном на переданное значение
 * @param {string} string строка с шаблоном
 * @param {string} pattern шаблон для поиска в строке
 * @param {string} newValue новое значение для вставки вместо шаблона
 * @return {string} resultString строка с замененным шаблоном на новое значение
 */
function getReplacetStringText(string, pattern, newValue) {
  let resultString = string.replace(pattern, newValue);
  return resultString;
}


/**
 * Функция ищет ответ по числовому диапазону согласно набранным клиентом баллов в опроснике формы и возвращает его
 * @param {Array} valuesScoreDiapasonCorrespondingResponseText массив, где первые две ячейки - диапазон (число минимум, число максимум), третья ячейка - текст для вставки в письмо
 * @param {number} score сколько пользователь набрал баллов в опроснике формы
 * @return {string} correspondingResponseText выбранный текст согласно набранных пользователем баллов в опроснике
 */
function getTextSelectionResult(valuesScoreDiapasonCorrespondingResponseText, score) {
  let correspondingResponseText = '';
  for (let i = 0; i < valuesScoreDiapasonCorrespondingResponseText.length; i++) {
    if (score >= valuesScoreDiapasonCorrespondingResponseText[i][0] && score <= valuesScoreDiapasonCorrespondingResponseText[i][1]) {
      correspondingResponseText = valuesScoreDiapasonCorrespondingResponseText[i][2];
    }
  }

  return correspondingResponseText;
}



/**
 * Функция возвращает среднее от суммы значений массивам
 * @param {Array} arrOfNumbers массив чисел
 * @return {float} среднее от суммы значений в массиве
 */
function getAverage(arrOfNumbers) {
  let summ = arrOfNumbers.reduce((acc, nextItem) => {
    return acc += nextItem;
  })
  let average = (summ / arrOfNumbers.length).toFixed(2);
  Logger.log(average)
  return average;
}

/**
 * Функция возвращает среднее из найденных в основном массиве ответов согласно заданной строке с номерами ответов
 * @param {string} range диапазон с строкой номеров ответов через запятую
 * @param {Array} answersResults_F2 массив с ответами 
 * @returm {float} resAvg среднее от найденных ответов
 */
function getArrOfAnswerToCalculate(range, answersResults_F2) {
  let arrOfAnswToCalculate_F2 = sheetSettings.getRange(range).getValues()[0].toString();
  let arrOfDaysToCalculate = getSeparatedArray(arrOfAnswToCalculate_F2);
  // Массив с ответами
  let arrOfAnswToCalToCalculate = [];
  arrOfDaysToCalculate.map(item => {
    arrOfAnswToCalToCalculate.push(answersResults_F2[item-1]);
  })
  let resAvg = getAverage(arrOfAnswToCalToCalculate);
  return resAvg;
}
