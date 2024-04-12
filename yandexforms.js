function processIncomingEmails() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var keyword = "Заявка на перевозку"; // Ключевое слово в теме письма
    var senderEmail = "64c7fcbcc769f19adcc7cb03@forms-mailer.yaconnect.com"; // Адрес отправителя
  
    var threads = GmailApp.search("label:inbox is:unread from:" + senderEmail); // Поиск непрочитанных писем
  
    for (var i = 0; i < threads.length; i++) {
      var messages = threads[i].getMessages();
      for (var j = 0; j < messages.length; j++) {
        var message = messages[j];
        var subject = message.getSubject();
        var isUnread = message.isUnread(); // Проверка наличия статуса "непрочитано"
        
        if (isUnread && subject.indexOf(keyword) !== -1) {
          var dateReceived = message.getDate(); // Дата и время получения письма
          var body = message.getPlainBody(); // Получение текста письма
            
          // Регулярки для определения границ данных
          var dateRegex = /Дата:\s*{([^}]+)}/;
          var timeRegex = /Время:\s*{([^}]+)}/;
          var startRegex = /НТП:\s*{([^}]+)}/;
          var midRegex = /ПТ:\s*{([^}]+)}/;
          var finRegex = /КТП:\s*{([^}]+)}/;
          var fiozRegex = /ФИОЗ:\s*{([^}]+)}/;
          var aimRegex = /Цель:\s*{([^}]+)}/;
          var fiopRegex = /ФИОП:\s*{([^}]+)}/;
          var carRegex = /МАШ:\s*{([^}]+)}/;
          var comRegex = /КОМ:\s*{([^}]+)}/;
          var oneRegex = /1:\s*{([^}]+)}/;
          var twoRegex = /2:\s*{([^}]+)}/;
          var threeRegex = /3:\s*{([^}]+)}/;
          var trackRegex = /ID заявки:\s*{([^}]+)}/;
          var fourRegex = /4:\s*{([^}]+)}/;
          var orgRegex = /ОРГ:\s*{([^}]+)}/;
          var eventRegex = /ТМ:\s*{([^}]+)}/;
          var eventdateRegex = /ДМ:\s*{([^}]+)}/;
  
          // Вытаскиваем данные функцией ниже
          var dateValue = extractValue(body, dateRegex);
          var timeValue = extractValue(body, timeRegex);
          var startValue = extractValue(body, startRegex);
          var midValue = extractValue(body, midRegex);
          var finValue = extractValue(body, finRegex);
          var fiozValue = extractValue(body, fiozRegex);
          var aimValue = extractValue(body, aimRegex);
          var fiopValue = extractValue(body, fiopRegex);
          var carValue = extractValue(body, carRegex);
          var comValue = extractValue(body, comRegex);
          var oneValue = extractValue(body, oneRegex);
          var twoValue = extractValue(body, twoRegex);
          var threeValue = extractValue(body, threeRegex);
          var trackValue = extractValue(body, trackRegex);
          var fourValue = extractValue(body, fourRegex);
          var orgValue = extractValue(body, orgRegex);
          var eventValue = extractValue(body, eventRegex);
          var eventdateValue = extractValue(body, eventdateRegex);
  
          // Вставка данных в таблицу
          sheet.appendRow([dateReceived, dateValue, timeValue, oneValue, twoValue, startValue, midValue, finValue, fiozValue, aimValue, fiopValue, carValue, comValue, threeValue, "", "", "", trackValue, "", orgValue, eventValue, eventdateValue]);
  
          // Пометить письмо как прочитанное
          message.markRead();
        }
      }
    }
  }
  
  // Функция для "вытягивания" значений
  function extractValue(text, regex) {
    var match = text.match(regex);
    return match ? match[1].trim() : "";
  }
  