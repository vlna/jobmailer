function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Jobmailer')
      .addItem('Send mail(s)', 'jobmailer')
      .addSeparator()
      .addItem('Create parameters', 'parameters')
      .addSeparator()
      .addItem('Help', 'help')
      .addToUi();
}

function help() {
  var ui = SpreadsheetApp.getUi();
  ui.alert("Select four columns and up to 100 lines with application date, mail, job and answer info/date,\n"
           + "then choose Jobmailer => Send mail(s)");
}

function parameters() {
  if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("parameters") === null) {
    SpreadsheetApp.getActiveSpreadsheet().insertSheet("parameters");
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getRange("A1").setValue("SUBJECT");
    sheet.getRange("B1").setValue("Nevyřízená žádost o pracovní místo");
    sheet.getRange("A2").setValue("MAIL");
    sheet.getRange("B2").setValue("Vážený pane, vážená paní,<BR><BR>\n"
                                  + "Dne {{{Date}}} jsem reagoval na nabídku pracovního místa\n"
                                  + "<I>{{{Job}}}</I>, avšak dosud jsem neobdržel žádnou odpověď.\n"
                                  + "Věřím, že došlo k pouhému nedopatření a již brzy mě budete kontaktovat.<BR><BR>\n"
                                  + "Se srdečným pozdravem<BR><BR>\n"
                                  + "Valentýn Dobrotivý");
    sheet.getRange("A3").setValue("GRACE");
    sheet.getRange("B3").setValue(14);
    
  } else {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Sheet parameters already exist.\nDelete/rename it to create new one.");
  }
}


function jobmailer() {
  var parametry = readparametry("parameters");
  
  Logger.log(parametry);
  
  var ui = SpreadsheetApp.getUi();
  
  var style= 'style="font-family:Calibri;"'
  
  var range = SpreadsheetApp.getActiveSheet().getSelection().getActiveRange();
  var values = range.getValues();
  var displayvalues = range.getDisplayValues();

  var subject = parametry["SUBJECT"];
  var template = parametry["MAIL"];
  var grace = parametry["GRACE"];
  
  if (range.getNumColumns() != 4) {
    Logger.log("Selection must have four columns!")
    throw "Selection must have four columns!"
  }

  if (subject == "") {
    Logger.log("Mail subject SUBJECT cannot be empty!")
    throw "Mail subject SUBJECT cannot be empty!"
  }

  if (template == "") {
    Logger.log("Mail template MAIL cannot be empty!")
    throw "Mail template MAIL cannot be empty!"
  }

  if(isNaN(grace)) {
    Logger.log("Grace period GRACE must be number!")
    throw "Grace period GRACE must be number!"
  }
  
  var today = new Date();
  today.setHours(0,0,0,0);
  
  var len = range.getNumRows();
  var counter = 0;
 
  for (var row = 0; row < len; row++) {
    if (displayvalues[row][1].indexOf('@') > 0) {

      var mailto = displayvalues[row][1].toLowerCase();
      
      if (displayvalues[row][3]=="" && ( today-values[row][0] > grace*86400000)) {

        MailApp.sendEmail({to: mailto, subject: subject,
                                       htmlBody: '<HTML>' + '<BODY ' + style + '>' + template
                                       .replace(new RegExp('{{{Date}}}', 'g'), displayvalues[row][0])
                                       .replace(new RegExp('{{{Job}}}', 'g'), displayvalues[row][2])
                                       + '</BODY></HTML>'
        }); //replyTo: replyto, noReply: true
        counter++;
        Logger.log("Sent to: " + mailto);
        Utilities.sleep(2000);
      } else {
        Logger.log("Skipped: " + mailto);
          
      }
    }
  }
  ui.alert(counter + " mail(s) sent.");
}


function readparametry(sheetname) {
  var parametry = {};
  
  var parametrysheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  var parametrycounter = 1;
  
  if (parametrysheet != null) {
    while ((parametrysheet.getRange(parametrycounter , 1).getValue()) != "") {
      //parametry[parametry.length] = parametrysheet.getRange(vipcounter++ , 1).getValue().toUpperCase();
      
      parametry[parametrysheet.getRange(parametrycounter, 1).getValue().toUpperCase()] = parametrysheet.getRange(parametrycounter , 2).getValue();
      parametrycounter++;
      
    }
  }
  return parametry;
}
