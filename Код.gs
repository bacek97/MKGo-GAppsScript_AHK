function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheets()[0];
//  var lastCell = s.getRange(s.getLastRow(), s.getLastColumn());
//  lastCell.setValue(123);
  var sortedFiles = sortHtmPdf(e);
  var tmpSS = copyHtmTmp(sortedFiles[0]).setTrashed(true);
  var tmpS = SpreadsheetApp.openById(tmpSS.getId()).getSheets()[0];
  var objHTM = readHTM(tmpS);
  if (objHTM == null) return ContentService.createTextOutput(sortedFiles[0]);
  var resp = prepareDir(DriveApp.getFolderById(e.parameter.folderId), objHTM, sortedFiles);
  writeRow(s,objHTM);
  addPhone(objHTM);
  var row = s.getLastRow() + 1;
  return ContentService.createTextOutput(resp);
}

function addPhone(objHTM) {
  var contact = ContactsApp.createContact(null, null, null);
  contact.setGivenName("КПо-"+objHTM.KPnm+"⎯"+objHTM.feeo.join(" ")+"⎯"+objHTM.date.reverse().join("."));
  contact.addPhone(ContactsApp.Field.MOBILE_PHONE, objHTM.phon);
  ContactsApp.getContactGroup("System Group: My Contacts").addContact(contact)
}

function prepareDoc(template, objHTM) {
  var months = ["", "Января", "Февраля", "Марта", "Апреля", "Мая", "Июня", "Июля", 
                "Августа", "Сентября", "Октября", "Ноября", "Декабря"];
  var docbody=DocumentApp.openById(template.getId()).getBody();
  docbody.replaceText('OOO', objHTM.KPnm);
  docbody.replaceText(' 21 ', ' ' + objHTM.date[2] + ' ');
  docbody.replaceText('Декабря', months[objHTM.date[1]]);
  docbody.replaceText('2012', objHTM.date[0]);
  docbody.replaceText('Договоров', objHTM.feeo[0]);
  docbody.replaceText('Шаблон', objHTM.feeo[1]);
  docbody.replaceText('Распечатович', objHTM.feeo[2]);
  docbody.replaceText('Дoговоров_Шаблoн_Распечатoвич', objHTM.feeo.join(" "));
  docbody.replaceText('3я_улСтроителей_д25_кв12', objHTM.addr);
  docbody.replaceText("78005553535", objHTM.phon.split("+")[1]);
}

function prepareDir(folderKP,objHTM,sortedFiles) {
  var oldFolder = folderKP.searchFolders('title contains "КП-'+objHTM.KPnm+'"');
  // Перемещение старых папок в Архив
  while (oldFolder.hasNext()) {
    DriveApp.getFileById(oldFolder.next().getId()).moveTo(DriveApp.getFolderById("1r3nvqyf2V98Oc8dVJ1e-TjhqrmWUk-lB"))  
//  folder.setTrashed(true); //Access Denied
}
  var newFolderName = (objHTM.date.reverse().join(".")+" КП-"+objHTM.KPnm+" "+objHTM.price+" "+sortedFiles[0].getName().slice(0,-4)).slice(2);
  var newFolder = folderKP.createFolder(newFolderName);
  DriveApp.getFileById('1-4djyun2Q03E6md9WW6vaaFQmq7DJnOS').makeCopy(newFolder); // Скан.bat
  var template = DriveApp.getFileById('1kfA9ykGCBH4FI4beOx9Y0FPPLObXJu8IVrX-aQZzJlI').makeCopy("Договор-"+objHTM.KPnm+" "+objHTM.feeo[0],newFolder);
  if (sortedFiles[1] != null) sortedFiles[1].makeCopy("КП-"+objHTM.KPnm+" "+objHTM.feeo[0]+".pdf",newFolder);
  prepareDoc(template, objHTM);
  return (newFolderName+";"+template.getId());
}

function readHTM(tmpS) {
  var objHTM = {};
  var range = tmpS.getRange("A1:D5");
  if (range.getValues()[2][0] != "Плательщик:") return (null);
  
  objHTM.feeo = range.getValues()[3][1].split(" ИНН")[0].split(" ");
  if (objHTM.feeo.length != 3) return (null);
  
  objHTM.date = range.getValues()[0][0].split(" ")[1].split(".")
  if (objHTM.date[2] < 2000 || 2222 < objHTM.date[2]) return (null);
  
  objHTM.KPnm = range.getValues()[1][0].split("№ ")[1];
  if (objHTM.KPnm < 1 || 999 < objHTM.KPnm) return (null);
  
  objHTM.addr = range.getValues()[3][0];
  if (objHTM.addr.length < 2 || 44 < objHTM.addr.length) return (null);
  
  objHTM.phon = range.getValues()[4][1].split("Тел. ")[1].split(" Факс")[0]
  if (objHTM.phon.length != 12 || objHTM.phon[0] != "+") return (null);
  
  var vseItogo = tmpS.getRange("A:A").createTextFinder('Итого').findAll();
  objHTM.price = tmpS.getRange(vseItogo[vseItogo.length-1].getRow(), 2).getValue().split(" руб.")[0];
  
  tmpS.getRange("E2").setFormula('=IFS(COUNTIFS(A:A;"Цвет")=COUNTIFS(A:A;"Цвет";B:B;"");3;COUNTIFS(A:A;"Цвет";B:B;"*Лам*") > 0; 30;COUNTIFS(A:A;"Цвет";B:B;"*Лам*") = 0; 15)');
  objHTM.ddln = tmpS.getRange("E2").getValue()
  
  
  return (objHTM)
}

function writeRow(s,objHTM) {
  var row = s.getLastRow() + 1;
  s.getRange("A" + row).setValue(objHTM.feeo.join(" "));
  s.getRange("B" + row).setValue("КП-"+objHTM.KPnm+"\n"+objHTM.phon);
  s.getRange("C" + row).setValue("⎯⎯⎯⎯");
  s.getRange("D" + row).setFormula((objHTM.ddln < 30) ? "=WORKDAY($K"+row+";"+objHTM.ddln+")" : "=$K"+row+"+"+objHTM.ddln);
  s.getRange("E" + row).setFormula((objHTM.ddln < 30) ? "=WORKDAY($K"+row+";"+objHTM.ddln+")" : "=$K"+row+"+"+objHTM.ddln);
//  s.getRange("E" + row).setValue("⎯⎯⎯⎯⎯");
  
  s.getRange("F" + row).setValue("⎯⎯⎯⎯⎯");
  s.getRange("G" + row).setValue("⎯⎯⎯⎯⎯");
  s.getRange("I" + row).setValue(objHTM.addr);
  s.getRange("K" + row).setValue(objHTM.date.join("."));
  s.getRange("H" + row).setValue(objHTM.price);
  s.getRange("J"+row).setFormula("=$J"+(row-1)+"+1");
  s.getRange("L"+row).setFormula('=SUMIFS(H:H;K:K;(">="&TEXT(EDATE(K:K;0);"MM/YYYY"));K:K;"<"&TEXT(EDATE(K:K;1);"MM/YYYY"))');
  s.getRange("M"+row).setFormula('=SUMIFS(H:H;K:K;">=1/"&YEAR(K:K);K:K;"<1/"&YEAR(K:K)+1)');
  s.getRange("N"+row).setFormula('=SUM(H:H)');
  s.getRange((row-1)+":"+(row-1)).copyTo(s.getRange((row)+":"+(row)), {formatOnly:true})
  return (objHTM);
}

function sortHtmPdf(e) {
  var folder = DriveApp.getFolderById(e.parameter.folderId);
  var aFN = [e.parameter.fileName1, e.parameter.fileName2];
  if (aFN[0].slice(-4).toLowerCase() == ".htm" && aFN[1].slice(-4).toLowerCase() != ".htm") {
  aFN = [
    (folder.searchFiles('title contains "'+aFN[0]+'"').hasNext() ? folder.searchFiles('title contains "'+aFN[0]+'"').next() : null),
    (aFN[1].slice(-4).toLowerCase() == ".pdf" && folder.searchFiles('title contains "'+aFN[1]+'"').hasNext()) ? folder.searchFiles('title contains "'+aFN[1]+'"').next() : null];
} else if (aFN[0].slice(-4).toLowerCase() != ".htm" && aFN[1].slice(-4).toLowerCase() == ".htm") {
    aFN = [
    (folder.searchFiles('title contains "'+aFN[1]+'"').hasNext() ? folder.searchFiles('title contains "'+aFN[1]+'"').next() : null),
    (aFN[0].slice(-4).toLowerCase() == ".pdf" && folder.searchFiles('title contains "'+aFN[0]+'"').hasNext()) ? folder.searchFiles('title contains "'+aFN[0]+'"').next() : null];
} else {
  Logger.log("Нет HTM");
  aFN = ([null, null]);
}
  return (aFN);
}

function copyHtmTmp(source) {
  var tmpXLSX = Drive.Files.insert({title: "tmpXLSX", mimeType: (MimeType.MICROSOFT_EXCEL)}, source.getBlob());
  DriveApp.getFileById(tmpXLSX.id).setTrashed(true);
//  DriveApp.createFile('New HTML File', source.getBlob(), MimeType.MICROSOFT_EXCEL);
  
    var tmpHTM = Drive.Files.copy({title: 'tmpHTM'}, tmpXLSX.id, {convert: true});
    return (DriveApp.getFileById(tmpHTM.getId()));
}

// https://docs.google.com/spreadsheets/d/1T1Gn-XeexJsnbZ4C6XOHQ_8D637JsoyoEFi84Xjkao4/export?exportFormat=pdf&gid=0&#38;range=A:H&#38;top_margin=0.25&#38;bottom_margin=0.25&#38;left_margin=0.3&#38;right_margin=0.2

