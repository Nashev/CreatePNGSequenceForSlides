/**
 * @OnlyCurrentDoc
 *
 * Наличие @OnlyCurrentDoc выше как-то влияет на то, какие права будет просить Гугл у пользователя для выполнения этого пакета скриптов
 */

function onOpen(e) {
  // штука, вызываемая обычно при открытии документа. Может добавлять пункты меню, что и делает.

  if (Session.getActiveUserLocale() == 'ru') {
    SlidesApp.getUi().createAddonMenu()
      //.addItem('Start', 'doTest') // 
      //.addItem('Start', 'savePNGs')
      .addItem('Старт', 'savePNGs2')
      .addItem('О дополнении', 'showAboutBox')
      .addToUi(); // добавляет созданые пункты меню в интерфейс
  } else {
    SlidesApp.getUi().createAddonMenu()
      //.addItem('Start', 'doTest') // 
      //.addItem('Start', 'savePNGs')
      .addItem('Start', 'savePNGs2')
      .addItem('About', 'showAboutBox')
      .addToUi(); // добавляет созданые пункты меню в интерфейс
  }
}

function onInstall(e) {
  // штука для вызова onOpen не при открытии документа, а при добавлении к нему этого пакета с макросами
  onOpen(e);
}

function doTest(e) {
  var d = DriveApp.getRootFolder()
  var x = {'URL':d.getUrl(), 'FolderName':d.getName()}
  // вот таким способом можно посмотреть не под отладчиком содержимое переменной с полями:
  SlidesApp.getUi().alert(JSON.stringify(x))
  // или вот таким:
  Logger.log(x)
}

function LocalizeFileName(FileName) {
   if (Session.getActiveUserLocale() == 'ru') {
     return FileName + '_ru'
   } else {
     return FileName + '_en'
   }   
}

function showAboutBox(e) {
  SlidesApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile(LocalizeFileName('AboutPage')),
    'Create PNG sequence'
  );
}

function savePNGs() {
  // Страничка ResultPage вызовет функцию internalSavePNGs и красиво покажет её результат
  SlidesApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile(LocalizeFileName('ResultPage')),
    'Create PNG sequence'
  );
}

function internalSavePNGs() {
  var presentation = SlidesApp.getActivePresentation();
  var prID = presentation.getId()
  
  var f = DriveApp.getFileById(prID) // файл презентации в гугл-диске
  // берём папку в гугл-диске пользователя, где лежит презентация. Если ещё не сохранял себе - то берём корневую папку гугл-диска.
  if (f.getParents().hasNext()) {
    var d = f.getParents().next()
  } else {
    var d = DriveApp.getRootFolder()
  }
  // делаем во взятой папке папку с именем презентации и датой
  d = d.createFolder(presentation.getName() + '_export_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss"))
  // перебираем слайды, и сохраняем их в PNG
  for(s in presentation.getSlides()){
    var tb = Slides.Presentations.Pages.getThumbnail(prID, presentation.getSlides()[s].getObjectId(), {'thumbnailProperties.mimeType': 'PNG', 'thumbnailProperties.thumbnailSize': 'LARGE'})
    var pngFile = d.createFile(UrlFetchApp.fetch(tb.contentUrl))
    pngFile.setName(presentation.getName() + '_page' + s + '.png')
  }
  return {'URL':d.getUrl(), 'FolderName':d.getName()}
}

function savePNGs2() {
  // Страничка ResultPage вызовет функцию internalSavePNGs и красиво покажет её результат
  SlidesApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile(LocalizeFileName('ResultPage2')),
    'Create PNG sequence'
  );
}

function internalSavePNGs2_GetSlideList() {
  var presentation = SlidesApp.getActivePresentation();
  var prID = presentation.getId()
  
  var f = DriveApp.getFileById(prID) // файл презентации в гугл-диске
  // берём папку в гугл-диске пользователя, где лежит презентация. Если ещё не сохранял себе - то берём корневую папку гугл-диска.
  if (f.getParents().hasNext()) {
    var d = f.getParents().next()
  } else {
    var d = DriveApp.getRootFolder()
  }
  // делаем во взятой папке папку с именем презентации и датой
  d = d.createFolder(presentation.getName() + '_export_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss"))
  // перебираем слайды, и сохраняем их в PNG
  var slides = [];
  for(s in presentation.getSlides()){
    slides.push(presentation.getSlides()[s].getObjectId());
  }
  var result = {'URL':d.getUrl(), 'FolderName':d.getName(), 'PresentationName': presentation.getName(), 'PresentationID': prID, 'Slides': slides};
  //SlidesApp.getUi().alert(JSON.stringify(result));
  return result;
}

function internalSavePNGs2_SaveSlide(Index, PresentationName, PrID, SlID, FldName) {
  Logger.log(1);
  //SlidesApp.getUi().alert(Index + '-' + PrID + '-' + SlID + '-' + FldName);
  var tb = Slides.Presentations.Pages.getThumbnail(PrID, SlID, {'thumbnailProperties.mimeType': 'PNG', 'thumbnailProperties.thumbnailSize': 'LARGE'})
  Logger.log(tb);
      
  var f = DriveApp.getFileById(PrID); // файл презентации в гугл-диске
  Logger.log(f);
  // берём папку в гугл-диске пользователя, где лежит презентация. Если ещё не сохранял себе - то берём корневую папку гугл-диска.
  if (f.getParents().hasNext()) {
    Logger.log(2);
    var d = f.getParents().next()
  } else {
    Logger.log(3);
    var d = DriveApp.getRootFolder()
  }
  Logger.log(d);
  Logger.log(FldName);
  Logger.log(d.getFoldersByName(FldName).hasNext());
  if (!d.getFoldersByName(FldName).hasNext()) {
    Utilities.sleep(1500) // иногда почему-то поначалу созданную папку здесь не видно, а потом видно. Ждям этого 'потом'...
  }
  Logger.log(d.getFoldersByName(FldName).hasNext());
  d = d.getFoldersByName(FldName).next()
  Logger.log(d);
  var pngFile = d.createFile(UrlFetchApp.fetch(tb.contentUrl))
  Logger.log(pngFile);
  pngFile.setName(PresentationName + '_page' + ("000000000" + Index.toString()).slice(-5) + '.png')
  Logger.log('ok');
  return Index
}

