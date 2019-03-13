/**
 * @OnlyCurrentDoc
 *
 * Наличие @OnlyCurrentDoc выше как-то влияет на то, какие права будет просить Гугл у пользователя для выполнения этого пакета скриптов
 */

function onOpen(e) {
  // штука, вызываемая обычно при открытии документа. Может добавлять пункты меню, что и делает.
  SlidesApp.getUi().createAddonMenu()
    .addItem('Сохранить все слайды как папку с PNG-файлами на Google Drive', 'savePNGs')
    .addToUi(); // добавляет созданые пункты меню в интерфейс
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
}

function savePNGs() {
  // Страничка ResultPage вызовет функцию internalSavePNGs и красиво покажет её результат
  SlidesApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('ResultPage'),
    'Сохранение всех слайдов как папки с PNG-файлами на Google Drive'
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
