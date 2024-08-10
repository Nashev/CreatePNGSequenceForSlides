# CreatePNGSequenceForSlides
Helper for http://github.com/Nashev/MultiPult: allow export of animation, made through Google Presentation

How to use:
* Add this files as a project content at https://script.google.com through editing scripts for your own presentation at https://docs.google.com/presentation 
  * Open presentation
  * open menu ~~Tools / Script editor~~ **Extensions / Apps Script**
  * Add each file and fill him
* Reload presentation
* Use addon's submenu at menu Addons
* Browse and download folder with created PNGs

Other way is use it as library in all your presentations when you once have it in one of them. 
* Open presentation where you have a full copy of this solution. Go to  **Extensions / Apps Script**. then on scripts go to **Project Settings** and copy a **Script ID**
* Open target presentation, go to **Extensions / Apps Script**. then on scripts go to **Project Settings** and check the **Show "appsscript.json" manifest file in editor** checkbox
* Go to **Editor** and there add a Library. Paste there the copied **Script ID** and name library 'CreatePNGsequence'
* in Files into the file 'Code.gs' put this:
```
function onInstall() {
  CreatePNGsequence.onInstall
}

function onOpen(e) {
  CreatePNGsequence.onOpen(e);
}

function savePNGs2() {
  CreatePNGsequence.savePNGs2();
}

function showAboutBox() {
  CreatePNGsequence.showAboutBox();
}

function internalSavePNGs2_GetSlideList() {
  return CreatePNGsequence.internalSavePNGs2_GetSlideList()
}

function internalSavePNGs2_SaveSlide(Index, PresentationName, PrID, SlID, FldName) {
  return CreatePNGsequence.internalSavePNGs2_SaveSlide(Index, PresentationName, PrID, SlID, FldName)
}
```

and into the 'appscript.json' add value "enabledAdvancedServices" into the "dependencies" object:

```
  ...,
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "Slides",
        "serviceId": "slides",
        "version": "v1"
      }
    ],,
    "libraries": [
    ...
```
after this changes it must start working.
