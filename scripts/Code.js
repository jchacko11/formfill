//TODO: sanitize, withUser Object, margins, prevent overflows (ellipses), Display names of output sheets, error handling

/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
 */

var DIALOG_TITLE = 'Select Fields to Pre-fill';
var SIDEBAR_TITLE = 'Form Fill';

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Manage Form Fill', 'showSidebar')
      .addSeparator()
      .addItem('Choose File', 'showDialog')
      .addItem('Reset Form Fill', 'clearDocProperties')
      .addToUi();
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('stepper/stepper')
      .evaluate()
      .setTitle(SIDEBAR_TITLE)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Opens a dialog.
 */
function showPicker() {
  var ui = HtmlService.createTemplateFromFile('picker/picker')
      .evaluate()
      .setWidth(800)
      .setHeight(525)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(ui, DIALOG_TITLE);
}

/**
 * Opens a dialog.
 */
function showChoosePrintables() {
  var ui = HtmlService.createTemplateFromFile('choosePrintables/choosePrintables')
      .evaluate()
      .setWidth(800)
      .setHeight(525)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(ui, "Choose Printable Information");
}

//creates a new sheet with selected questions as column headers
function newSheet() {
  clearProp("shortenedUrls")
  clearProp("prefillStatus")
  clearProp("printableColumns")
  clearProp("printStatus")

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var currentSheet = spreadsheet.insertSheet()

  //name current sheet, if already taken, and a random number to it
  try{
    currentSheet.setName(truncate(getProp("formName"), 25) + " Prefill")
  }catch(e){
    currentSheet.setName(truncate(getProp("formName"), 20) + " Prefill " + randTime())
  }

  setProp("sheetName", currentSheet.getName())

  var split = "$$|$||$||$|$||$|||$$$$|$|"
  setProp("sheetId", currentSheet.getSheetId().toString(10))

  var selectedQs = getProp("selectedQsName").split(split)
  var selectedQsId = getProp("selectedQs").split(split)
  var formId = getProp("")

  var range = currentSheet.getRange(1, 1, 2, selectedQs.length)

  currentSheet.hideRows(2)
  currentSheet.setFrozenRows(2)

  //delete extraneous rows and columns
  currentSheet.deleteColumns(selectedQs.length, currentSheet.getMaxColumns() - (selectedQs.length))
  currentSheet.deleteRows(200, currentSheet.getMaxRows() - 200)

  var values = [selectedQs, selectedQsId]

  range.setFontWeight("bold").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setValues(values)

  //protect range
  var unprotectedRange = currentSheet.getRange(3, 1, currentSheet.getMaxRows()-3, currentSheet.getMaxColumns())
  currentSheet.protect().setWarningOnly(true).setUnprotectedRanges([unprotectedRange])

  var rangey = currentSheet.getRange(3, 1, currentSheet.getMaxRows()-2)
  var form = FormApp.openById(getProp("formId"))
  for(var i = 0; i < selectedQs.length; i++){
    var currentItem = form.getItemById(parseInt(selectedQsId[i], 10))
    if(currentItem.getType() != FormApp.ItemType.DATE && currentItem.getType() != FormApp.ItemType.DATETIME && currentItem.getType() != FormApp.ItemType.DURATION && currentItem.getType() != FormApp.ItemType.TIME){
      rangey.setNumberFormat('@STRING@');
    }

    if(currentItem.getType() == FormApp.ItemType.MULTIPLE_CHOICE){
      var choiceOptions = [];
      for each (var item in currentItem.asMultipleChoiceItem().getChoices()){
        choiceOptions.push(item.getValue())
      }

      var multipleChoiceRule = SpreadsheetApp.newDataValidation().requireValueInList(choiceOptions).build();
      rangey.setDataValidation(multipleChoiceRule);
    }

    if(currentItem.getType() == FormApp.ItemType.LIST){
      var choiceOptions = [];
      for each (var item in currentItem.asListItem().getChoices()){
        choiceOptions.push(item.getValue())
      }

      var listRule = SpreadsheetApp.newDataValidation().requireValueInList(choiceOptions).build();
      rangey.setDataValidation(listRule);
    }

    if(currentItem.getType() == FormApp.ItemType.SCALE){
      var lowerBound = currentItem.asScaleItem().getLowerBound();
      var upperBound = currentItem.asScaleItem().getUpperBound();

      var scaleRule = SpreadsheetApp.newDataValidation().requireValueInList(getIntsBetween(lowerBound, upperBound)).build();
      rangey.setDataValidation(scaleRule);
    }



    rangey = rangey.offset(0, 1)
  }
}

function createPrintables(){

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var currentSheet = spreadsheet.insertSheet()
  setProp("printSheet", currentSheet.getSheetId().toFixed())
  var shortenedUrls = getProp("shortenedUrls").split("<>")

  //name current sheet, if already taken, and a random number to it
  try{
    currentSheet.setName(truncate(getProp("formName"), 25) + " Printables")
  }catch(e){
    currentSheet.setName(truncate(getProp("formName"), 20) + " Printables " + randTime())
  }


  currentSheet.deleteColumns(3, currentSheet.getMaxColumns()-2)
  currentSheet.deleteRows(shortenedUrls.length, currentSheet.getMaxRows()-shortenedUrls.length)

  currentSheet.setRowHeights(1, currentSheet.getMaxRows(), 300)
  currentSheet.setColumnWidth(1, 300).setColumnWidth(2, 425)

  var range1 = currentSheet.getRange(1, 1, currentSheet.getMaxRows())
  var range2 = range1.offset(0, 1)
  var qrCodes = [];

  for each (var link in shortenedUrls){
    console.info(link)
    console.info(link.error)
    if(link.error || link == '[object Object]'){
      qrCodes.push(['=IMAGE("https://developers.google.com/maps/documentation/maps-static/images/error-image-generic.png")'])
    }else{
      qrCodes.push(['=IMAGE("https://chart.googleapis.com/chart?cht=qr&chs=500x500&chl=' + link + '")'])
    }
  }
  var sheetName = getSheetById(parseInt(getProp("sheetId"), 10)).getName();

  var sheetName = "'" + sheetName + "'!"
  var item1 = 'Indirect("'
  var item2 = '", false)'
  var char = ", char(10)"

  //item1
  try{
  var selected = getProp("printableColumns").split(",");
  console.log(parseInt(selected[0], 10))
  console.log(parseInt(selected[0], 10)-1)
  //range2.setValue(parseInt(selected[0], 10)-1)
  //R[2]C[i]

  var final;
  switch(selected.length){
    case 1:
      final = "Concatenate(" + item1 + sheetName + "R[2]C[" + (parseInt(selected[0], 10)-1) + "]" + item2 + ")";
      break;
    case 2:
      final = "Concatenate(" + item1 + sheetName + "R[2]C[" + (parseInt(selected[0], 10)-1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[1], 10)-1) + "]" + item2 + ")"
      break;
    case 3:
      final = "Concatenate(" + item1 + sheetName + "R[2]C[" + (parseInt(selected[0], 10)-1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[1], 10)-1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[2], 10)-1) + "]" + item2 +")"
      break;
    case 4:
      final = "Concatenate(" + item1 + sheetName + "R[2]C[" + (parseInt(selected[0], 10)-1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[1], 10)-1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[2], 10)-1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[3], 10)-1) + "]" + item2 + ")"
      break;
    case 5:
      final = "Concatenate(" + item1 + sheetName + "R[2]C[" + (parseInt(selected[0], 10)-1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[1], 10)-1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[2], 10)-1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[3], 10)-1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[4], 10)-1) + "]" + item2 + ")"
      break;
    default:
      break;
  }

  range2.setFormula(final)

  range2.setFontSize(24)
  range2.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("middle")
  }catch(e){
   console.error(e)
  }

  range1.setFormulas(qrCodes)
  setProp('printStatus', 'true')

}

function initPicker() {
  return {
    locale: 'en',
    token: ScriptApp.getOAuthToken(),
    origin: "https://script.google.com",
    parentFolder: "xyz",
    developerKey: PropertiesService.getScriptProperties().getProperty("devKey1"),
    dialogDimensions: {
      width: 700,
      height: 425
    },
    picker: {
      viewMode: "LIST",
      mineOnly: true,
      //mimeTypes: "image/png,image/jpeg,image/jpg",
      multiselectEnabled: false,
      allowFolderSelect: false,
      navhidden: true,
      hideTitle: true,
      includeFolders: true,
    }
  };
}

//list selectable questions (not video, image, etc.)
function listQuestions(id){
  var form = FormApp.openById(id)
  var items = form.getItems();
  var output = [];
  for(var i = 0; i<items.length; i++){
    var type = items[i].getType()
    if(type != FormApp.ItemType.IMAGE && type != FormApp.ItemType.VIDEO && type != FormApp.ItemType.PAGE_BREAK && type != FormApp.ItemType.SECTION_HEADER && type != FormApp.ItemType.GRID && type != FormApp.ItemType.CHECKBOX_GRID){
      output.push([sanitize(items[i].getTitle()), (items[i].getId())]);
    }

  }
  Logger.log(output)
  return output;
  //1v9dycRRmDFOrmtRjjD2l7V_ZhVgLGlSaH_AJHEPh6KA
}

function main(){
  //listQuestions("1v9dycRRmDFOrmtRjjD2l7V_ZhVgLGlSaH_AJHEPh6KA");
  /*
  Logger.log(shorten(["https://docs.google.com/forms/d/e/sdfghjk-bB0dVEmVxFu5jZiw/viewform?usp=sf_link",
                      "https://docs.google.com/forms/d/e/abc/viewform?usp=sf_link",
                      "https://docs.google.com/forms/d/e/def/viewform?usp=sf_link"], "SHORT"));
                      */
  //var form = FormApp.openById(PropertiesService.getDocumentProperties().getProperty("formId"))
  var form = FormApp.openByUrl("https://docs.google.com/forms/d/19TNmtiw7BKGZHISgw7DWa9pgXTW0OcowCFkYsg0ndGs/edit")
  var items = form.getItems()

  //var selectedQ = getProp("selectedQsName").split("$$|$||$||$|$||$|||$$$$|$|");

  var response = form.createResponse()
  var item = items[2].asMultipleChoiceItem();
  var itemResponse = item.showOtherOption(true).createResponse("Other")
  response.withItemResponse(itemResponse)

  var url = response.toPrefilledUrl();
  Logger.log(shorten([repairFormUrl(url)], "UNGUESSABLE"))

}

//TODO make split global var
function prefillForm(shortenType){
  clearProps(["shortenedUrls", "prefillStatus", "printableColumns", "printStatus", "printSheet"])

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var currentSheet = getSheetById(parseInt(getProp("sheetId"), 10))

  var protection = currentSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  if(protection){
    protection.remove()
  }

  //var currentSheet = spreadsheet.getSheetByName("Sheet11")
  var split = "$$|$||$||$|$||$|||$$$$|$|"

  var form = FormApp.openById(getProp("formId"))
  var items = form.getItems()

  var selectedQs = getProp("selectedQsName").split(split)
  var selectedQsId = getProp("selectedQs").split(split)

  if(selectedQsId.length == currentSheet.getMaxColumns()){
    currentSheet.insertColumnAfter(currentSheet.getMaxColumns()).setColumnWidth(currentSheet.getMaxColumns(), 170)
    currentSheet.getRange(1, currentSheet.getMaxColumns()).setValue("Prefilled Links")
  }

  //clear data validation
  currentSheet.getRange(3, (selectedQs.length + 1), currentSheet.getMaxRows()-2).setDataValidation(null)

  var range = currentSheet.getRange(1, 1, currentSheet.getLastRow(), selectedQs.length)
  //var outputRange = currentSheet.getRange("A14")
  var urls =[]

  for(var i = 0; i < currentSheet.getLastRow()-2; i++){
    var userResponse = range.getValues()[i+2]
    console.log(userResponse)
    var response = form.createResponse()

    //https://stackoverflow.com/a/26395487/1677912
    for(var j = 0; j < selectedQs.length; j++){
      var resp = userResponse[j];
      var currentItem = form.getItemById(parseInt(selectedQsId[j], 10))
      console.log("Question Title: " + currentItem.getTitle())
      if(resp){
        try{
          console.log("Response: " + resp)
          switch(currentItem.getType()){
            case FormApp.ItemType.TEXT:
              var item = currentItem.asTextItem();
              response.withItemResponse(item.createResponse(resp))
              break;
            case FormApp.ItemType.PARAGRAPH_TEXT:
              item = currentItem.asParagraphTextItem();
              response.withItemResponse(item.createResponse(resp))
              break;
            case FormApp.ItemType.LIST:
              item = currentItem.asListItem();
              response.withItemResponse(item.createResponse(resp))
              break;
            case FormApp.ItemType.MULTIPLE_CHOICE:
              item = currentItem.asMultipleChoiceItem();
              response.withItemResponse(item.createResponse(resp))
              break;
            case FormApp.ItemType.CHECKBOX:
              item = currentItem.asCheckboxItem();
              // In a form submission event, resp is an array, containing CSV strings. Join into 1 string.
              // In spreadsheet, just CSV string. Convert to array of separate choices, ready for createResponse().
              if (typeof resp !== 'string')
                resp = resp.join(',');      // Convert array to CSV
              resp = resp.split(/ *, */g);   // Convert CSV to array
              response.withItemResponse(item.createResponse(resp))
              break;
            case FormApp.ItemType.DATE:
              item = currentItem.asDateItem();
              resp = new Date( resp );
              resp.setDate(resp.getDate());
              response.withItemResponse(item.createResponse(resp))
              break;
            case FormApp.ItemType.DATETIME:
              item = currentItem.asDateTimeItem();
              resp = new Date( resp );
              resp.setHours(resp.getHours() - 6);
              response.withItemResponse(item.createResponse(resp))
              break;
            case FormApp.ItemType.DURATION:
              item = currentItem.asDurationItem();
              //if (typeof resp !== 'string')
              //  resp = resp.join(':');      // Convert array to Colon SV
              console.log("Duration Item")
              resp = resp.split(/( *: *)/g);   // Convert Colon SV to array
              console.log(resp)
              response.withItemResponse(item.createResponse(resp[0], resp[2], resp[4]))
              break;
            case FormApp.ItemType.SCALE:
              item = currentItem.asScaleItem();
              resp = +resp;
              response.withItemResponse(item.createResponse(resp))
              break;
            case FormApp.ItemType.TIME:
              item = currentItem.asTimeItem();
              //if (typeof resp !== 'string')
              //  resp = resp.join(':');      // Convert array to Colon SV
              //resp = resp.split(/( *: *)/g);   // Convert Colon SV to array
              //resp = new Date( resp );
              //resp = new Date( resp );
              console.error(resp.toString())
              resp.setHours(resp.getHours() - 6);

              console.error("Time: " + resp.getUTCHours() + ":" + resp.getUTCMinutes())
              response.withItemResponse(item.createResponse(resp.getUTCHours(), resp.getUTCMinutes()))
              break;
            default:
              item = null;  // Not handling GRID, IMAGE, PAGE_BREAK, SECTION_HEADER
              break;
          }
        }catch(e){
          console.error(e)
        }
      }else{
        console.log("Skipped " + currentItem.getTitle())
      }

    }
    try{
      var url = response.toPrefilledUrl();
      console.log(url)
      urls.push(url)
      console.log("url pushed")
    }catch(e){
      console.log(e)
      urls.push("")
    }

  }

  var out = []
  var outputRange = currentSheet.getRange(3, (selectedQs.length + 1), currentSheet.getLastRow()-2)

  if(shortenType == "noshort"){
    for each (var link in urls){
      out.push([link])
    }
    setProp("shortenedUrls", urls.join("<>"))
  }else{

    if(shortenType=="short"){
      var shortened = shorten(urls, "SHORT")
    }else{
      var shortened = shorten(urls, "UNGUESSABLE")
    }
    for each (var link in shortened){
      if(link.error){
        out.push(["Error"])
      }else{
        out.push([link])
      }
    }
    console.log(shortened)
    setProp("shortenedUrls", shortened.join("<>"))
  }

  outputRange.setValues(out)
  setProp("prefillStatus", "true")

  //urls.push(shorten([repairFormUrl(url)], "UNGUESSABLE"))
}

function getHeaders(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var currentSheet = getSheetById(parseInt(getProp("sheetId"), 10))
  if(currentSheet){
    var range = currentSheet.getRange(1, 1, 1, currentSheet.getMaxColumns())
    return range.getValues()[0]
  }else{
    return ["Error"];
  }
}

/*
function getSheetById(spreadsheet, sheetId){
  for each (var sheet in spreadsheet.getSheets()) {
    if(sheet.getSheetId()==sheetId){
      return sheet;
    }
  }
}
*/
