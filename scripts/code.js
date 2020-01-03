//list selectable questions (not video, image, etc.)
function listQuestions(id) {
  var form = FormApp.openById(id)
  var items = form.getItems();
  var output = [];

  //if item is supported, push title and id to output
  for (var i = 0; i < items.length; i++) {
    var type = items[i].getType()
    if (type != FormApp.ItemType.IMAGE && type != FormApp.ItemType.VIDEO && type != FormApp.ItemType.PAGE_BREAK && type != FormApp.ItemType.SECTION_HEADER && type != FormApp.ItemType.GRID && type != FormApp.ItemType.CHECKBOX_GRID) {
      output.push([sanitize(items[i].getTitle()), (items[i].getId())]);
    }

  }
  return output;
}

//creates a new sheet with selected questions as column headers
function newSheet() {

  clearProps(["shortenedUrls", "prefillStatus", "printableColumns", "printStatus"])

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var currentSheet = spreadsheet.insertSheet()

  //name current sheet, if already taken, and a random number to it
  try {
    currentSheet.setName(truncate(getProp("formName"), 25) + " Prefill")
  } catch (e) {
    currentSheet.setName(truncate(getProp("formName"), 20) + " Prefill " + randTime())
  }

  setProp("sheetName", currentSheet.getName())
  setProp("sheetId", currentSheet.getSheetId().toString(10))

  //get selected questions and ids
  var selectedQs = getProp("selectedQsName").split(SPLIT)
  var selectedQsId = getProp("selectedQs").split(SPLIT)

  var range = currentSheet.getRange(1, 1, 2, selectedQs.length)

  //format the first two rows
  currentSheet.hideRows(2)
  currentSheet.setFrozenRows(2)
  var values = [selectedQs, selectedQsId]
  console.log(values)
  range.setFontWeight("bold").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setValues(values)

  //delete extraneous rows and columns
  currentSheet.deleteColumns(selectedQs.length + 1, currentSheet.getMaxColumns() - (selectedQs.length))
  currentSheet.deleteRows(200, currentSheet.getMaxRows() - 200)

  //protect range
  var unprotectedRange = currentSheet.getRange(3, 1, currentSheet.getMaxRows() - 2, currentSheet.getMaxColumns())
  currentSheet.protect().setWarningOnly(true).setUnprotectedRanges([unprotectedRange])

  //get range of the first column of user input
  var rangey = currentSheet.getRange(3, 1, currentSheet.getMaxRows() - 2)

  var form = FormApp.openById(getProp("formId"))

  //format each range seperately
  for (var i = 0; i < selectedQs.length; i++) {
    var currentItem = form.getItemById(parseInt(selectedQsId[i], 10))
    //format dates and times
    if (currentItem.getType() != FormApp.ItemType.DATE && currentItem.getType() != FormApp.ItemType.DATETIME && currentItem.getType() != FormApp.ItemType.DURATION && currentItem.getType() != FormApp.ItemType.TIME) {
      rangey.setNumberFormat('@STRING@');
    } else if (currentItem.getType() == FormApp.ItemType.DATE) {
      rangey.setNumberFormat('m/d/yyyy');
    } else if (currentItem.getType() == FormApp.ItemType.DATETIME) {
      rangey.setNumberFormat('m"/"d"/"yyyy" "h":"mm" "am/pm');
    } else if (currentItem.getType() == FormApp.ItemType.DURATION) {
      rangey.setNumberFormat('[h]:mm:ss');
    } else if (currentItem.getType() == FormApp.ItemType.TIME) {
      rangey.setNumberFormat('h:mm am/pm');
    }

    //add in cell dropdown menu for multiple choice items
    if (currentItem.getType() == FormApp.ItemType.MULTIPLE_CHOICE) {
      var choiceOptions = [];
      for each (var item in currentItem.asMultipleChoiceItem().getChoices()) {
        choiceOptions.push(item.getValue())
      }
      var multipleChoiceRule = SpreadsheetApp.newDataValidation().requireValueInList(choiceOptions).build();
      rangey.setDataValidation(multipleChoiceRule);
    }

    //add in cell dropdown menu for list items
    if (currentItem.getType() == FormApp.ItemType.LIST) {
      var choiceOptions = [];
      for each (var item in currentItem.asListItem().getChoices()) {
        choiceOptions.push(item.getValue())
      }
      var listRule = SpreadsheetApp.newDataValidation().requireValueInList(choiceOptions).build();
      rangey.setDataValidation(listRule);
    }

    //add in cell dropdown menu for scale items
    if (currentItem.getType() == FormApp.ItemType.SCALE) {
      var lowerBound = currentItem.asScaleItem().getLowerBound();
      var upperBound = currentItem.asScaleItem().getUpperBound();
      var scaleRule = SpreadsheetApp.newDataValidation().requireValueInList(getIntsBetween(lowerBound, upperBound)).build();
      rangey.setDataValidation(scaleRule);
    }

    //move on to next column
    rangey = rangey.offset(0, 1)
  }
}

function prefillForm(shortenType) {
  clearProps(["shortenedUrls", "prefillStatus", "printableColumns", "printStatus", "printSheet"])

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var currentSheet = getSheetById(parseInt(getProp("sheetId"), 10))

  //temporarily change timezone to GMT
  var timeZone = spreadsheet.getSpreadsheetTimeZone()
  spreadsheet.setSpreadsheetTimeZone("Etc/GMT")

  //removing sheet protection
  var protection = currentSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  if (protection) {
    protection.remove()
  }

  var form = FormApp.openById(getProp("formId"))
  var items = form.getItems()

  //get selected questions and their ids
  var selectedQs = getProp("selectedQsName").split(SPLIT)
  var selectedQsId = getProp("selectedQs").split(SPLIT)

  //if Prefilled Links column doesn't exist, create it
  if (selectedQsId.length == currentSheet.getMaxColumns()) {
    currentSheet.insertColumnAfter(currentSheet.getMaxColumns()).setColumnWidth(currentSheet.getMaxColumns(), 170)
    currentSheet.getRange(1, currentSheet.getMaxColumns()).setValue("Prefilled Links")
  }

  //clear data validation
  currentSheet.getRange(3, (selectedQs.length + 1), currentSheet.getMaxRows() - 2).setDataValidation(null)

  var range = currentSheet.getRange(1, 1, currentSheet.getLastRow(), selectedQs.length)
  var outputRange = currentSheet.getRange(3, (selectedQs.length + 1), currentSheet.getLastRow() - 2)

  //clear error notes and progress colors
  range.clearNote();
  outputRange.setBackground("white")

  var urls = []
  var lastRow = currentSheet.getLastRow();

  for (var i = 0; i < lastRow - 2; i++) {
    if(getProp("emergencyStop") == 'true'){
      setProp("emergencyStop", "false")
      showSidebar()
      return;
    }
    //show user working status
    currentSheet.getRange(i + 3, selectedQs.length + 1).setValue("Working...").setBackground("#fce8b2")

    //get response row
    var userResponse = range.getValues()[i + 2]

    var response = form.createResponse()

    //https://stackoverflow.com/a/26395487/1677912
    for (var j = 0; j < selectedQs.length; j++) {
      //get response from row
      var resp = userResponse[j];
      var currentItem = form.getItemById(parseInt(selectedQsId[j], 10))
      console.log("Question Title: " + currentItem.getTitle())
      if (resp) {
        try {
          console.log("Response: " + resp)
          //create responses
          switch (currentItem.getType()) {
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
                resp = resp.join(','); // Convert array to CSV
              resp = resp.split(/ *, */g); // Convert CSV to array
              response.withItemResponse(item.createResponse(resp))
              break;
            case FormApp.ItemType.DATE:
              item = currentItem.asDateItem();
              console.log(resp.toString());
              response.withItemResponse(item.createResponse(resp))
              break;
            case FormApp.ItemType.DATETIME:
              item = currentItem.asDateTimeItem();
              console.log(resp.toString());
              response.withItemResponse(item.createResponse(resp))
              break;
            case FormApp.ItemType.DURATION:
              item = currentItem.asDurationItem();
              console.log(resp.toString())
              console.log("Duration: " + resp.getUTCHours() + ":" + resp.getUTCMinutes() + ":" + resp.getUTCSeconds())
              response.withItemResponse(item.createResponse(resp.getUTCHours(), resp.getUTCMinutes(), resp.getUTCSeconds()))
              break;
            case FormApp.ItemType.SCALE:
              item = currentItem.asScaleItem();
              resp = +resp;
              response.withItemResponse(item.createResponse(resp))
              break;
            case FormApp.ItemType.TIME:
              item = currentItem.asTimeItem();
              console.log(resp.toString())
              console.log("Time: " + resp.getUTCHours() + ":" + resp.getUTCMinutes())
              response.withItemResponse(item.createResponse(resp.getUTCHours(), resp.getUTCMinutes()))
              break;
            default:
              item = null; // Not handling GRID, IMAGE, PAGE_BREAK, SECTION_HEADER
              break;
          }
        } catch (e) {
          console.error(e)
          var userError = e.toString();
          //create user friendly errors
          switch (currentItem.getType()) {
            case FormApp.ItemType.LIST:
              userError = userError.replace("Exception: ", "")
              break;
            case FormApp.ItemType.MULTIPLE_CHOICE:
              userError = userError.replace("Exception: ", "")
              break;
            case FormApp.ItemType.CHECKBOX:
              userError = userError.replace("Exception: ", "")
              break;
            case FormApp.ItemType.DATE:
              userError = ("Invalid response. Make sure cell is formatted as date.")
              break;
            case FormApp.ItemType.DATETIME:
              userError = ("Invalid response. Make sure cell is formatted as date time.")
              break;
            case FormApp.ItemType.DURATION:
              userError = ("Invalid response. Make sure cell is formatted as duration.")
              break;
            case FormApp.ItemType.SCALE:
              userError = ("Invalid response. Make sure value is within the bounds of the scale.")
              break;
            case FormApp.ItemType.TIME:
              userError = ("Invalid response. Make sure cell is formatted as time.")
              break;
            default:
              userError = ("Error") // Not handling GRID, IMAGE, PAGE_BREAK, SECTION_HEADER
              break;
          }
          //cell specific error
          currentSheet.getRange(i + 3, j + 1).setNote(userError)

        }
      } else {
        console.log("Skipped " + currentItem.getTitle())
      }

    }
    try {
      var url = response.toPrefilledUrl();
      console.log(url)
      urls.push(url)
      console.log("url pushed")
      currentSheet.getRange(i + 3, selectedQs.length + 1).setValue("Response Created").setBackground("#b7e1cd")
    } catch (e) {
      console.log(e)
      currentSheet.getRange(i + 3, selectedQs.length + 1).setValue("Error").setBackground("#f4c7c3")
      urls.push("")
    }

  }

  var out = []

  if (shortenType == "noshort") {
    for each (var link in urls) {
      if(link){
        out.push([link])
      }else{
        out.push(['Error'])
      }
    }
    setProp("shortenedUrls", urls.join("<>"))
  } else {
    //bulk shorten using SHORT or UNGUSSABLE
    if (shortenType == "short") {
      var shortened = shorten(urls, "SHORT")
    } else {
      var shortened = shorten(urls, "UNGUESSABLE")
    }
    for each (var link in shortened) {
      if (link.error) {
        out.push(["Error"])
      } else {
        out.push([link])
      }
    }
    console.log(shortened)
    setProp("shortenedUrls", shortened.join("<>"))
  }

  //output urls
  outputRange.setValues(out)
  //reset timezone to default
  spreadsheet.setSpreadsheetTimeZone(timeZone)
  setProp("prefillStatus", "true")
}

//get the available spreadsheet columns to place on printables
//returns string[]
function getHeaders() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var currentSheet = getSheetById(parseInt(getProp("sheetId"), 10))
  if (currentSheet) {
    var range = currentSheet.getRange(1, 1, 1, currentSheet.getMaxColumns())
    var headers = range.getValues()[0];
    var out = [];

    for each(var header in headers) {
      out.push(sanitize(header))
    }

    return out;
  } else {
    return ["Error"];
  }
}

function createPrintables() {

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var currentSheet = spreadsheet.insertSheet()
  setProp("printSheet", currentSheet.getSheetId().toFixed())
  var shortenedUrls = getProp("shortenedUrls").split("<>")

  //name current sheet, if already taken, and a random number to it
  try {
    currentSheet.setName(truncate(getProp("formName"), 25) + " Printables")
  } catch (e) {
    currentSheet.setName(truncate(getProp("formName"), 20) + " Printables " + randTime())
  }

  //delete extraneous rows and columns
  currentSheet.deleteColumns(3, currentSheet.getMaxColumns() - 2)
  currentSheet.deleteRows(shortenedUrls.length, currentSheet.getMaxRows() - shortenedUrls.length)

  //set row and column sizes
  currentSheet.setRowHeights(1, currentSheet.getMaxRows(), 300)
  currentSheet.setColumnWidth(1, 300).setColumnWidth(2, 425)

  var range1 = currentSheet.getRange(1, 1, currentSheet.getMaxRows())
  var range2 = range1.offset(0, 1)
  var qrCodes = [];

  for each (var link in shortenedUrls) {
    console.info(link)
    console.info(link.error)
    if (link.error || link == '[object Object]') {
      //show error image
      qrCodes.push(['=IMAGE("https://developers.google.com/maps/documentation/maps-static/images/error-image-generic.png")'])
    } else {
      //formula to show qr code
      qrCodes.push(['=IMAGE("https://chart.googleapis.com/chart?cht=qr&chs=500x500&chl=' + link + '")'])
    }
  }
  range1.setFormulas(qrCodes)

  //vars to help build formulas
  var sheetName = getSheetById(parseInt(getProp("sheetId"), 10)).getName();
  sheetName = "'" + sheetName + "'!"
  var item1 = 'Indirect("'
  var item2 = '", false)'
  var char = ", char(10)"

  try {
    var selected = getProp("printableColumns").split(SPLIT);
    console.log(parseInt(selected[0], 10))
    console.log(parseInt(selected[0], 10) - 1)

    //build formulas for data column of printables
    var final;
    switch (selected.length) {
      case 1:
      final = "Concatenate(" + item1 + sheetName + "R[2]C[" + (parseInt(selected[0], 10) - 1) + "]" + item2 + ")";
      break;
      case 2:
      final = "Concatenate(" + item1 + sheetName + "R[2]C[" + (parseInt(selected[0], 10) - 1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[1], 10) - 1) + "]" + item2 + ")"
      break;
      case 3:
      final = "Concatenate(" + item1 + sheetName + "R[2]C[" + (parseInt(selected[0], 10) - 1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[1], 10) - 1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[2], 10) - 1) + "]" + item2 + ")"
      break;
      case 4:
      final = "Concatenate(" + item1 + sheetName + "R[2]C[" + (parseInt(selected[0], 10) - 1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[1], 10) - 1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[2], 10) - 1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[3], 10) - 1) + "]" + item2 + ")"
      break;
      case 5:
      final = "Concatenate(" + item1 + sheetName + "R[2]C[" + (parseInt(selected[0], 10) - 1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[1], 10) - 1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[2], 10) - 1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[3], 10) - 1) + "]" + item2 + char + ", " + item1 + sheetName + "R[2]C[" + (parseInt(selected[4], 10) - 1) + "]" + item2 + ")"
      break;
      default:
      break;
    }

    //set data and formatting second column
    range2.setFormula(final)
    range2.setFontSize(24)
    range2.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("middle")
  } catch (e) {
    console.error(e)
  }
  setProp('printStatus', 'true')
}
