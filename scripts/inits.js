//TODO: sanitize, withUser Object, margins, prevent overflows (ellipses), Display names of output sheets, error handling

/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
 */

var DIALOG_TITLE = 'Select Questions';
var SIDEBAR_TITLE = 'Form Fill';
var SPLIT = "$$|$||$||$|$||$|||$$$$|$|"
var RED_BUTTON = false;

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Launch', 'showSidebar')
    .addSeparator()
    .addItem('Select Questions', 'showPicker')
    .addItem('Stop Prefilling', 'emergencyStop')
    .addItem('Reset', 'clearDocProperties')
    .addItem('Settings', 'showSettings')
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
  setProp("emergencyStop", 'false')
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
  //setProp("emergencyStop", 'false')
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

/**
 * Open Settings
 */
function showSettings() {
  //setProp("emergencyStop", 'false')
  var ui = HtmlService.createTemplateFromFile('settings/settings')
    .evaluate()
    .setWidth(300)
    .setHeight(250)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(ui, "Settings");
}

//https://ctrlq.org/code/20393-google-file-picker-example
function initPicker() {
  return {
    locale: 'en',
    token: ScriptApp.getOAuthToken(),
    origin: "https://script.google.com",
    developerKey: PropertiesService.getScriptProperties().getProperty("devKey1"),
    dialogDimensions: {
      width: 700,
      height: 425
    },
    picker: {
      viewMode: "LIST",
      mineOnly: true,
      multiselectEnabled: false,
      allowFolderSelect: false,
      navhidden: true,
      hideTitle: true,
    }
  };
}

function emergencyStop(){
  setProp("emergencyStop", 'true')
}
