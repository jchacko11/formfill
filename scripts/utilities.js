function repairFormUrl( url ) {
  return url.replace(
    /(entry\.\d{9})([=+,0-9a-zA-Z()]+)(&\1)([=+,0-9a-zA-Z()]+)/g,
    "$1$2$3=__other_option__$3.other_option_response$4"
  );
}

function sanitize(input) {

  return input.replace(/[&<>"'/]/ig, function (m) {
    return {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#x27;',
      "/": '&#x2F;',
    }[m];
  });
}

function getSheetById(id) {
  return SpreadsheetApp.getActive().getSheets().filter(
    function(s) {return s.getSheetId() === id;}
  )[0];
}

function getRandomInt(max) {
  return Math.floor(Math.random() * Math.floor(max));
}

//clear document properties, reset
function clearDocProperties() {
  PropertiesService.getDocumentProperties().deleteAllProperties();
  showSidebar();
}

//check user ready state
function isReady(){
  var formName = getProp("formName")
  if(formName) {
    var selectedQ = getProp("selectedQsName");
    if(selectedQ){
     var data = selectedQ.split("$$|$||$||$|$||$|||$$$$|$|");
     if(data.length != 0){
       var prefilled = getProp("prefillStatus")
       if(prefilled == "true"){
         return 2;
       }else{
         return 1;
       }
     }
    }
  }
  return false;
}

function truncate(input, maxlength) {
  if (input.length > maxlength)
    return input.substring(0, maxlength - 3) + '...';
  else
    return input;
}

//get a 'random' number based on time in ms
function randTime() {
  var d = new Date();
  var n = d.getTime();
  return Math.floor((n - 1576000000000)/100);
  //return n.toString().substring(7); ;

}

//gets all ints getween lower and upper bound inclusive
function getIntsBetween(lower, upper) {
  var list = [];
  for (var i = lower; i <= upper; i++) {
    list.push(i);
  }
  return list;
}

//https://stackoverflow.com/questions/3115982/how-to-check-if-two-arrays-are-equal-with-javascript
function arraysEqual(a, b) {
  if (a === b) return true;
  if (a == null || b == null) return false;
  if (a.length != b.length) return false;

  for (var i = 0; i < a.length; ++i) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}
