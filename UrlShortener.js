//batch shorten Urls
//takes array of urls to shorten and 'SHORT' or 'UNGUESSABLE' for suffixLength
function shorten(urlArray, suffixLength){
  var request = []
    var out = []
    
    for(var i = 0; i < urlArray.length; i++){
      //use createRequest function to load request var with fetch requests
      request.push(createRequest(urlArray[i], suffixLength))
    }
    
    var postContent = UrlFetchApp.fetchAll(request);
  
    for(var i = 0; i < postContent.length; i++){
      //load output variable with short Links
      try{
        var parsed = JSON.parse(postContent[i])
        if(parsed.shortLink){
          out.push(parsed.shortLink)
        }else{
          out.push(parsed)
        }
      }catch(e){
        out.push(e);
      }
    }
    return out;
}

//creates a fetch request for fetchAll method
//takes url to shorten and 'SHORT' or 'UNGUESSABLE' for suffixLength
function createRequest(url, suffixLength){
    var AUTH_KEY = PropertiesService.getScriptProperties().getProperty("firebaseAuthKey");
    var postdata = {
      "dynamicLinkInfo": {
        "domainUriPrefix": "https://s.formfill.app",
        "link": url
      },
      "suffix": {
        "option": suffixLength
      }
    };
    var postoptions = {
      'url' : 'https://firebasedynamiclinks.googleapis.com/v1/shortLinks?key=' + AUTH_KEY,
      'method' : 'post',
      'contentType': 'application/json',
      // Convert the JavaScript object to a JSON string.
      'payload' : JSON.stringify(postdata),
      'muteHttpExceptions': true
    };
    return postoptions;
}