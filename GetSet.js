//getters and setters

function setCurrentForm(value){
  var docProperties = PropertiesService.getDocumentProperties();
  if(docProperties){
    docProperties.setProperty("formId", (value))
  }
}

function getCurrentForm(){
  //check if undefines
  var docProperties = PropertiesService.getDocumentProperties();
  if(docProperties){
    var formId = docProperties.getProperty("formId")
    if(formId){
      return formId;
    }
  }
}

function clearCurrentForm(){
  var docProperties = PropertiesService.getDocumentProperties();
  if(docProperties){
    var formId = docProperties.deleteProperty("formId")
  }
}

function setSelected(val){
  var docProperties = PropertiesService.getDocumentProperties();
  if(docProperties){
    docProperties.setProperty("selectedQs", (val))
  }
}

function getSelected(){
  //check if undefines
  var docProperties = PropertiesService.getDocumentProperties();
  if(docProperties){
    var selectedQs = docProperties.getProperty("selectedQs")
    if(selectedQs){
      return selectedQs;
    }
  }
}

function setProp(property, val){
  var docProperties = PropertiesService.getDocumentProperties();
  if(docProperties){
    docProperties.setProperty(property, (val))
  }
}

function getProp(property){
  //check if undefines
  var docProperties = PropertiesService.getDocumentProperties();
  if(docProperties){
    var selectedQs = docProperties.getProperty(property)
    if(selectedQs){
      return selectedQs;
    }
  }
}

function clearProp(property){
  var docProperties = PropertiesService.getDocumentProperties();
  if(docProperties){
    var formId = docProperties.deleteProperty(property)
  }
}
