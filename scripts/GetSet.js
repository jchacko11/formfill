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

//set multiple properties
function setProps(properties, values){
  var docProperties = PropertiesService.getDocumentProperties();
  if(docProperties){
    for(var i = 0; i < properties.length; i++){
      try{
        docProperties.setProperty(properties[i], values[i])
      }catch(e){
        console.error("Property write error: " + e)
      }
    }
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
    docProperties.deleteProperty(property)
  }
}

function clearProps(properties){
  var docProperties = PropertiesService.getDocumentProperties();
  if(docProperties){
    for each (var property in properties){
      docProperties.deleteProperty(property)
    }
  }
}
