// Create Menu
var body = DocumentApp.getActiveDocument().getBody();
var paragraphs = body.getParagraphs();

var maxTokens = PropertiesService.getScriptProperties().getProperty('maxTokens');

function onOpen(e) {
  PropertiesService.getScriptProperties().setProperty('maxTokens', 50);
  var ui = DocumentApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('GPT Connect')
      .addItem('Run on Full Document', 'runFullDoc')
      .addItem("Run on Selected Text", 'runSelected' )
      .addSeparator()
      .addSubMenu(ui.createMenu('Settings')
        .addItem('Max Tokens', 'editTokens')
        .addItem('API Key', 'addAPIKey'))
      .addToUi();
}

function addAPIKey() {
  var ui = DocumentApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Add Your GPT3 API Key',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    PropertiesService.getScriptProperties().setProperty('apiKey', text)
  }
}

function editTokens() {
  var ui = DocumentApp.getUi(); // Same variations.

  var maxTokens = PropertiesService.getScriptProperties().getProperty('maxTokens');

  var result = ui.prompt(
      'Max tokens is currently set at ' + maxTokens + '.',
      'Adjust number (1 token = 4 characters):',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    PropertiesService.getScriptProperties().setProperty('maxTokens', text)
  }
}

function runFullDoc(){
  runGPT(body.getText());
}

function runSelected(){
  var selectedText = getSelection();
  var response = runGPT(selectedText);
  insertText(selectedText, response);
}

function insertText(find, insert){
  for (let i = 0; i < paragraphs.length; i++) {
    var text = paragraphs[i].getText();
        
    if (text.includes(find)==true) 
    
    {
      body.insertParagraph(i, insert)
     }
 }
}

// Connect to GPT3

var copy = body.getText();

// Use editAsText to obtain a single text element containing
// all the characters in the document.
var text = body.editAsText();

function runGPT(content) {
  var maxTokens = PropertiesService.getScriptProperties().getProperty('maxTokens');
  maxTokens = parseInt(maxTokens);
  var apiKey = PropertiesService.getScriptProperties().getProperty('apiKey');
  apiKey = 'Bearer '+ apiKey;
  console.log(apiKey);
  var data = {
    'model': 'text-davinci-003',
    'prompt': content, 
    "temperature": 0, 
    "max_tokens": maxTokens
  };
  var url = "https://api.openai.com/v1/completions";
  var options = {
      "method": "post",
      "contentType": "application/json",
      "headers": {
          "Authorization": apiKey
      },
      "payload": JSON.stringify(data)
  };
  console.log(options);
  var response = UrlFetchApp.fetch(url, options);
  var parse = JSON.parse(response);
  console.log(parse);
  return parse.choices[0].text;
}

function getSelection() {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();
  var ui = DocumentApp.getUi();


  if (!selection) {
    ui.alert("Please highlight some text...");
  }
  else {
    var elements = selection.getSelectedElements();
    // Report # elements. For simplicity, assume elements are paragraphs
    if (elements.length > 1) {
    }
    else {
      var element = elements[0].getElement();
      var startOffset = elements[0].getStartOffset();      // -1 if whole element
      var endOffset = elements[0].getEndOffsetInclusive(); // -1 if whole element
      var selectedText = element.asText().getText();       // All text from element
      // Is only part of the element selected?
      if (elements[0].isPartial())
        selectedText = selectedText.substring(startOffset,endOffset+1);

      // Google Doc UI "word selection" (double click)
      // selects trailing spaces - trim them
      selectedText = selectedText.trim();
      endOffset = startOffset + selectedText.length - 1;

      // Now ready to hand off to format, setLinkUrl, etc.
      return selectedText;
    }
  }
}