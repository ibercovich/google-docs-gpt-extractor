var templateId = 'google_doc_id_with_relevant_table';
var maxTokens = 100000;// 5000; //for context window limit
var openAIModel = 'gpt-4-1106-preview';// 'gpt-3.5-turbo-16k'; //gpt-4

function extractPrompt() {

  kpis = getKpiTemplate();

  const prompt = `Below you will find notes taken by various analysts about a company being diligenced by a VC.
  The task is to extract some investment-related KPIs from this text. Be succinct, but if a given KPI has multiple answers, list all of them. 
  Below, you can find the list of KPIs of interest:


  "${(kpis.map(subArray => subArray['kpi'])).join('", "')}". 

  
  If a given KPI seems to not be present in the document, just leave "N/A". Don't skip any KPIs. \n

  Format: return results in the following format. Only return results in this JSON format and nothing else. \n

  {
    kpi_1: {response: "response 1"} ,
    kpi_2: {response: "response 2"} ,
    kpi_3: {response: ["response 3a", "response 3b", "response 3c"]} ,
    kpi_4: {response: ["response 4a", "response 4b"]} ,
    kpi_5: {response: "response 5"} ,
  }

  \n
  Notes:\n
  `
  return prompt;
}

function getKpiTemplate() {
  var kpis = [];
  // Open the source document that contains the table
  var sourceDocumentId = templateId;
  var sourceDoc = DocumentApp.openById(sourceDocumentId);
  var sourceBody = sourceDoc.getBody();

  // Assuming the table we want to copy is the first one in the source document
  var sourceTable = sourceBody.getTables()[0];

  // Copy the contents from the source table to the target table
  // skip header
  for (var i = 1; i < sourceTable.getNumRows(); i++) {
    var sourceRow = sourceTable.getRow(i);
    var kpi = sourceRow.getCell(0).getText();
    var benchmark = sourceRow.getCell(3).getText();
    kpis.push({ 'kpi': kpi, 'benchmark': benchmark ? benchmark : null });
  }
  return kpis;
}

function kpiToTable(kpiObj) {

  function getObjectProp(obj, prop, defaultValue) {
    return obj.hasOwnProperty(prop) ? obj[prop] : defaultValue;
  }

  // Open the source document that contains the table
  var sourceDocumentId = templateId;
  var sourceDoc = DocumentApp.openById(sourceDocumentId);
  var sourceBody = sourceDoc.getBody();

  // Assuming the table we want to copy is the first one in the source document
  var sourceTable = sourceBody.getTables()[0];

  // Open the target document where you want to copy the table
  var targetDoc = DocumentApp.getActiveDocument();
  var targetBody = targetDoc.getBody();

  // Create a new table at the beginning of the target document 
  var targetTable = targetBody.insertTable(0);  //.appendTable()

  // Copy the contents from the source table to the target table
  for (var i = 0; i < sourceTable.getNumRows(); i++) {
    var sourceRow = sourceTable.getRow(i);
    var targetRow = targetTable.appendTableRow();

    for (var j = 0; j < sourceRow.getNumCells(); j++) {
      var sourceCell = sourceRow.getCell(j);
      var rowKpi = sourceRow.getCell(0).getText();
      var value = getObjectProp(kpiObj, rowKpi, false);
      if (j == 1 && value) {
        value = JSON.stringify(value['response']);
        var targetCell = targetRow.appendTableCell(value);
      }
      // else if (j == 2 && value) {
      //   value = JSON.stringify(value['benchmark']);
      //   var targetCell = targetRow.appendTableCell(value);

      // }
      else {
        var targetCell = targetRow.appendTableCell(sourceCell.getText());
      }

      // Copy the cell's formatting if necessary
      // For example, to copy the background color:
      // targetCell.setBackgroundColor(sourceCell.getBackgroundColor());
      targetCell.setAttributes(sourceCell.getAttributes());

      // Repeat for other formatting as needed
    }
  }

  targetBody.insertParagraph(1, ""); // insert empty paragraph

  // Clone the table's visual formatting
  targetTable.setAttributes(sourceTable.getAttributes());
}

function extractKpi() {
  // TODO: need to break out long documents
  var docText = DocumentApp.getActiveDocument().getBody().getText();
  var prompt = extractPrompt() + docText;

  var response = callGPT(prompt);
  let objectResponse = JSON.parse(response);
  kpiToTable(objectResponse);
}

function callGPT(prompt) {
  // Replace with your OpenAI API Key
  const openAIKey = 'your_open_AI_key';

  // Form the data payload to send to OpenAI
  const data = {
    model: openAIModel, //'gpt-4', // Replace with the model you intend to use
    response_format: { "type": "json_object" },
    messages: [
      { role: 'system', content: 'You are a financial assistan working for a venture capital firm called ScOp VC.' },
      { role: 'user', content: prompt },
    ],
  };

  // Set up the API request parameters
  const params = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${openAIKey}`,
    },
    payload: JSON.stringify(data),
    muteHttpExceptions: false,
  };

  // Make the API request
  const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', params);

  // Optional: Parse the response and do something with it, like inserting it back into the document
  const responseJson = JSON.parse(response.getContentText());
  if (responseJson.choices && responseJson.choices.length > 0) {
    return responseJson.choices[0].message.content;
  }
  return false;
}

function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Extract KPIs', 'extractKpi')
    .addToUi();
}
