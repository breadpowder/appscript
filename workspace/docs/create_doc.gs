// Define your OpenAI API key
const OPENAI_API_KEY = '';
function onOpen() {
 DocumentApp.getUi()
   .createMenu('ChatGPT')
   .addItem('Start Session', 'showDialog')
   .addToUi();
}


function showDialog() {
const html = HtmlService.createHtmlOutputFromFile('InputForm')
.setWidth(400)
.setHeight(300);
DocumentApp.getUi().showModalDialog(html, 'Enter your prompt');
}




function startSession(prompt) {
appendText(`User: ${prompt}`);
const chatGptResponse = getChatGptResponse(prompt);
applyFormattedText(`ChatGPT-4:\n ${chatGptResponse}`);
}


function appendText(text) {
 const doc = DocumentApp.getActiveDocument();
 const body = doc.getBody();
 body.appendParagraph(text);
}


function applyFormattedText(responseText) {
 const doc = DocumentApp.getActiveDocument();
 const body = doc.getBody();
 const lines = responseText.split('\n'); // Split the response into separate lines


 lines.forEach(line => {
   const headerMatch = line.match(/^(#+)\s*(.*)$/);
   if (headerMatch) {
     const level = headerMatch[1].length; // Count of '#' symbols
     const text = headerMatch[2]; // The text following the '#' symbols
     const paragraph = body.appendParagraph(text);
     switch (level) {
       case 1:
         paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
         break;
       case 2:
         paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING3);
         break;
       case 3:
         paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING4);
         break;
       case 4:
         paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING5);
         break;
       case 5:
         paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING6);
         break;
       default:
         paragraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);
     }
   }


   else if (line.trim().startsWith('- **')) {
     // Process bullet points that start with bold text
     const listItem = body.appendListItem('');
     listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
     processFormattedText(line.trim().substring(2), listItem); // Remove the '- ' and process the rest
   } else {
     // Process other lines that may contain bold formatting
     const paragraph = body.appendParagraph('');
     processFormattedText(line, paragraph);
   }
 });
}


function processFormattedText(text, container) {
 let parts = text.split('**');
 let isBold = false; // Track whether the current section should be bold


 parts.forEach((part, index) => {
   isBold = index === 1


   if (part.length > 0) {
   container.appendText(part).setBold(isBold);


   }
 });
}






function getChatGptResponse(prompt) {
 const url = 'https://api.openai.com/v1/chat/completions';
 const headers = {
   'Authorization': `Bearer ${OPENAI_API_KEY}`,
   'Content-Type': 'application/json'
 };
 const messages = getChatContext();
 messages.push({ role: 'user', content: prompt });
 const payload = {
   'model': 'gpt-4',
   'messages': messages,
   'max_tokens': 450,
   'temperature': 0.7
 };
  const options = {
   'method': 'post',
   'headers': headers,
   'payload': JSON.stringify(payload)
 };
  const response = UrlFetchApp.fetch(url, options);
 const jsonResponse = JSON.parse(response.getContentText());
 const responseText = jsonResponse.choices[0].message.content.trim();
 updateChatContext(prompt, responseText);
 return responseText;
}
function getChatContext() {
 const properties = PropertiesService.getUserProperties();
 const contextString = properties.getProperty('chatContext') || '[]';
 return JSON.parse(contextString);
}
function updateChatContext(prompt, response) {
 const properties = PropertiesService.getUserProperties();
 const context = getChatContext();
 context.push({ role: 'user', content: prompt });
 context.push({ role: 'assistant', content: response });
 properties.setProperty('chatContext', JSON.stringify(context));
}