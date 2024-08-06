/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 */
function doGet() {
  return ContentService.createTextOutput('I just successfully handled your GET request.');
}
function onOpen() {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Open Sidebar', 'showSidebar')
      .addToUi();
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Text Selector');
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Gets the text the user has selected.
 *
 * @return {String} The selected text.
 */
function getSelectedText() {
  const selection = DocumentApp.getActiveDocument().getSelection();
  const text = [];
  if (selection) {
    const elements = selection.getSelectedElements();
    for (let i = 0; i < elements.length; ++i) {
      if (elements[i].isPartial()) {
        const element = elements[i].getElement().asText();
        const startIndex = elements[i].getStartOffset();
        const endIndex = elements[i].getEndOffsetInclusive();
        text.push(element.getText().substring(startIndex, endIndex + 1));
      } else {
        const element = elements[i].getElement();
        if (element.editAsText) {
          const elementText = element.asText().getText();
          if (elementText) {
            text.push(elementText);
          }
        }
      }
    }
  }
  return text.join('\n') || 'No text selected.';
}

/**
 * Replaces the selected text with the specified text.
 *
 * @param {String} newText The text with which to replace the selection.
 * @return {String} Success message.
 */
function replaceSelectedText(newText) {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    let replaced = false;
    const elements = selection.getSelectedElements();
    for (let i = 0; i < elements.length; ++i) {
      if (elements[i].isPartial()) {
        const element = elements[i].getElement().asText();
        const startIndex = elements[i].getStartOffset();
        const endIndex = elements[i].getEndOffsetInclusive();
        element.deleteText(startIndex, endIndex);
        element.insertText(startIndex, newText);
        replaced = true;
      } else {
        const element = elements[i].getElement();
        if (element.editAsText) {
          element.asText().setText(newText);
          replaced = true;
        }
      }
    }
    return replaced ? 'Text successfully replaced.' : 'No text was replaced.';
  } else {
    return 'No text selected.';
  }
}

