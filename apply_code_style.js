// is called by google docs when a document is open
// adds a menu with a menu item that applies a style to the currently selected text
function onOpen() {
  DocumentApp.getUi()
  .createMenu('Extras')
  .addItem('Apply code style', 'applyCodeStyle')
  .addToUi();
}

// definition of a style to be applied
var style = {
  bold: false,
  backgroundColor: "#DDDDDD",
  fontFamily: DocumentApp.FontFamily.CONSOLAS,
  fontSize: 9
};

// helper function that strips the selected element and passes it to a handler
function withElement(processPartial, processFull) {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      if (element.getElement().editAsText) {
        var text = element.getElement().editAsText();
        if (element.isPartial()) {
          var from = element.getStartOffset();
          var to = element.getEndOffsetInclusive();
          processPartial(element, text, from, to);
        } else {
          processFull(element, text);
        }
      }
    }
  }
}

// called in response to the click on a menu item
function applyCodeStyle() {
  return withElement(
    applyPartialStyle.bind(this, style),
    applyFullStyle.bind(this, style)
  );
}

// applies the style to a selected text range
function applyPartialStyle(style, element, text, from, to) {
  text.setFontFamily(from, to, style.fontFamily);
  text.setBackgroundColor(from, to, style.backgroundColor);
  text.setBold(from, to, style.bold);
  text.setFontSize(from, to, style.fontSize);
}

// applies the style if the entire element is selected
function applyFullStyle(style, element, text) {
  text.setFontFamily(style.fontFamily);
  text.setBackgroundColor(style.backgroundColor);
  text.setBold(style.bold); 
  text.setFontSize(style.fontSize);
}

