// Add our option to the menu
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu().addItem('Transpose', 'main').addToUi();
}

// Install
function onInstall(e) {
  onOpen(e);
}

// Entry
function main() {
  var body = DocumentApp.getActiveDocument().getBody();
  
  var ui = DocumentApp.getUi();
  var result = ui.prompt("Transpose", "How many notes do you want to transpose?", ui.ButtonSet.OK_CANCEL);
  
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  
  if (button == ui.Button.OK) {
    N = parseInt(text);
    recurse(body);
  }
}

function recurse(elem) {
  if (elem.getType() == DocumentApp.ElementType.TEXT) {
    // base case - call replace function
    replaceNotes(elem);
  } else {
    for (var i = 0; i < elem.getNumChildren(); i++) {
      recurse(elem.getChild(i));
    }
  }
}

// User inputs
var N = 1;

// TODO: Definitely can improve on this regex. Right now we only allow certain punctuation marks before/after the note, but we can be more lenient imo
var OVERALL_REGEX = /^[()\-,.]*[A-G][#b]?[()\-,.]*$/
var NOTE_REGEX = /[A-G][#b]?/

var noteArray = ["A", "A#", "B", "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#"];
var flatMap = {
  Ab: "G#",
  Bb: "A#",
  Cb: "B",
  Db: "C#",
  Eb: "D#",
  Fb: "E",
  Gb: "F#",
};

function transpose(note) {
  // Convert all flats to sharps
  if (note.indexOf("b") === note.length - 1) {
    note = flatMap[note];
  }

  var start = noteArray.indexOf(note);
  if (start < 0) {
    throw "Invalid note found";
  }
  
  // TODO: why am I randomly adding length*5? Bizarre... Perhaps to ensure a very large positive value to be then modded?
  return noteArray[(start + N + (noteArray.length * 5)) % noteArray.length];
}

function replaceNotes(elem) {
  var rawText = elem.getText();
  var attributes = elem.getAttributes();
  var pieces = rawText.split(' ');
  var newPieces = [];
  pieces.forEach(function(piece) {  
    var match = piece.match(OVERALL_REGEX);
    if (match && match.length > 0) {
      newPieces.push(piece.replace(NOTE_REGEX, transpose));
    } else {
      newPieces.push(piece);
    }
  });
  elem.setText(newPieces.join(" "));
  elem.setAttributes(attributes);
  
  /*var rangeElem = elem.findText(REGEX);
  elem.replaceText("hello", "LMAOZEDONG");*/
}
