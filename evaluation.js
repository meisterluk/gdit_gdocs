//
// Evaluation script.
// Script to provide tools for the evaluation spreadsheet.
// Provides a "Search" menu item.
//
//     :project:      gdit_gdocs
//     :author:       meisterluk
//     :date:         11.10.03
//     :version:      0.3.0beta
//     :license:      GPLv3
//

//----------------------------------------------------------------------
// GLOBAL CONFIGURATION
//----------------------------------------------------------------------

MATRNR_ROW    = 3; // row with matr.numbers
NAME_ROW      = 2; // row with twiki names

//----------------------------------------------------------------------
// FUNCTIONALITY
//----------------------------------------------------------------------

//
// Retrieve a mark by total points of a student.
// Note. Has *always* to return an integer.
//
// @param points int    total points to calculate mark for
// @return int          the corresponding mark
//
function pointsToMark(points)
{
  if (points >= 88)
    return 1;
  else if (points >= 76)
    return 2;
  else if (points >= 63)
    return 3;
  else if (points >= 51)
    return 4;
  else
    return 5;
}

//
// Levenshtein distance calculation
// via phpjs
//     http://phpjs.org/functions/levenshtein:463
//
// @param s1 string    first string
// @param s2 string    second string to calculate similarity of
// @return int
//
function levenshtein(s1, s2)
{
  if (s1 == s2) {
    return 0;
  }

  var s1_len = s1.length;
  var s2_len = s2.length;
  if (s1_len === 0) {
    return s2_len;
  }
  if (s2_len === 0) {
    return s1_len;
  }

  var split = false;
  try {
    split = !('0')[0];
  } catch (e) {
    split = true; // Earlier IE may not support access by string index
  }
  if (split) {
    s1 = s1.split('');
    s2 = s2.split('');
  }

  var v0 = new Array(s1_len + 1);
  var v1 = new Array(s1_len + 1);

  var s1_idx = 0,
  s2_idx = 0,
  cost = 0;
  for (s1_idx = 0; s1_idx < s1_len + 1; s1_idx++) {
    v0[s1_idx] = s1_idx;
  }
  var char_s1 = '',
  char_s2 = '';
  for (s2_idx = 1; s2_idx <= s2_len; s2_idx++) {
      v1[0] = s2_idx;
      char_s2 = s2[s2_idx - 1];

      for (s1_idx = 0; s1_idx < s1_len; s1_idx++) {
        char_s1 = s1[s1_idx];
      cost = (char_s1 == char_s2) ? 0 : 1;
      var m_min = v0[s1_idx + 1] + 1;
      var b = v1[s1_idx] + 1;
      var c = v0[s1_idx] + cost;
      if (b < m_min) {
        m_min = b;
      }
      if (c < m_min) {
        m_min = c;
      }
      v1[s1_idx + 1] = m_min;
    }
    var v_tmp = v0;
    v0 = v1;
    v1 = v_tmp;
  }
  return v0[s1_len];
}

//
// String1 looks like string2?
//
// @param string1 string  first string to compare with second
// @param string2 string  second string
// @return bool shall I consider string1 to be similar to string2?
//
function stringSimilar(string1, string2)
{
  var sensitivity = 3; // levensthein distance sensivitity
  return levenshtein(string1, string2) < sensitivity;
}

//
// Number1 looks like number2?
// Note. levenshtein for numbers :-P
//
// @param number1 int   first number to compare with second
// @param number2 int   second number
// @return bool shall I consider number1 to be similar to number2?
//
function numberSimilar(number1, number2)
{
  if (number1 === number2)
    return true;

  /*if(number1- (Math.pow(10, 4) * parseInt(number1/10000)) === number2)
    return true;*/
  
  if (number1 > 1000)
  {
    var num1 = String(number1);
    var num2 = String(number2);
    
    if (num1.length < num2.length)
    {
      var tmp = num1;
      num1 = num2;
      num2 = tmp;
    }

    for (var i=-num1.length+1; i<num1.length; i++)
    {
      if (i > 0)
      {
        if (num1.substr(0, i).length > 3 && num1.substr(0, i) === num2)
          return true;
      } else if (i < 0) {
        if (num1.substr(i).length > 3 && num1.substr(i) === num2)
          return true;
      }
    }
  }
  
  return false;
}

//
// Search for a student by Matriculation number
// Note. Calls an InputBox to request parameter
// Note. Points focus to search result or shows error message.
//
function searchMatrNr()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  var identifier = Browser.inputBox
                   ("Suche nach Matrnr im aktuellen Sheet: ");
  var found = false;
  if (identifier === "cancel")
    return true;

  identifier = parseInt(identifier.toString().replace(/[^0-9]/g, ""));
  if (isNaN(identifier)) {
    Browser.msgBox("Es wurde keine Zahl eingegeben");
    return searchMatrNr();
  } else {
    if (sheet.getName() === "Benotung")
    {
      for (var row=5; row<=range.getLastRow(); row++)
      {
        var val = range.getCell(row, 2).getValue();
        val = parseInt(val.toString().replace(/[^0-9]/g, ''));
        if (!isNaN(val) && numberSimilar(val, identifier))
        {
          sheet.setActiveCell(range.getCell(row, 2));
          found = true;
          break;
        }
      }
    } else {
      for (var col=1; col<=range.getLastColumn(); col++)
      {
        var val = range.getCell(MATRNR_ROW, col).getValue();
        val = parseInt(val.toString().replace(/[^0-9]/g, ''));
        if (!isNaN(val) && numberSimilar(val, identifier))
        {
          sheet.setActiveCell(range.getCell
                  (range.getActiveCell.getRow(), col));
          found = true;
          break;
        }
      }
    }
  }
  if (!found)
    Browser.msgBox("Nicht gefunden! :-(");
}

//
// Search for a student by Matr.nr. or TwikiName
// Note. Calls an InputBox to request parameter
// Note. Points focus to search result or shows error message.
//
function searchTwikiName()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  var identifier = Browser.inputBox
                   ("Suche nach TwikiName im aktuellen Sheet: ");
  var found = false;
  if (identifier === "cancel")
    return;

  identifier = identifier.toString().replace(/\s+/, '').toLowerCase();
  if (sheet.getName() === "Benotung")
  {
    for (var row=5; row<=range.getLastRow(); row++)
    {
      var val = range.getCell(row, 5).getValue();
      val = val.toString().replace(/\s+/,'').toLowerCase();
      if (stringSimilar(val, identifier))
      {
        sheet.setActiveCell(range.getCell(row, 5));
        found = true;
        break;
      }
    }
  } else {
    for (var col=1; col<=range.getLastColumn(); col++)
    {
      val = range.getCell(NAME_ROW, col).getValue();
      val = val.toString().replace(/\s+/,'').toLowerCase();
      if (stringSimilar(val, identifier))
      {
        sheet.setActiveCell(range.getCell(HIGHLIGHT_ROW, col));
        found = true;
        break;
      }
    }
  }
  if (!found)
    Browser.msgBox("Nicht gefunden! :-(");
}

//----------------------------------------------------------------------
// CALLBACK
//----------------------------------------------------------------------

//
// Trigger on document opening
//
function onOpen()
{
  // Add a search menu.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var menuEntries = [
    {name: "Nach Matrikelnummer", functionName: "searchMatrNr"},
    {name: "Nach Twikiname",      functionName: "searchTwikiName"}
  ];
  ss.addMenu("Suche", menuEntries);
}
