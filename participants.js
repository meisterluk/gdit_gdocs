//
// Participants script.
// Provides functionality to take participants and spread them across
// spreadsheets for each tutorial group.
//
//     :project:      gdit_gdocs
//     :author:       meisterluk
//     :date:         11.10.14
//     :version:      1.0.0public
//     :license:      LGPLv3
//

//----------------------------------------------------------------------
// GLOBAL CONFIGURATION
//----------------------------------------------------------------------

SOURCE_SHEET_STUD = 'Studenten';
SOURCE_SHEET_TUTS = 'Tutoren';

SOURCE_STUD_MATRNR_COL = 2;
SOURCE_TUTS_MATRNR_COL = 8;
SOURCE_STUD_TWNAME_COL = 5;
SOURCE_TUTS_TWNAME_COL = 7;

// inputBoxes do not allow default values. This var cannot be used.
//COPY_SPREADSHEET = "[GDIT] Gruppe N";
BASENAME_GROUP_SSS = '[GDIT] Gruppe %d';

// data structure to store students & tutors
data = {'students' : [], 'tutors' : {}, 'groups' : [], 'email' : {}};


//----------------------------------------------------------------------
// LIBRARY
//----------------------------------------------------------------------

//
// String trim function. Removes whitespace characters at beginning and
// end of string.
//
// @param string string   the string to trim
// @return string         the trimmed string
//
function trim(string)
{
    string = string.replace(/^\s+/, '');
    return string.replace(/\s+$/, '');
}

//
// Get the real content of a cell.
// Note. This function *always* returns strings.
//
// @param cell_value string  the cell value to parse
// @return string            the real content of the cell
//
function content(cell_value)
{
  if (typeof(cell_value) !== "string")
    return cell_value;
  return trim(cell_value.toString());
}

//
// A bad quick & dirty (probably not perfectly working) hack to check
// whether or not object is empty.
//
// @param obj object    the object to check for properties
// @return bool         is object empty?
//
function isEmpty(obj)
{
    for (var attr in obj) {
        if (obj.hasOwnProperty(attr))
            return false;
    }
    return true;
}

//
// Extend Array for membership testing method
// Note. Google Apps script does not allow overloading prototypes,
//       therefore this is a stupid function for prefix notation.
//
// @param parameter array   the array
// @param value object      some element
// @return bool             does value exist in parameter?
//
function contains(parameter, value)
{
  // some evil test for array property
  if (parameter instanceof Array)
  {
    for (var i=0; i<parameter.length; i++)
      if (parameter[i] == value)
        return true;
  } else { // assumption, it's an object
    for (var index in parameter)
      if (index == value)
        return true;
  }
  return false;
}

//
// Returns the corresponding letter (A1 notation) for a column index.
// Note. Currently A--ZZ is supported.
// Example:
//     >>> columnId(3)
//     "C"
//
// @param column_index int    the column index to use
// @return string             the column ID in A1 notation
//
function columnId(column_index)
{
  if (column_index > 0 && column_index < 27)
    return String.fromCharCode(64 + column_index);
  else if (column_index > 26 && column_index < 703)
  {
    column_index -= 26;
    var first_letter = 1;
    var second_letter = 0;
    for (; column_index>26; column_index-=26)
      first_letter += 1;
    second_letter = column_index;
    return String.fromCharCode(64 + first_letter)
         + String.fromCharCode(64 + second_letter);
  } else
    return String.fromCharCode(65);
}

//----------------------------------------------------------------------
// FUNCTIONALITY
//----------------------------------------------------------------------

//
// Remove "Copy of" label in sheet names when they get duplicated.
//
// @param sheetname string    the sheet name
// @return string the sheetname without "Copy of" text
//
function getOriginalSheetname(sheetname)
{
  return sheetname.replace(/^(Copy of|Kopie von)\s?/, "");
}

//
// Insert group_id into basename for Evaluation Spreadsheets.
//
// @param basename string    the basename to use
// @param group_id int       the group_id to use
// @return string specify spreadsheet name
//
function specifyGroup(basename, group_id)
{
  return basename.toString().replace(/%d/, group_id.toString());
}

//
// Because spreadsheets are identified by ID and not name, multiple
// spreadsheets can have the same name. Therefore GDocs does not provide
// a getSpreadsheetByName function for SpreadsheetApp. However, here we
// have an evil hack to use this functionality anyway.
//
// @param name string     the name to search for
// @return null|Spreadsheet null on error or Spreadsheet instance
//
function getSpreadsheetByName(name)
{
  var docs = new Array();
  docs = DocsList.getFilesByType("spreadsheet");
  for (var doc in docs)
  {
    if (docs[doc].getName() === name)
    {
      // docs[doc] contains File instance
      var id = docs[doc].getId();
      // get Spreadsheet instance
      return SpreadsheetApp.openById(id);
    }
  }
  return null;
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = sheet.getDataRange();
  var identifier = Browser.inputBox("Suche nach Matrnr: ");
  var found = false;

  if (identifier === "cancel")
    return;

  identifier = parseInt(identifier.toString().replace(/[^0-9]/g, ''));
  if (identifier === "cancel") {
    return true;
  } else if (isNaN(identifier)) {
    Browser.msgBox("Es wurde keine Zahl eingegeben");
    return searchMatrNr();
  } else {
    var base_spreadsheet = sheet;
    
    // for students sheet
    var sheet = ss.getSheetByName(SOURCE_SHEET_STUD);
    if (sheet === null)
    {
      Browser.msgBox("Studenten sheet konnte leider nicht gefunden "
                   + "werden :-(");
      return false;
    }
    range = sheet.getDataRange();

    for (var row=5; row<=range.getLastRow(); row++)
    {
      var val = range.getCell(row, SOURCE_STUD_MATRNR_COL).getValue();
      val = parseInt(val.toString().replace(/[^0-9]/g, ''));
      if (!isNaN(val) && numberSimilar(val, identifier))
      {
        sheet.setActiveCell(range.getCell(row, SOURCE_STUD_MATRNR_COL));
        found = true;
        break;
      }
    }

    // for tutors sheet
    var sheet = ss.getSheetByName(SOURCE_SHEET_TUTS);
    if (sheet === null)
    {
      Browser.msgBox("Tutoren sheet konnte leider nicht gefunden "
                   + "werden :-(");
      ss.setActiveSheet(base_spreadsheet);
      return false;
    }
    range = sheet.getDataRange();

    for (var row=5; row<=range.getLastRow(); row++)
    {
      var val = range.getCell(row, SOURCE_TUTS_MATRNR_COL).getValue();
      val = parseInt(val.toString().replace(/[^0-9]/g, ''));
      if (!isNaN(val) && numberSimilar(val, identifier))
      {
        sheet.setActiveCell(range.getCell(row, SOURCE_TUTS_MATRNR_COL));
        found = true;
        break;
      }
    }
  }
  if (!found)
  {
    ss.setActiveSheet(base_spreadsheet);
    Browser.msgBox("Nicht gefunden! :-(");
  }
}

//
// Search for a student by Matr.nr. or TwikiName
// Note. Calls an InputBox to request parameter
// Note. Points focus to search result or shows error message.
//
function searchTwikiName()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = sheet.getDataRange();
  var identifier = Browser.inputBox("Suche nach TwikiName: ");
  var found = false;
  if (identifier === "cancel")
    return;

  identifier = identifier.toString().replace(/\s+/, '').toLowerCase();
  var base_spreadsheet = sheet;
    
  // for students sheet
  var sheet = ss.getSheetByName(SOURCE_SHEET_STUD);
  if (sheet === null)
  {
    Browser.msgBox("Studenten Sheet konnte leider nicht gefunden "
                 + "werden :-(");
    return false;
  }
  range = sheet.getDataRange();

  for (var row=5; row<=range.getLastRow(); row++)
  {
    var val = range.getCell(row, SOURCE_STUD_TWNAME_COL).getValue();
    val = val.toString().replace(/\s+/,'').toLowerCase();
    if (stringSimilar(val, identifier))
    {
      sheet.setActiveCell(range.getCell(row, SOURCE_STUD_TWNAME_COL));
      found = true;
      break;
    }
  }

  // for tutors sheet
  var sheet = ss.getSheetByName(SOURCE_SHEET_TUTS);
  if (sheet === null)
  {
    Browser.msgBox("Tutoren sheet konnte leider nicht gefunden "
                 + "werden :-(");
    ss.setActiveSheet(base_spreadsheet);
    return false;
  }
  range = sheet.getDataRange();
  
  for (var row=5; row<=range.getLastRow(); row++)
  {
    var val = range.getCell(row, SOURCE_TUTS_TWNAME_COL).getValue();
    val = val.toString().replace(/\s+/,'').toLowerCase();
    if (stringSimilar(val, identifier))
    {
      sheet.setActiveCell(range.getCell(row, SOURCE_TUTS_TWNAME_COL));
      found = true;
      break;
    }
  }

  if (!found)
    Browser.msgBox("Nicht gefunden! :-(");
}

//
// Domain-specific function to create query to calculate total points.
//
// @param sheet Sheet     the Sheet to read data from
// @param spec_col int    the index of the column with points
// @param col int         the index of the column we are creating for
// @param start int       the first row to look for data
// @param end int         the last row to look for data
// @param bonus int       the row ID with the bonus value
// @return string the formula query to calculate total points
//
function createTpointsQuery(sheet, spec_col, col, start, end, bonus)
{
  try {
    // CTQ_ROWS_CACHE holds all row indizes with all points for the
    // various exercises.
    CTQ_ROWS_CACHE;
    if (CTQ_SHEET !== sheet.getName())
      throw("Is different sheet. Update cache, please.");
  } catch (err) {
    // if CTQ_ROWS_CACHE is undefined
    CTQ_ROWS_CACHE = [];
    var range = sheet.getRange(1, 1, end, spec_col);
    for (var row=start; row<=end; row++)
    {
      if (typeof(range.getCell(row, spec_col).getValue()) === "number")
      {
        CTQ_ROWS_CACHE.push(row);
      }
    }
    CTQ_SHEET = sheet.getName();
  }
  var rows = [];
  for (var index in CTQ_ROWS_CACHE)
  {
    var row = CTQ_ROWS_CACHE[index];
    rows.push("IF(" + columnId(col) + row + '="x",1,0)*$'
              + columnId(spec_col) + "$" + row);
  }
  rows.push(columnId(col) + bonus);
  return '=SUM(' + rows.join(', ') + ')';
}


//----------------------------------------------------------------------
// CALLBACK
//----------------------------------------------------------------------

//
// Generate callback to invoke generation of spreadsheets per group.
function generate()
{
  var base = Browser.inputBox("Name des Spreadsheets von dem das "
                            + "Layout kopiert wird: ");
  var ss = getSpreadsheetByName(base);
  if (ss === null)
  {
    Browser.msgBox("Konnte Spreadsheet nicht finden :-(");
    return false;
  }
  
  readData();
  generateSpreadsheets(ss);
}

//
// Read data to be used in spreadsheet generation
// Note. Modifies the global data variable.
//
function readData()
{
  var STUDS_SHEET = "Studenten";
  var TUTS_SHEET = "Tutoren";
  
  var group_col = 1;
  var matrnr_col = 2;
  var fname_col = 3;
  var sname_col = 4;
  var tname_col = 5;
  
  var base_path = [3, 2];
  var first_students_row = 5;
  var first_tutors_row = 4;

  var tut_group_col = 1;
  var tut_fname_col = 7;
  var tut_sname_col = 6;
  var tut_mail_col = 12;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(STUDS_SHEET);
  var range = sheet.getDataRange();
  
  // read data[meta]
  data['meta'] = [];
  for (var col=2; col<5; col++)
  {
    var value = content(range.getCell(1, col).getValue());
    if (value)
      data['meta'].push(value);
  }
  
  // read data[base]
  data['base'] = content(range.getCell(base_path[0], base_path[1])
                              .getValue());
  
  // read students & groups
  data['students'] = [];
  var start = Math.min(group_col, matrnr_col, fname_col, sname_col,
                       tname_col);
  var end = Math.max(group_col, matrnr_col, fname_col, sname_col,
                     tname_col);
  for (var row=first_students_row; row<=range.getLastRow(); row++)
  {
    // skip empty lines
    if (content(range.getCell(row, matrnr_col).getValue()) === "")
      continue;
    
    var student = {};
    for (var col=start; col<=end; col++)
    {
      var value = content(range.getCell(row, col).getValue());
      switch (col)
      {
        case group_col:
          student['group'] = parseInt(value);
          break;
        case matrnr_col:
          student['matrnr'] = parseInt(value);
          break;
        case fname_col:
          student['name'] = value;
          break;
        case sname_col:
          student['second_name'] = value;
          break;
        case tname_col:
          student['twiki_name'] = value;
          break;
        default:
          // skip it
      }
    }
    if (!contains(data['groups'], student['group']))
      data['groups'].push(student['group']);
    if (!isEmpty(student))
      data['students'].push(student);
  }

  // read tutors
  var sheet_tuts = ss.getSheetByName(TUTS_SHEET);
  var range = sheet_tuts.getDataRange();
  
  data['tutors'] = {};
  data['email'] = {};
  for (var row=first_tutors_row; row<=range.getLastRow(); row++)
  {
    var group = content(range.getCell(row, tut_group_col).getValue());
    var fname = content(range.getCell(row, tut_fname_col).getValue());
    var sname = content(range.getCell(row, tut_sname_col).getValue());
    var email = content(range.getCell(row, tut_mail_col).getValue());
    
    if (group === "" || typeof(group) !== "number")
      continue; // skip empty line
    
    data['tutors'][group] = sname + " " + fname;
    data['email'][group] = email;
  }

  // remove groups without tutor
  var limit = 3; // show 3 messages at maximum.
  var groups = []; // a whitelist
  for (var _ in data['groups'])
  {
    var group = data['groups'][_];
    
    if (data['tutors'][group] === undefined)
    {
      if (limit-- > 0)
        Browser.msgBox("Gruppe " + group + " hat keinen Tutor. Ein "
                   + "Spreadsheet für diese Gruppe wird nicht erzeugt");
      Logger.log("Missing tutor for group #" + group);
    } else {
      groups.push(group);
    }
  }
  data['groups'] = groups;
  
  // Checks
  if (data['groups'].length === 0)
    Browser.msgBox("Ich konnte gar keine Gruppen finden?!");
  if (data['students'].length === 0)
    Browser.msgBox("Ich konnte keine Studenten finden :-(");
  
  return true;
}

function generateSpreadsheets(base_ss)
{
  // TODO: These fields require a better usability design
  
  // Benotung Spreadsheet
  var ben_first_student_row = 5;
  
  var ben_group_col       = 1;
  var ben_matrnr_col      = 2;
  var ben_name_col        = 3;
  var ben_sname_col       = 4;
  var ben_tname_col       = 5;

  var ben_first_ex_col    = 10;
  var ben_last_ex_col     = 12;
  
  var ben_totalpoints_col = 14;
  var ben_mark_col        = 15;
  
  // Exercise(s) Spreadsheet
  var total_points        = [[48, 4], [50, 4], [80, 4]];
  var bonus               = [49, 51, 81];
  var write               = [45, 47, 77];

  var spec_col            = 3;
  var first_student_col   = 4;
  var twiki_name_row      = 2;
  var martnr_row          = 3;

  for (var g_index in data['groups'])
  {
    var group = data['groups'][g_index];
    var ss_name = specifyGroup(BASENAME_GROUP_SSS, group);
    var ss = SpreadsheetApp.create(ss_name);

    var rights_config = {editorAccess: true, emailInvitations: true};
    ss.addCollaborators(data['email'][group], rights_config);
    
    // get sheets & - names
    var sheets = base_ss.getSheets();
    var sheet_names = [];
    for (var s_index in sheets)
      sheet_names.push(sheets[s_index].getName());
    
    for (var index in sheets)
    {
      var sheet = sheets[index].copyTo(ss);
      sheet.setName(getOriginalSheetname(sheet.getName()));

      // write metadata
      var range = sheet.getRange(1, 1, 1, data['meta'].length+1);
      range.getCell(1, 1).setValue(data['tutors'][group]);
      for (var index_ in data['meta'])
        range.getCell(1, parseInt(index_)+1)
             .setValue(data['meta'][parseInt(index_)]);
      
      if (content(sheet.getName()) === "Benotung")
      {
        // write Benotungssheet data
        var range = sheet.getRange(1, 1, 
               data['students'].length + ben_first_student_row, 17);
      
        var gs_counter = 0; // student in group counter
        for (var s_index in data['students'])
        {
          if (data['students'][s_index]['group'] !== group)
            continue;

          var row = gs_counter + ben_first_student_row;
          range.getCell(row, ben_group_col)
               .setValue(data['students'][s_index]['group']);
          range.getCell(row, ben_matrnr_col)
               .setValue(data['students'][s_index]['matrnr']);
          range.getCell(row, ben_name_col)
               .setValue(data['students'][s_index]['name']);
          range.getCell(row, ben_sname_col)
               .setValue(data['students'][s_index]['second_name']);
          range.getCell(row, ben_tname_col)
               .setValue(data['students'][s_index]['twiki_name']);

          for (var col=ben_first_ex_col; col<=ben_last_ex_col; col++)
          {
            var sheet_ref = sheet_names[col - ben_first_ex_col + 1];
            var ex_ref    = col - ben_first_ex_col;
            var reference = "'" + sheet_ref + "'!"
                          + columnId(total_points[ex_ref][1]
                          + gs_counter)
                          + write[ex_ref];
            range.getCell(row, col).setFormula("=IF(" + reference
                    + "<0,0," + reference + ")");
          }
          
          range.getCell(row, ben_totalpoints_col)
               .setFormula("=SUM(" + columnId(ben_first_ex_col) + row
                        + ":" + columnId(ben_last_ex_col) + row + ")");
          range.getCell(row, ben_mark_col)
               .setFormula("=IF(ISERR(pointsToMark("
                  + columnId(ben_totalpoints_col) + col + ')), "?", '
                  + 'pointsToMark(' + columnId(ben_totalpoints_col)
                  + col + '))');
          gs_counter++;
        }
      } else {
        // write exercisesheet data
        var range = sheet.getRange(1, 1, sheet.getLastRow(),
                    first_student_col + data['students'].length - 1);

        var gs_counter = 0; // student in group counter
        for (var s_index in data['students'])
        {
          if (data['students'][s_index]['group'] !== group)
            continue;
          
          var col = first_student_col + gs_counter;
          var ref_t = "'Benotung'!" + columnId(ben_tname_col)
                    + (ben_first_student_row + gs_counter);
          var ref_m = "'Benotung'!" + columnId(ben_matrnr_col)
                    + (ben_first_student_row + gs_counter);

          range.getCell(twiki_name_row, col).setFormula
            ('=HYPERLINK(CONCATENATE("' + data['base'] + '", ' + ref_t +
             '), ' + ref_t + ')');
          range.getCell(martnr_row, col).setFontWeight("bold");
          range.getCell(martnr_row, col).setFormula('=' + ref_m);
          
          var exercise_number = (parseInt(index) - 1);
          var start = Math.max(twiki_name_row, martnr_row);
          var tpts = total_points[exercise_number];
          var formula = createTpointsQuery(sheet, spec_col, col, start,
                                   tpts[0], bonus[exercise_number]);
          range.getCell(write[exercise_number],
                        tpts[1] + gs_counter).setFormula(formula);
          gs_counter++;
        }
      }
      // delete default sheet
      var default_sheet = ss.getSheetByName("Sheet1");
      if (default_sheet !== null)
      {
        ss.setActiveSheet(default_sheet);
        ss.deleteActiveSheet();
      }
    }

    Browser.msgBox("Spreadsheet für Gruppe " + group + " erzeugt");
  }
  Browser.msgBox("Spreadsheets erzeugt :)");
}

//
// Trigger on document opening
//
function onOpen()
{
  // Add a search menu.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var menuEntriesGenerate = [
    {name: "Generiere Bewertungsspreadsheet pro Gruppe",
     functionName: "generate"}
  ];
  var menuEntriesSearch = [
    {name: "Nach Matrikelnummer", functionName: "searchMatrNr"},
    {name: "Nach Twikiname",      functionName: "searchTwikiName"}
  ];
  ss.addMenu("Generieren", menuEntriesGenerate);
  ss.addMenu("Suche", menuEntriesSearch);
}
