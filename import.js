//
// Import script.
// Provides functionality to import data from TUG Online CSV dump
// and create a "[GDI] Teilnehmer" spreadsheet.
//
//     :project:      gdit_gdocs
//     :author:       meisterluk
//     :date:         11.10.03
//     :version:      0.3.0beta
//     :license:      GPLv3
//
// Table of Contents:
//     1. Global Configuration
//     2. Library
//     3. Functionality
//     4. Callbacks
//

//----------------------------------------------------------------------
// GLOBAL CONFIGURATION
//----------------------------------------------------------------------

IMPORT_COL               = 1;
IMPORT_ROW               = 1;
PLACEHOLDER              = '%s';
MENU_ITEM_SET            = "1. Parameter setzen";
MENU_ITEM_CHECK          = "(2.) Daten überprüfen";
MENU_ITEM_IMPORT         = "3. Daten importieren";
USER_BASEURL_DEFAULT  = 'http://gdi.ist.tugraz.at/gdi11/bin/view/Main/';
SSHEET_NAME_DEFAULT      = '[GDI] Teilnehmer';
IMPORT_SHEETNAME_DEFAULT = ''; // if empty, filled with active sheetname
_E_NO = "Nur zur Information: Die Anzahl der %s ist "
      + "0. Ich hoffe das ist korrekt.";
E_NO_TUTORS = _E_NO.replace("%s", "Tutoren");
E_NO_STUDENTS = _E_NO.replace("%s", "Studenten");
E_START = "Ich konnte keine Tabelle entdecken. Abbruch. :-(";
        + "Ich werde es trotzdem als Feldname akzeptieren.";
//E_NONEXISTENT_FIELD = "Zelle (%d, %d) [%s] enthält keinen bekannten "
//                    + " source_key. Handelt es sich wirklich um "
//                    + "TUG Online Quelldaten?";
E_FIELD_NOT_FOUND = "Ich werde leider abbrechen müssen, da die "
                  + "folgenden [essenziellen] Felder nicht gefunden "
                  + "werden konnten: ";
E_GMAIL = "Bei %s scheint es sich um keine gmail Emailadresse zu "
        + "handeln. Dies ist jedoch notwendig um die Berechtigungen "
        + "zu setzen.";
E_EMPTY_SHEET = "Es sind keine Daten vorhanden.";

// global data structure
data = {};

// field names to read
fields_config = {
    // key: uppercase, trimmed name of the column
    // value: two value tuple:
    //        [0] the field ID for internal usage (you don't want to change it)
    //        [1] if true, the value is assumed to be equal for all participants
    //            if false, the value differs for each participants
    //            if null, the value will not be read and processed
    // Note. In current implementation, if given key is not defined in
    //       fields_config, it behaves like [null, null].
    '^(REG\\w+_NUM\\w+|MATR\\w+)$'    : ["matrnr", false],
    '^(CODE\\w+STUDY\\w+|STUDIUM)$'   : ["study", false],
    '(FAMILY|SEC(OND)?|NACH)\\w*NAME' : ["familyname", false],
    '(FIRST|FORE|VOR)\\w*NAME'        : ["name", false],
    'E?MAIL'                          : ["mail", false],
    '^(SEMESTER.*STUD(IUM|Y))$'       : ["st_semester", false],
    '(ASSESS|PRÜF|EXAM)'              : ["assessment", true],
    '(NUM|ID)\\w+(COURSE|LV)'         : ["courseID", true],
    '(SEM)\\w+(COURSE|LV)'            : ["semester", true],
    '(COURSE|LV)\\w+TYP'              : ["ctype", true],
    '(COURSE|LV)\\w+TIT(LE|EL)'       : ["title", true],
    '(EXAMINER|PROF)'                 : ["prof", true],
    '(GROUP|GRUPPE)'                  : ["gname", false]
};

//----------------------------------------------------------------------
// LIBRARY
//----------------------------------------------------------------------

//
// Notification object.
// 0 = note, 1 = warn, 2 = error
//
log = {
    volume : 1,
    msgs   : [],
    quiet  : function () { this.volume = 100; },
    silent : function () { this.volume = 1; },
    loud   : function () { this.volume = 0; },
    
    log    : function (msg) { this.msgs.push([msg, 0]); this.show(); },
    warn   : function (msg) { this.msgs.push([msg, 1]); this.show(); },
    error  : function (msg) { this.msgs.push([msg, 2]); this.show(); },
    
    show   : function () {
        var last = this.msgs[this.msgs.length - 1];
        if (last[1] >= this.volume)
          Browser.msgBox(last[0]);
    },
    all    : function () {
        for (var msg in this.msgs)
          if (msg[1] >= this.volume)
            Browser.msgBox(msg[0]);
    }
};

//
// Persistent Data Storage
// Note. Uses currently temporary sheets.
// Note. Stores data in JSON format.
//
pds = {
    name     : "Persistent Data Storage",
    tmpsheet : "gdit_gdocs temporary sheet",
    store    : function (data) {
      // store data persistently
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var old_sheet = ss.getActiveSheet();
      var sheet = ss.getSheetByName(this.tmpsheet);
      if (sheet === null)
        sheet = ss.insertSheet(this.tmpsheet);
      else
        sheet.clear();
      var range = sheet.getDataRange();
      var data = JSON.stringify(data);
      range.getCell(1, 1).setValue(data);
      var rs = ss.setActiveSheet(old_sheet);
    },
    load     : function () {
      // load data from temporary sheet
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var old_sheet = ss.getActiveSheet();
      var sheet = ss.getSheetByName(this.tmpsheet);
      if (sheet === null)
        return false;
      var range = sheet.getDataRange();
      var value = range.getCell(1, 1).getValue();
      ss.setActiveSheet(old_sheet);
      
      return JSON.parse(value);
    },
    close    : function () {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(this.tmpsheet);
      if (sheet !== null)
      {
        ss.setActiveSheet(sheet);
        ss.deleteActiveSheet();
      }
    }
};

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
// Create the abbreviation of a title.
// Note. Perfect for course titles:
// Example:
//     >>> abbreviation("Grundlagen der Informatik")
//     "GDI"
//
// @param param string    the phrase to create abbreviation for
// @return string         the abbreviation
//
function abbreviation(param)
{
  var parts = param.toString().split(/\W+/);
  var abbr = "";
  for (var key in parts)
    abbr += parts[key].charAt(0).toUpperCase();
  return abbr;
}

//
// Create a CamelCased Word based on input string
// eg.    "Karl-victor Woit"  =>  "KarlVictorWoit"
//
// @param word string   the word to process
// @return string       the camel cased version of the word
//
function camel_case(word)
{
  //   ~)
  //   (_/^\
  //    /|~|\
  //   / / / |
  var parts = word.toString().split(/\W+/);
  var cameled = "";
  for (var key in parts)
    cameled += parts[key].charAt(0).toUpperCase() +parts[key].substr(1);
  return cameled;
}

//
// Get a name and return the corresponding Twiki username.
//
// @param first_name string  a string including first name of the student
// @param last_name string   a string including last name of the student
// @return string            the TWiki username
//
function twiki_username(first_name, last_name)
{
  var name = first_name + last_name;
  
  name = name.replace("ä", "ae");
  name = name.replace("ö", "oe");
  name = name.replace("ü", "ue");
  name = name.replace("ß", "ss");
  name = name.replace("-", "");
  name = name.replace("é", "e");

  return camel_case(name);
}

//
// Get some string and compare it with the cell value.
// Note. This function defines how the user-defined value of a cell
//       is compared to a string.
// Note. This function defines a case-insensitive comparison.
//
// @param cell_value string  the cell value to compare with param
// @param string string      the param to compare with cell value
// @return bool              are cell_value and param equal?
//
function compare(cell_value, param)
{
  return trim(cell_value.toString().toUpperCase()) === param.toString();
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
// Check whether or not string ends with substr.
//
// @param string string    string to search substr in
// @param substr string    string to search for
// @return bool            does string end with substr?
//
function endswith(string, substr)
{
  return string.substr(-substr.length) === substr;
}

//
// The group identifier is stored as a string.
// We want to extract Array(group_id, datetime.day, datetime.start,
// datetime.end) from the string.
//
// @param string string    a string identifying tutorium datetime
//                         eg. "Gruppe: Gruppe  3, Montag 11-12  "
// @return object Array[group_id, day, start, end]
//                (here: [3, "Montag", 11, 12])
//
function groupSplit(string)
{
  var split = string.match(/Gr(uppe(n|nnr|nnummer)?|oup)\s+(\d+)[-,.]/);
  if (split)
    var group_id = parseInt(split[3]);
  else
    var group_id = 0;

  var split = string.match(/([MTWFSD]\w*)\s+(\d{1,2})-(\d{1,2})/);
  if (split)
  {
    var wday = split[1];
    var time_start = parseInt(split[2]);
    var time_end = parseInt(split[3]);
  } else {
    var wday = "?";
    var time_start = 0;
    var time_end = 0;
  }

  return [group_id, wday, time_start, time_end];
}

//
// Get semesterID by current date.
// Example:
//     >>> getCurrentSemesterId()
//     "WS 2011/12"
//
// @return string   semesterID
//
function getCurrentSemesterId()
{
  var today = new Date();
  if (today.getMonth() > 6)
    var semesterId = "WS " + today.getFullYear() + "/"
                   + (today.getFullYear() % 100);
  else
    var semesterId = "SS " + (today.getFullYear() + 1);
  return semesterId;
}

//----------------------------------------------------------------------
// FUNCTIONALITY
//----------------------------------------------------------------------

//
// Create a UserInterface to request parameters for creation of
// new spreadsheet. The user input is going to be stored in global
// variables in this file.
// Note. This function requires the callback submitParameters.
//
function requestParameters()
{
  autoloadConfig();

  var ass = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle('Parameter setzen');
  
  var INTRO = "Setze hier Parameter, welche zum Erzeugen des Teilnehmer"
            + "-spreadsheets verwendet werden:";
  
  var BASE = 'Basis URL (Twiki Username wird am Ende angefügt '
           + 'und soll einen validen Link zur Benutzerseite '
           + 'enthalten)';
  var SSHEET = 'Name des neuen Teilnehmer-Spreadsheets';
  var SHEET = 'Sheetname des Import Spreadsheets (vermutlich Name '
            + 'des aktuellen Sheets)';

  if (IMPORT_SHEETNAME_DEFAULT === '')
    IMPORT_SHEETNAME_DEFAULT = ass.getActiveSheet().getName();
  
  var grid = app.createGrid(3, 2);
  grid.setWidget(0, 0, app.createLabel(BASE));
  grid.setWidget(0, 1, app.createTextBox().setName('base_url')
                          .setText(USER_BASEURL_DEFAULT));
  grid.setWidget(1, 0, app.createLabel(SSHEET));
  grid.setWidget(1, 1, app.createTextBox().setName('ss_name')
                          .setText(SSHEET_NAME_DEFAULT));
  grid.setWidget(2, 0, app.createLabel(SHEET));
  grid.setWidget(2, 1, app.createTextBox().setName('import_sheet')
                          .setText(IMPORT_SHEETNAME_DEFAULT));
    
  var submit = app.createButton('Setzen');
  var handler = app.createServerClickHandler('submitParameters');
  handler.addCallbackElement(grid);
  submit.addClickHandler(handler);

  var panel = app.createVerticalPanel();
  panel.add(app.createLabel(INTRO));
  panel.add(grid);
  panel.add(submit);
  app.add(panel);
  ass.show(app);
}

//
// Read configuration from PDS.
//
function autoloadConfig()
{
  var config = pds.load();
  if (config !== false) {
    if (config['USER_BASEURL_DEFAULT'] !== undefined)
      USER_BASEURL_DEFAULT = config['USER_BASEURL_DEFAULT'];
    if (config['SSHEET_NAME_DEFAULT'] !== undefined)
      SSHEET_NAME_DEFAULT = config['SSHEET_NAME_DEFAULT'];
    if (config['IMPORT_SHEETNAME_DEFAULT'] !== undefined)
      IMPORT_SHEETNAME_DEFAULT = config['IMPORT_SHEETNAME_DEFAULT'];
  }
}

//
// data contains the data structure to create participants spreadsheet.
// Assuming it was already filled, we can do any checks for validity of
// the data structure.
//
// @return bool   indicating success or failure
//
function checkDatastructure()
{
  if (data['tutors'].length === 0)
    log.warn(E_NO_TUTORS);

  if (data['students'].length === 0)
    log.warn(E_NO_STUDENTS);

  for (var tutor_nr in data.tutors)
  {
    if (!endswith(data.tutors[tutor_nr]["mail"], "gmail.com"))
    {
      Logger.log(tutor_nr);
      log.warn(E_GMAIL.replace(/%s/, data.tutors[tutor_nr]["mail"]));
      break;
    }
  }

  return true;
}

//
// Create participants spreadsheet and write data accordingly.
//
function writeData()
{
  if (data['students'].length === 0)
  {
    Browser.msgBox("writeData() wurde aufgerufen, aber keine Daten "
             + "zum Verarbeiten sind vorhanden. Wurde readData() "
             + "vorher aufgerufen? Vielleicht konnten die gelesenen "
             + "Daten nicht gespeichert werden?!");
    return false;
  }

  autoloadConfig();
  // do enable, if you want the temporary storage sheet to be deleted,
  // if writeData is called
  //pds.close();

  // Create participants spreadsheet
  var ss = SpreadsheetApp.create(SSHEET_NAME_DEFAULT);
  var sheet = ss.getActiveSheet();
  
  sheet.setName("Studenten");
  
  var range = sheet.getRange(1, 1, data.students.length + 5, 6);
  var first_column = sheet.getRange(1, 1, 1, 6);
  
  var semesterId = getCurrentSemesterId();

  data["prof"] = data["prof"].replace(/(.*)Dipl\.-? ?Ing\.?$/, "DI $1");
  
  first_column.setFontSize(12);
  range.getCell(1, 1).setValue("Studenten");
  range.getCell(1, 2).setValue(abbreviation(data['title'])
                                + " (" + data['ctype'] + ")");
  range.getCell(1, 3).setValue(semesterId);
  range.getCell(1, 4).setValue(data["prof"]);
  
  range.getCell(3, 1).setValue("Basis URL:");
  sheet.getRange(3, 2, 1, 3).mergeAcross().setValue(USER_BASEURL_DEFAULT);
  
  sheet.getRange(4, 1, 1, 6).setFontWeight("bold");
  range.getCell(4, 1).setValue("Gruppe");
  range.getCell(4, 2).setValue("Matrikelnummer");
  range.getCell(4, 3).setValue("Vorname");
  range.getCell(4, 4).setValue("Nachname");
  range.getCell(4, 5).setValue("Twiki Name");
  range.getCell(4, 6).setValue("Link");

  var row_nr = 6;
  var grp = [];
  for (var student in data['students'])
  {
    var st_data = data['students'][student];
    grp = groupSplit(st_data['gname']);
    var user = twiki_username(st_data["name"], st_data["familyname"]);
    
    range.getCell(row_nr, 1).setValue(grp[0]);
    range.getCell(row_nr, 2).setValue(st_data["matrnr"]);
    range.getCell(row_nr, 3).setValue(st_data["name"]);
    range.getCell(row_nr, 4).setValue(st_data["familyname"]);
    range.getCell(row_nr, 5).setValue(user);
    range.getCell(row_nr, 6).setFormula('=HYPERLINK(CONCATENATE($B$3, E'
                                      + row_nr + '), E' + row_nr + ')');
    row_nr++;
  }
  
  sheet = ss.insertSheet("Tutoren");
  var range = sheet.getRange(1, 1, data.tutors.length + 4, 14);
  
  range.getCell(1, 1).setValue("Tutoren");
  sheet.getRange(1, 1, 1, 14).setFontSize(12);
  sheet.getRange(1, 2, 1, 2).mergeAcross().setValue(data['title'])
                            .setHorizontalAlignment("center");
  range.getCell(1, 4).setValue(semesterId);
  range.getCell(1, 5).setValue(data['ctype']);
  
  sheet.getRange(3, 1, 1, 14).setFontWeight("bold");
  range.getCell(3, 1).setValue("Gruppe #");
  sheet.getRange(3, 2, 1, 2).mergeAcross()
                            .setValue("Tutoriumszeit (von | bis)")
                            .setHorizontalAlignment("center");
  range.getCell(3, 4).setValue("Tutoriumstag");
  range.getCell(3, 5).setValue("Platz");
  range.getCell(3, 6).setValue("Nachname");
  range.getCell(3, 7).setValue("Vorname");
  range.getCell(3, 8).setValue("Matrikelnummer");
  range.getCell(3, 9).setValue("Studium");
  range.getCell(3, 10).setValue("Studiensemester");
  range.getCell(3, 11).setValue("Anmeldedatum");
  range.getCell(3, 12).setValue("Email");
  range.getCell(3, 13).setValue("Notiz");
  
  sheet.getRange(4, 3, data.students.length, 1)
       .setHorizontalAlignment("left");

  var row_nr = 5;
  grp = [];
  for (var tutor in data['tutors'])
  {
    var tu_data = data['tutors'][tutor];
    grp = groupSplit(tu_data["gname"]);
    var user = twiki_username(tu_data["name"], tu_data["familyname"]);

    range.getCell(row_nr, 1).setValue(grp[0]);
    range.getCell(row_nr, 2).setValue(grp[2] + ":00");
    range.getCell(row_nr, 3).setValue(grp[3] + ":00");
    range.getCell(row_nr, 4).setValue(grp[1]);
    range.getCell(row_nr, 5).setValue("unklar");
    range.getCell(row_nr, 6).setValue(tu_data["familyname"]);
    range.getCell(row_nr, 7).setValue(tu_data["name"]);
    range.getCell(row_nr, 8).setValue(tu_data["matrnr"]);
    range.getCell(row_nr, 9).setValue(tu_data["study"]);
    range.getCell(row_nr, 10).setValue(tu_data["st_semester"]);
    range.getCell(row_nr, 11).setValue(data["assessment"]);
    range.getCell(row_nr, 12).setValue(tu_data["mail"]);
    range.getCell(row_nr, 13).setValue("");
    
    row_nr++;
  }
  
  return true;
}

//
// Read data from source sheet to global storage object.
// Note. This function hardcodes the structure of the source data.
//
// @param sheet object  a Sheet instance to read
// @return bool         success (true) or failure (false)
//
function readData(sheet)
{
  var range = sheet.getDataRange();
  data['tutors'] = [];
  data['students'] = [];

  // defines where data retrieval will start
  var source_start_col  = IMPORT_COL;
  var source_start_row  = IMPORT_ROW;
  var source_end_col    = range.getLastColumn();
  var source_end_row    = range.getLastRow();

  if (source_end_row === 1)
  {
    log.error(E_EMPTY_SHEET);
    return false;
  }

  var found_start = false;
  for (var offset=0; offset<5; offset++)
  {
    var current = content(range.getCell(source_start_row + offset,
                  source_start_col).getValue());
    var next    = content(range.getCell(source_start_row + offset + 1,
                  source_start_col + 1).getValue());

    if (current != false && next != false)
    {
      source_start_row += offset;
      found_start = true;
      break;
    }
  }

  if (!found_start)
  {
    log.warn(E_START.replace(/%s/, source_start_col)
                    .replace(/%s/, source_start_row));
    return false;
  }

  // read data from tabular structure

  var active_column = [];
  var is_first_valid_column = true;
  var tmp_flag = true;

  // create a list of all fields, which should be filled up
  // for each student
  var missing_fields = {};
  for (var index in fields_config)
    if (fields_config[index][1] !== null)
      missing_fields[fields_config[index][0]] = true;

  for (var col=source_start_col; col<=source_end_col; col++)
  {
    var students_counter = 0;
    var tutors_counter = 0;
    active_column = [];
    rows_loop:
    for (var row=source_start_row; row<=source_end_row; row++)
    {
      var val = content(range.getCell(row, col).getValue());

      // evaluate type
      if (range.getCell(row, col).getFontWeight() === "bold")
        var type = 'tutors';
      else
        var type = 'students';

      // if is first row, expect field name
      if (row === source_start_row)
      {
        for (var regex in fields_config)
        {
          var re = new RegExp(regex, "i");
          if (val.match(re))
          {
            active_column = fields_config[regex];
            missing_fields[fields_config[regex][0]] = false;
            break;
          }
        }

        // if field name cannot be found, write error and skip column
        if (active_column.length === 0)
        {
          Logger.log("'" + val + "' is not a known TUG Online field");
          break rows_loop;
        } else {
          if (tmp_flag && is_first_valid_column)
            tmp_flag = false;
          else if (!tmp_flag && is_first_valid_column)
            is_first_valid_column = false;
        }
      } else {
        // if is not first row

        // if is first column, create a new student/tutor instance
        if (is_first_valid_column)
          data[type].push({});

        if (active_column[1] === true)
          data[active_column[0]] = val;
        else if (active_column[1] === false)
        {
          //data[type][data[type].length - 1][active_column[0]] = val;
          if (type === "students")
            data[type][students_counter++][active_column[0]] = val;
          else
            data[type][tutors_counter++][active_column[0]] = val;
        } else if (active_column[1] === null)
          break rows_loop;
        else
          log.error("Invalide fields_config Konfiguration");
      }
    }
  }

  // <hack why="because current CSV format does not supply enough information"
  //       what="hardcoding data" which="bad">
  if (!("assessment" in data))
    data["assessment"] = "16.12.2011";
  if (!("courseID" in data))
    data["courseID"] = "716.231";
  if (!("semester" in data))
    data["semester"] = "11W";
  if (!("ctype" in data))
    data["ctype"] = "UE";
  if (!("title" in data))
    data["title"] = "Grundlagen der Informatik";
  if (!("prof" in data))
    data["prof"] = "Vo"+"it K"+"ar"+"l Di"+"pl.-Ing.";
  for (var index in data.students)
  {
    if (!("study" in data['students'][index]))
      data['students'][index]['study'] = "F 033 000";
    if (!("st_semester" in data['students'][index]))
      data['students'][index]['st_semester'] = 0;
  }
  for (var index in data.tutors)
  {
    if (!("study" in data['tutors'][index]))
      data['tutors'][index]['study'] = "F 033 000";
    if (!("st_semester" in data['tutors'][index]))
      data['tutors'][index]['st_semester'] = 0;
  }

  var mark = ["assessment", "courseID", "semester", "ctype", "title",
              "prof", "study", "st_semester"];
  for (var index in mark)
    missing_fields[mark[index]] = false;
  // </hack>

  var missing_items = [];
  for (var index in missing_fields)
    if (missing_fields[index])
      missing_items.push(index);

  if (missing_items.length > 0)
  {
    log.error(E_FIELD_NOT_FOUND + missing_items.join(", "));
    return false;
  }
  Logger.log(data);

  // Okay, now we have the data stored in var data.

  return true;
}


//----------------------------------------------------------------------
// CALLBACKS
//----------------------------------------------------------------------

//
// Will be called when spreadsheet is opened.
//
function onOpen()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Add menu item for import functionality
  var menuEntries = [
    {name: MENU_ITEM_SET, functionName: "set"},
    {name: MENU_ITEM_CHECK, functionName: "check"},
    {name: MENU_ITEM_IMPORT, functionName: "import"},
    {name: "Hilfe", functionName: "help"}
  ];
  ss.addMenu("Import", menuEntries);
}

//
// Main routine to set parameters
//
function set()
{
  requestParameters();
}

//
// Main routine to do the check for validity of data.
//
function check()
{
  log.loud();
  var success = readData(SpreadsheetApp.getActiveSpreadsheet()
                                        .getActiveSheet());
  success = success && checkDatastructure();
  if (success)
    Browser.msgBox("Test bestanden :-)");
  else
    Browser.msgBox("Test nicht bestanden!");
}

//
// Main routine to import data.
//
function import()
{
  log.silent();
  var success = readData(SpreadsheetApp.getActiveSpreadsheet()
                                        .getActiveSheet());
  if (success)
    Browser.msgBox("Daten gelesen :-) Werde Spreadsheet erzeugen");
  else {
    Browser.msgBox("Abbruch.");
    return false;
  }
  checkDatastructure();
  if (writeData())
    Browser.msgBox("Spreadsheet '" + SSHEET_NAME_DEFAULT
                 + "' wurde erzeugt :-)");
  return true;
}

//
// Help context menu in MsgBox
//
function help()
{
  var ass = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle('Import Skript');
  
  var p1 = "Dieses Skript handhabt den Import der Daten vom "
         + "TUG Online CSV und erzeugt automatisiert den "
         + "'GDI Teilnehmer' Spreadsheet. Es stehen zwei Aufrufe "
         + "zur Verfügung: Einmal 'Import' > '" + MENU_ITEM_CHECK
         + "' zum Basischeck ob die Eingabedaten verarbeitet werden "
         + "können und zweitens 'Import' > '" + MENU_ITEM_IMPORT
         + "' zum wirklichen Importieren der Daten.";
  var p2 = "Wähle diese Menüpunkte, wenn du diese Funktionen nutzen "
         + "willst.";
  var panel = app.createVerticalPanel();
  panel.add(app.createLabel(p1));
  panel.add(app.createLabel(p2));
  app.add(panel);
  ass.show(app);
}

//
// Called by submit button of requestParameters UI
//
function submitParameters(e)
{
  // since this is not stored persistently ...
  USER_BASEURL_DEFAULT      = e.parameter.base_url;
  SSHEET_NAME_DEFAULT       = e.parameter.ss_name;
  IMPORT_SHEETNAME_DEFAULT  = e.parameter.import_sheet;
  // ... I will use pds:
  pds.store({
        USER_BASEURL_DEFAULT     : USER_BASEURL_DEFAULT,
        SSHEET_NAME_DEFAULT      : SSHEET_NAME_DEFAULT,
        IMPORT_SHEETNAME_DEFAULT : IMPORT_SHEETNAME_DEFAULT
  });
  
  // Clean up - get the UiApp object, close it, and return
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}
