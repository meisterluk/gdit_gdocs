//
// Statistics script.
// Provides functionality to create stats charts and export data to TUG
// Online CSV dump.
//
//     :project:      gdit_gdocs
//     :author:       meisterluk
//     :date:         11.10.14
//     :version:      1.0.0public
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

CELL11_TITLE = "Statistiken";
EVALUATION_MAX_POINTS = 200;
STATS_FIRST_ROW = 2;
STATS_FIRST_COL = 1;

MENU_ITEM_SET        = "1. Parameter setzen";
MENU_ITEM_READ       = "2. Daten lesen";
MENU_ITEM_CREATE     = "3. Erstelle Charts";
MENU_ITEM_SETEXPORT  = "4. Setze Export Parameter";
MENU_ITEM_WRITE      = "5. Exportiere ins TUG Online CSV";

SOURCE_MATRNR_COLUMN = 2;
SOURCE_TOTAL_POINTS_COLUMN = 14;
SOURCE_FIRSTEXERCISE_COLUMN = 10;
SOURCE_LASTEXERCISE_COLUMN = 12;

SOURCE_BASENAME = '[GDIT] Gruppe %d';
SOURCE_SHEETNAME = 'Benotung';
SOURCE_FIRST_GROUP = 1;
SOURCE_LAST_GROUP = 20;
SOURCE_MINPOINTS = 20;

IMPORT_SSNAME_DEFAULT = '[GDIT] Import';
IMPORT_SHEETNAME_DEFAULT = 'Sheet1';
IMPORT_FIELD_MATR = '^(REG\\w+_NUM\\w+|MATR\\w+)$';
IMPORT_FIELD_GRADE = 'GRADE';

CHART_WIDTH = 600;
CHART_HEIGHT = 400;
CHART_COLOR = "blue";

DEC = 2; // decimal points for numbers in spreadsheet

E_404 = "Spreadsheet '%s' konnte nicht gefunden werden";
E_READDATA = "Muss die Daten der Bewertungsspreadsheets neu einlesen "
           + "aber fand sie nicht. Wurden die Parameter geändert?";
E_EXPORT_NOTFOUND = "Spreadsheet '%s' mit dem Sheet '%s' konnte nicht "
                  + "gefunden werden";

// data structure to store data in
//    [students] => ([matrnr] => total_points)
//    [tutors]   => ([tutor] => (student1_matrnr, student2_matrnr, ...))
//    [exercises] => ([exercise] => ([matrnr] => points))
data = {'students' : {}, 'tutors' : {}, 'exercises' : {}};
add_info = {};

//----------------------------------------------------------------------
// LIBRARY
//----------------------------------------------------------------------

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
      else {
        old_data = this.load();
        if (!old_data)
          sheet.clear();
        else {
          data = mergeObjects(data, old_data);
        }
      }
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
// Merge properties of two objects. If property is defined in both
// objects, obj1 has proprity.
//
// @param obj1 object
// @param obj2 object
// @return obj an object with properties of obj1 and obj2
//
function mergeObjects(obj1, obj2)
{
  var obj = obj1;
  for (var attr in obj2)
    if (!(attr in obj))
      obj[attr] = obj2[attr];
  return obj1;
}

//
// String trim function. Removes whitespace characters at beginning and
// end of string.
//
// @param string string   the string to trim
// @return string         the trimmed string
//
function trim(string) {
    string = string.replace(/^\s+/, '');
    return string.replace(/\s+$/, '');
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
  //name = name.replace("-", ""); // TODO: I doubt, this is a good idea
  name = name.replace("é", "e");

  return camel_case(name);
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
// Take the tutor specification (typically Cell(1,1)) and return the
// *real* tutor part.
// Example:
//     >>> getRealTutorSpecification("Tutor FooBar")
//     "FooBar"
//
// @param tutor string    the tutor specification of the cell
// @return string the *real* tutor specification
//
function getRealTutorSpecification(tutor)
{
  // remove leading "Tutor "
  return tutor.toString().substr(6);  
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
// Get a list of numbers and return the average value.
//
// @param list array    a list of numbers
// @return float the average value
//
function getAverage(list)
{
  if (!list || list.length === 0)
    return 0.0;
  return list.reduce(function(a, b) { return a+b; }) / list.length;
}

//----------------------------------------------------------------------
// FUNCTIONALITY
//----------------------------------------------------------------------

//
// Return the number of students in data.
//
// @return int    the number of students (identified by matrnr) in data
//
function getNumberOfStudents()
{
  var counter = 0;
  for (var key in data['students'])
  {
    if (!isNaN(key))
      counter++;
  }
  return counter;
}

//
// Return the number of tutors in data.
//
// @return int    the number of tutors in data
//
function getNumberOfTutors()
{
  var counter = 0;
  for (var key in data['tutors'])
    counter++;
  return counter;
}

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
// Shorten string to a specific length.
// Add terminator if string is longer than limit.
// Note. Used, because titles of charts must not be too long.
//
// @param value string       the value to shorten
// @param limit int          the maximal length of the value
// @param terminator int     the terminator to append if too long
// @return string the shortened value
//
function shortenString(value, limit, terminator)
{
  if (terminator === undefined)
    terminator = ' ...';

  value = value.toString();
  if (value.length > limit)
    value = value.substr(0, limit - 4) + terminator.toString();
  return value;
}

//
// Domain-specific function to get a range with 2 columns and n rows.
// Will return the content as object.
//
//        col_1              col_2
//   +=============+=====================+
//   | Tutoren     | Durchschnittsnote   |   start_row
//   +-------------+---------------------+
//   | Lukas       |        3.5          |
//   | Thomas      |        4.2          |    end_row
//   +=============+=====================+
//
//    =>
//       {"topic" : "Tutoren", "field" : "Durchschnittsnote",
//        "Lukas" : 3.5, "Thomas" : 4.5}
//
// @param col_1 int      the id of the column with keys
// @param col_2 int      the id of the column with values
// @param start_row int  the id of the row with metadata
// @param end_row int    the id of the row to terminate parsing
// @return object|false  stats object or false on error
//
function readStat(col_1, col_2, start_row, end_row)
{
  if (col_1 > col_2 || end_row < start_row)
    return false;

  // +1 (in cell referencing) because one-based
  var ass = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ass.getActiveSheet();
  var range = sheet.getRange(start_row, col_1, end_row - start_row + 1,
                             col_2 - col_1 + 1);
  var first_col = 1;
  var last_col = col_2 - col_1 + 1;
  var first_row = 1;
  var last_row = end_row - start_row + 1;

  var topic = content(range.getCell(first_row, first_col).getValue());
  var field = content(range.getCell(first_row, last_col).getValue());
  
  var stats = { topic : topic, field : field };
  
  for (var row=first_row+1; row<=last_row; row++)
  {
    var key   = content(range.getCell(row, first_col).getValue());
    var value = content(range.getCell(row, last_col).getValue());
    
    // key must not be empty
    if (key === "")
      return false;
    
    stats[key] = value;
  }
  
  return stats;
}

//
// Create a UserInterface to request parameters for reading stats. The
// user input is going to be stored in global variables in this file.
// Note. This function requires the callback submitStatsParameters.
//
function requestStatsParameters()
{
  autoloadConfig();

  var ass = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle('Parameter setzen');
  
  var INTRO = "Setze die Parameter, um die Statistiken zu erzeugen:";
  
  var BASE = 'Name der Bewertungsspreadsheets (%d wird durch eine '
           + 'fortlaufende Nummer ersetzt)';
  var RANGE = 'Gebe den Anfang und das Ende der fortlaufenden Nummer '
            + 'an';
  var BOUND = 'Gebe eine Mindestpunktezahl an ab der die Studierenden '
            + ' in der Statistik gezählt werden (0 für "zähle alle")';

  var default_range = SOURCE_FIRST_GROUP + "-" + SOURCE_LAST_GROUP;
  var grid = app.createGrid(3, 2);
  grid.setWidget(0, 0, app.createLabel(BASE));
  grid.setWidget(0, 1, app.createTextBox().setName('basename')
                          .setText(SOURCE_BASENAME));
  grid.setWidget(1, 0, app.createLabel(RANGE));
  grid.setWidget(1, 1, app.createTextBox().setName('range')
                          .setText(default_range));
  grid.setWidget(2, 0, app.createLabel(BOUND));
  grid.setWidget(2, 1, app.createTextBox().setName('minpoints')
                          .setText(SOURCE_MINPOINTS.toString()));
    
  var submit = app.createButton('Setzen');
  var handler = app.createServerClickHandler('submitStatsParameters');
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
// Create a UserInterface to request parameters for exporting data to
// TUG Online sheet. The user input is going to be stored in global
// variables in this file.
// Note. This function requires the callback submitExportParameters.
//
function requestExportParameters()
{
  autoloadConfig();

  var ass = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle('Parameter setzen');

  var SSNAME = 'Name des Spreadsheets in den geschrieben werden soll';
  var MATR = 'RegEx ohne Groß-Kleinschreibung-Unterscheidung zum '
           + 'Erkennen der Matrikelnummerspalte';
  var GRADE = 'RegEx ohne Groß-Kleinschreibung-Unterscheidung zum '
            + 'Erkennen der Spalte für Noteneintragung';

  var grid = app.createGrid(3, 2);
  grid.setWidget(0, 0, app.createLabel(SSNAME));
  grid.setWidget(0, 1, app.createTextBox().setName('ssname')
                          .setText(IMPORT_SSNAME_DEFAULT));
  grid.setWidget(1, 0, app.createLabel(MATR));
  grid.setWidget(1, 1, app.createTextBox().setName('matr')
                          .setText(IMPORT_FIELD_MATR));
  grid.setWidget(2, 0, app.createLabel(GRADE));
  grid.setWidget(2, 1, app.createTextBox().setName('grade')
                          .setText(IMPORT_FIELD_GRADE));
    
  var submit = app.createButton('Setzen');
  var handler = app.createServerClickHandler('submitExportParameters');
  handler.addCallbackElement(grid);
  submit.addClickHandler(handler);

  var panel = app.createVerticalPanel();
  panel.add(grid);
  panel.add(submit);
  app.add(panel);
  ass.show(app);
}

//
// Configure statistics manually.
// Adds special attributes to specific stats objects to define their
// appearence.
//
// @param stats array   an Array of stats object.
// @return Array the Array of stats objects.
//
function configureStatsManually(stats)
{
  for (var stat_index in stats)
  {
    var stat = stats[stat_index];
    if (stat['topic'].match(/^Aufgaben?$/i))
      stats[stat_index]['column_append_title'] = true;
    if (stat['topic'].match(/Note/i))
      stats[stat_index]['minmax'] = [1, 5];
    if (stat['topic'].match(/Note/i)
     && content(stat['field']) === 'Studentenanteil')
    {
      stats[stat_index]['minmax'] = [0, 1];
    }
  }
  return stats;
}

//
// Read configuration from PDS.
//
function autoloadConfig()
{
  var config = pds.load();
  if (config !== false) {
    if (config['SOURCE_BASENAME'] !== undefined)
      SOURCE_BASENAME = config['SOURCE_BASENAME'];
    if (config['SOURCE_FIRST_GROUP'] !== undefined)
      SOURCE_FIRST_GROUP = config['SOURCE_FIRST_GROUP'];
    if (config['SOURCE_LAST_GROUP'] !== undefined)
      SOURCE_LAST_GROUP = config['SOURCE_LAST_GROUP'];
    if (config['SOURCE_MINPOINTS'] !== undefined)
      SOURCE_MINPOINTS = parseInt(config['SOURCE_MINPOINTS']);
    if (config['IMPORT_SSNAME_DEFAULT'] !== undefined)
      IMPORT_SSNAME_DEFAULT = config['IMPORT_SSNAME_DEFAULT'];
    if (config['IMPORT_FIELD_MATR'] !== undefined)
      IMPORT_FIELD_MATR = config['IMPORT_FIELD_MATR'];
    if (config['IMPORT_FIELD_GRADE'] !== undefined)
      IMPORT_FIELD_GRADE = config['IMPORT_FIELD_GRADE'];
  }
}

//----------------------------------------------------------------------
// CALLBACKS
//----------------------------------------------------------------------

//
// Called by submit button of requestStatsParameters UI
//
function submitStatsParameters(e)
{
  var range = e.parameter.range.split("-");
  if (range.length !== 2)
    var range_valid = false;
  else
    var range_valid = true;
  
  // since this is not stored persistently ...
  SOURCE_BASENAME      = e.parameter.basename;
  SOURCE_MINPOINTS     = e.parameter.minpoints;
  if (range_valid)
  {
    SOURCE_FIRST_GROUP      = range[0];
    SOURCE_LAST_GROUP       = range[1];
  }
  // ... I will use pds:
  pds.store({
        SOURCE_BASENAME         : SOURCE_BASENAME,
        SOURCE_FIRST_GROUP      : SOURCE_FIRST_GROUP,
        SOURCE_LAST_GROUP       : SOURCE_LAST_GROUP,
        SOURCE_MINPOINTS        : SOURCE_MINPOINTS
  });
  
  // Clean up - get the UiApp object, close it, and return
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

//
// Called by submit button of requestExportParameters UI
//
function submitExportParameters(e)
{
  // since this is not stored persistently ...
  IMPORT_SSNAME_DEFAULT   = e.parameter.ssname;
  IMPORT_FIELD_MATR       = e.parameter.matr;
  IMPORT_FIELD_GRADE      = e.parameter.grade;

  // ... I will use pds:
  pds.store({
        IMPORT_SSNAME_DEFAULT   : IMPORT_SSNAME_DEFAULT,
        IMPORT_FIELD_MATR       : IMPORT_FIELD_MATR,
        IMPORT_FIELD_GRADE      : IMPORT_FIELD_GRADE
  });
  
  // Clean up - get the UiApp object, close it, and return
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

//
// Called when document is opened.
//
function onOpen()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Add menu item for import functionality
  var menuEntries = [
    {name: MENU_ITEM_SET, functionName: "setStatsParameters"},
    {name: MENU_ITEM_READ, functionName: "readData"},
    {name: MENU_ITEM_CREATE, functionName: "createChart"},
    {name: MENU_ITEM_SETEXPORT, functionName: "setExportParameters"},
    {name: MENU_ITEM_WRITE, functionName: "export"},
    {name: "Hilfe", functionName: "help"}
  ];
  ss.addMenu("Export", menuEntries);
}

//
// Main routine to set parameters for stats
//
function setStatsParameters()
{
  requestStatsParameters();
}

//
// Main routine to set parameters for export
//
function setExportParameters()
{
  requestExportParameters();
}

//
// Will iterate over Evaluation spreadsheets, take matrnr & total_points
// association and write data to current spreadsheet.
//
// @param write bool     shall I write data to the current sheet?
// @param students_feedback bool
//        shall I give you feedback about how many students I have read?
//
function readData(write, students_feedback)
{
  autoloadConfig();
  
  // create map to get (column_id => exercise_id) association
  var exercises = {};
  var exercise_id = 1;
  for (var col=SOURCE_FIRSTEXERCISE_COLUMN;
               col<=SOURCE_LASTEXERCISE_COLUMN; col++)
  {
    exercises[col] = exercise_id;
    data['exercises'][exercise_id++] = {};
  }
    


  for (var group_id=SOURCE_FIRST_GROUP; group_id<=SOURCE_LAST_GROUP;
       group_id++)
  {
    var ss_name = specifyGroup(SOURCE_BASENAME, group_id);
    var ss = getSpreadsheetByName(ss_name);
    if (ss === null)
    {
      Logger.log(E_404.replace(/%s/, ss_name));
      continue;
    }
    var sheet = ss.getSheetByName(SOURCE_SHEETNAME);
    var range = sheet.getRange(1, 1, sheet.getMaxRows(),
        Math.max(SOURCE_MATRNR_COLUMN, SOURCE_TOTAL_POINTS_COLUMN));

    var tutor = getRealTutorSpecification
                    (range.getCell(1, 1).getValue());
    if (!(tutor in data['tutors']))
    {
      data['tutors'][tutor] = [];
    }

    if (add_info[0] === undefined  || add_info[1] === undefined)
    {
      add_info[0] = content(range.getCell(1, 2).getValue());
      add_info[1] = content(range.getCell(1, 3).getValue());
    }

    for (var row=1; row<sheet.getMaxRows(); row++)
    {
      // Matriculation number
      var matrnr = range.getCell(row, SOURCE_MATRNR_COLUMN).getValue();
      if (typeof matrnr !== "number")
        continue;
      
      // Total points
      var points = range.getCell(row, SOURCE_TOTAL_POINTS_COLUMN)
                        .getValue();
      // skip this student, if its less than SOURCE_MINPOINTS
      if (points < SOURCE_MINPOINTS)
        continue;

      if (typeof points !== "number")
        Logger.log("Obwohl die Matrikelnummer " + matrnr + " in '"
          + ss_name + "' eine valide Nummer ist, so sind es die "
          + "zugeordneten Gesamtpunkte '" + points.toString()
          + "' es nicht");

      // Exercises
      for (var col=SOURCE_FIRSTEXERCISE_COLUMN;
               col<=SOURCE_LASTEXERCISE_COLUMN; col++)
      {
        var ex_points = parseInt(content(range.getCell(row, col)
                                              .getValue()));
        data['exercises'][exercises[col]][matrnr] = ex_points;
      }

      data['students'][matrnr] = points;
      data['tutors'][tutor].push(matrnr);
    }
  }

  if (students_feedback !== false)
  {
    var counter = getNumberOfStudents();
    if (counter === 0)
      Browser.msgBox("Keine Studenten gefunden?! Eventuell ist der "
                   + "Name der Bewertungsspreadsheets falsch gesetzt "
                   + ":-(");
    else if (counter > 0 && counter < 10)
      Browser.msgBox("Nur " + counter + " Studenten gefunden :-/");
    else
      Browser.msgBox(counter + " Studenten gefunden :-)");
  }

  if (write !== false && counter > 0)
    writeData();
}

//
// Will write data to current sheet.
//
function writeData()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.clear();
  
  var stats = {};
  
  var range = sheet.getRange(1, 1, 15 + getNumberOfTutors(), 4);  
  var first_column = sheet.getRange(1, 1, 1, 4);

  first_column.setFontSize(12);
  range.getCell(1, 1).setValue(CELL11_TITLE);
  if (add_info)
  {
    range.getCell(1, 2).setValue(add_info[0]);
    range.getCell(1, 3).setValue(add_info[1]);
  }
  
  // Tutoren
  range.getCell(3, 1).setFontWeight("bold");
  range.getCell(3, 1).setValue("Tutor");
  range.getCell(3, 2).setValue("Durchschnittsnote");
  range.getCell(3, 3).setValue("Studenten");
  range.getCell(3, 4).setValue("Maximale Punkte");
  row_id = 4;
  for (var tutor in data['tutors'])
  {
    var average = [];
    for (var student in data['tutors'][tutor])
    {
      var points = data['students'][data['tutors'][tutor][student]];
      average.push(pointsToMark(points));
      if (stats[2] === undefined || stats[2] < points)
        stats[2] = points;
    }
    stats[0] = getAverage(average).toFixed(DEC);
    stats[1] = data['tutors'][tutor].length;
    
    range.getCell(row_id, 1).setValue(tutor);
    range.getCell(row_id, 2).setValue(stats[0]);
    range.getCell(row_id, 3).setValue(stats[1]);
    range.getCell(row_id, 4).setValue(stats[2]);
    
    stats = {};
    row_id++;
  }
  
  // Aufgabe
  range.getCell(row_id, 1).setFontWeight("bold");
  range.getCell(row_id, 1).setValue("Aufgabe");
  range.getCell(row_id, 2).setValue("Studenten");
  range.getCell(row_id, 3).setValue("Durchschnittspunkte");
  row_id++;
  for (var exercise in data['exercises'])
  {
    var participated_students = 0;
    var points = [];
    for (var student in data['exercises'][exercise])
    {
      var add = false;
      if (data['exercises'][exercise][student] > 0)
        participated_students++;
      points.push(data['exercises'][exercise][student]);
    }
    
    range.getCell(row_id, 1).setValue(exercise);
    range.getCell(row_id, 2).setValue(participated_students);
    range.getCell(row_id, 3).setValue(getAverage(points).toFixed(DEC));
    row_id++;
  }
  
  // Noten
  range.getCell(row_id, 1).setFontWeight("bold");
  range.getCell(row_id, 1).setValue("Note");
  range.getCell(row_id, 2).setValue("Studenten");
  range.getCell(row_id, 3).setValue("Studentenanteil");
  row_id++;
  var marks = {1:0, 2:0, 3:0, 4:0, 5:0};
  var nr_students = 0;
  for (var student in data['students'])
  {
    marks[pointsToMark(data['students'][student])]++;
    nr_students++;
  }

  for (var mark in marks)
  {
    range.getCell(row_id, 1).setValue(mark);
    range.getCell(row_id, 2).setValue(marks[mark]);
    range.getCell(row_id, 3).setValue((marks[mark] / nr_students)
                                         .toFixed(DEC));
    row_id++;
  }
}

//
// Create Charts for stats
//
function createChart()
{
  var ass = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ass.getActiveSheet();
  var data_range = sheet.getDataRange();

  var row_id = STATS_FIRST_ROW;

  // search for first row with bold text
  while (row_id <= data_range.getLastRow() &&
         data_range.getCell(row_id, 1).getFontWeight() != "bold")
    row_id++;

  if (row_id == data_range.getLastRow())
  {
    Browser.msgBox("Keine Statistiken gefunden. Der Titel der "
                 + "Statistiken sollte fett gedruckt in Spalte 1 sein. "
                 + "Abbruch.");
    return false;
  }

  // get all ranges ((start, end) pairs between topics)
  var ranges = [[row_id]];
  while (row_id <= data_range.getLastRow()
         && content(data_range.getCell(row_id, 1).getValue()) != "")
  {
    if (data_range.getCell(row_id, 1).getFontWeight() === "bold")
    {
      ranges[ranges.length - 1].push(row_id);
      ranges.push([row_id]);
    }
    row_id++;
  }
  ranges[ranges.length - 1].push(row_id);

  var stats = []; // each statistic will be an object in this Array

  // actually read the stats
  var col_id = STATS_FIRST_COL + 1;
  for (var range in ranges)
  {
    var begin = ranges[range][0];
    var end   = ranges[range][1];

    if (begin === end) // if there is only 1 row with bold text, skip it
      continue;

    while (col_id <= data_range.getLastColumn() &&
           content(data_range.getCell(begin, col_id).getValue()) != "")
    {
      var statistic = readStat(STATS_FIRST_COL, col_id, begin, end-1);
      if (statistic !== false)
        stats.push(statistic);
      else
        Logger.log("Problem with stats at (" +begin+ ", " +col_id+ ")");

      col_id++;
    }
    
    // reset to second column for next stats
    col_id = STATS_FIRST_COL + 1;
  }
  if (stats.length === 0)
  {
    Browser.msgBox("Keine Statistiken gefunden. Abbruch.");
    return false;
  } else {
    Browser.msgBox(stats.length + " Statistiken gelesen. Werde Charts "
                                + "erzeugen.");
  }

  // add configuration to stats  
  stats = configureStatsManually(stats);
  
  var options = ['column_append_title', 'minmax'];

  // Create stats charts
  var create_counter = 0;
  for (var stat_index in stats)
  {
    var stat = stats[stat_index];
    var title = stat['field'] + " pro " + stat['topic'];

    // evaluate type & number of rows [in a very inefficient way :-/ ]
    var counter = 0;
    for (var key in stat)
    {
      if (contains(options, key))
        continue;          // skip options
      if (contains(["field", "topic"], key))
        continue;          // skip metadata
      
      if (++counter > 1)
        continue;
    
      if (typeof key === "number")
        var type1 = Charts.ColumnType.NUMBER;
      else
        var type1 = Charts.ColumnType.STRING;

      if (typeof stat[key] === "number")
        var type2 = Charts.ColumnType.NUMBER;
      else
        var type2 = Charts.ColumnType.STRING;
    }

    var dataTable = Charts.newDataTable()
                          .addColumn(type1, stat['topic'])
                          .addColumn(type2, stat['field']);

    var counter = 0;
    for (var key in stat)
    {
      if (contains(options, key))
        continue;          // skip options
      if (contains(["field", "topic"], key))
        continue;          // skip metadata

      if (stat['column_append_title'] === true)
        dataTable.addRow([trim(stat['topic']) + " " + key, stat[key]]);
      else
        dataTable.addRow([key, stat[key]]);
      counter++;
    }
    dataTable.build();

    var chart = Charts.newColumnChart().setRange(1, 5)  
                      .setDataTable(dataTable)
                      .setColors([CHART_COLOR])
                      .setDimensions(CHART_WIDTH, CHART_HEIGHT)
                      .setXAxisTitle(stat['topic'])
                      .setYAxisTitle(stat['field'])
                      .setTitle(title);
    if (typeof(stat['minmax']) === "object"
        && stat['minmax'].length == 2)
    {
      chart.setRange(stat['minmax'][0], stat['minmax'][1]);
    }
    chart = chart.build();

    // Save the chart to our Document List
    var file = DocsList.createFile(chart);  
    file.rename("[Chart] " + title);
    create_counter++;
  }
  
  Browser.msgBox(create_counter + " Charts erzeugt :-)");

  return ui;
}

//
// Write data (matrnr => total_points) to TUG Online Spreadsheet.
//
// @return bool success (true) or failure (false)
//
function writeExportData()
{
  var found = true;
  var ass = getSpreadsheetByName(IMPORT_SSNAME_DEFAULT);
  if (ass === null)
    found = false;
  if (found)
  {
    var sheet = ass.getSheetByName(IMPORT_SHEETNAME_DEFAULT);
    if (sheet === null)
      found = false; 
  }
  if (!found)
  {
    Browser.msgBox(E_EXPORT_NOTFOUND
                       .replace(/%s/, IMPORT_SSNAME_DEFAULT)
                       .replace(/%s/, IMPORT_SHEETNAME_DEFAULT));
    return false;
  }
  var range = sheet.getDataRange();
  
  // read data['students']
  readData(false, false);

  if (!data['students'])
  {
    Browser.msgBox(E_READDATA);
    return false;
  }
  
  // Okay, now we have data['students'] with the data to export
  // and sheet/range point to the sheet to export data to.
  
  var matrnr_col = 0;
  var grade_col = 0;
  
  // search in first 5 rows for matrnr/tpoints columns
  search:
  for (var row=1; row<=5; row++)
  {
    for (var col=1; col<=range.getLastColumn(); col++)
    {
      var val = content(range.getCell(row, col).getValue());
      var regex1 = new RegExp(IMPORT_FIELD_MATR, "i");
      var regex2 = new RegExp(IMPORT_FIELD_GRADE, "i");
      if (val.match(regex1))
        matrnr_col = col;
      if (val.match(regex2))
        grade_col = col;
      
      if (matrnr_col !== 0 && grade_col !== 0)
        break search;
    }
  }
  if (matrnr_col === 0)
  {
    Browser.msgBox("Konnte Spalte mit Matrikelnummern nicht finden. "
                 + "Regex matcht nicht.");
    return false;
  }
  if (grade_col === 0)
  {
    Browser.msgBox("Konnte Spalte mit Gesamtpunkten nicht finden. "
                 + "Regex matcht nicht.");
    return false;
  }
  
  var wrote_something = false;
  for (var row=1; row<range.getLastRow(); row++)
  {
    var val_matrnr = content(range.getCell(row, matrnr_col).getValue());
    
    var tpts = data['students'][val_matrnr];
    var grade = pointsToMark(tpts);
    if (tpts !== undefined)
    {
      range.getCell(row, grade_col).setValue(grade);
      wrote_something = true;
      if (grade === 0)
        Logger.log("Student " + val_matrnr + " got grade 0!");
    } else
      if (typeof(val_matrnr) === "number")
        Logger.log("Cannot find student " + val_matrnr + " in "
                 + "data[students]");
  }
  if (!wrote_something)
  {
    Browser.msgBox("Ich konnte keiner einzigen Matrikelnummer die "
                 + "Gesamtpunkte zuordnen. Mehr ist nicht bekannt :-(");
  }

  return true;
}

//
// Export marks to TUG Online CSV.
//
function export()
{
  if (!writeExportData())
    Browser.msgBox("Konnte Daten nicht exportieren. Abbruch.");
  else
    Browser.msgBox("Daten exportiert.");
}

//
// Help context menu in MsgBox
//
function help()
{
  var ass = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle('Export Skript');
  
  var p1 = "Dieses Skript ermöglicht es die Ergebnisse der Benotung "
         + "aus den einzelnen Gruppen zusammenzuführen. Im zweiten "
         + "Schritt können entsprechende Statistiken generiert werden. "
         + "Ein optionaler dritter Schritt erlaubt es eine TUG Online "
         + "CSV in einem Spreadsheet mit den Noten zu füllen.";
  var p2 = "Wähle diese Menüpunkte, wenn du diese Funktionen nutzen "
         + "willst.";
  var panel = app.createVerticalPanel();
  panel.add(app.createLabel(p1));
  panel.add(app.createLabel(p2));
  app.add(panel);
  ass.show(app);
}
