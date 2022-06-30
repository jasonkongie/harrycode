
/**Class to more easily access sheets in the active spreadsheet.*/
class SheetHandler //allows for easy access of sheets
{

  //columnDict is a dict with key of attribute name (any string) and value of column index (column A is value 1, B is 2, etc.)
  //after passing in columnDict, it is used to more easily get the column index of an attribute
  /**
   * Creates a SheetHandler
   * @param {String} sheetName The name of the Sheet in the active Spreadsheet.
   * @param {Object} columnDict Object with string attributes to their corresponding column indexes.
   * Example: {'time':0, 'name':1, 'homeroomName':2, 'subject':3, 'email':4}
   */
  constructor(sheetName, columnDict) //example: this.teachersSheet = new SheetHandler('Teachers', {'time':0, 'name':1, 'homeroomName':2, 'subject':3, 'email':4});
  {
    this.sheetName = sheetName;
    this.sheet = Spreadsheets.activeSpreadsheet.getSheetByName(sheetName);
    if (this.sheet == undefined) throw 'Spreadsheet named "' + sheetName + '" not found'; //sheet cannot be undefined, means there was an error retrieving the sheet
    this.fullDataRange = this.sheet.getDataRange();
    this.sheetValues = this.fullDataRange.getValues();
    this.backgroundValues = this.fullDataRange.getBackgrounds();
    if (columnDict == undefined) //cannot be undefined because we need info about the columns
    {
      throw 'columnDict for SheetHandler "' + this.sheetName + '" is uninitialized, please set the columnDict attribute of SheetHandler with a dictionary containing attribute names (keys) and their corresponding column indices (values) starting with column A as value 1, B as 2, C as 3, etc.';
    }
    this.columnDict = columnDict; //key: attribute name; value: column index
    
  }

  /**Returns the column index of the requested attribute using the SheetHandler's columnDict
   * @param {String} attributeName The name of the attribute (specified in columnDict).
   * @returns {number} The column index of the attribute.
  */
  getAttributeColumnIndex(attributeName) //returns value using columnDict
  {
    if (this.columnDict[attributeName] == undefined) //error: we have no info about the column index of the specified attribute
    {
      throw 'columnDict for SheetHandler for sheet named "' + this.sheetName + '" had no key "' + attributeName + '". Is parameter attributeName typed correctly? Is columnDict in SheetHandler contructor(sheetName, columnDict) initialized correctly? (Check the SpreadsheetHolder class)';
      return;
    }
    return this.columnDict[attributeName];
  }

  /**Returns the cell value of the attribute at the specified row in the sheet
   * @param {String} attributeName The name of the attribute (specified in columnDict).
   * @param {number} rowIndex The index of the row to get.
   * @returns {String} The value of the cell.
  */
  getAttribute(attributeName, rowIndex) //row index 0 in the sheet are the labels/questions, rows after that is data
  {
    return this.sheetValues[rowIndex][this.getAttributeColumnIndex(attributeName)];
  }

  /**Returns the cell background value of the attribute at the specified row in the sheet
   * @param {String} attributeName The name of the attribute (specified in columnDict).
   * @param {number} rowIndex The index of the row to get.
   * @return {String} The value of the background cell in a hexadecimal string.
  */
  getAttributeBackground(attributeName, rowIndex)
  {
    return this.backgroundValues[rowIndex][this.getAttributeColumnIndex(attributeName)];
  }

  /**Sets the cell value of the attribute at the specified row in the sheet. Remember to upload the changes.
   * @param {String} attributeName The name of the attribute (specified in columnDict).
   * @param {number} rowIndex The index of the row to set.
   * @param {String} newValue The new value.
  */
  setAttribute(attributeName, rowIndex, newValue)
  {
    this.sheetValues[rowIndex][this.getAttributeColumnIndex(attributeName)] = newValue;
  }

  /**Sets the cell background value of the attribute at the specified row in the sheet. Remember to upload the changes.
   * @param {String} attributeName The name of the attribute (specified in columnDict).
   * @param {number} rowIndex The index of the row to set.
   * @param {String} newValue The new value.
  */
  setAttributeBackground(attributeName, rowIndex, newValue)
  {
    this.backgroundValues[rowIndex][this.getAttributeColumnIndex(attributeName)] = newValue;
  }

  /**Sets the cell background values of all cells at the specified row in the sheet. Remember to upload the changes.
   * @param {number} rowIndex The index of the row to set.
   * @param {String} newValue The new value for each cell.
  */
  setRowBackground(rowIndex, newValue)
  {
    for (var i = 0; i < this.backgroundValues[rowIndex].length; i++)
    {
      this.backgroundValues[rowIndex][i] = newValue;
    }
  }
  
  /**Batch sets the values and background values of the sheet handled by SheetHandler.
  */
  uploadChanges() //batch edits the sheet to have the values of this.sheetValues
  {
    Logger.log("Uploading changes to sheet " + this.sheetName);
    //https://developers.google.com/apps-script/reference/spreadsheet/range#setvaluesvalues
    this.fullDataRange.setBackgrounds(this.backgroundValues);
    return this.fullDataRange.setValues(this.sheetValues);
  }

}

var Spreadsheets =
{
  //valid keys: time, email, name, grade, homeroomName, purpose, status, destination, location
  activeSpreadsheet: SpreadsheetApp.getActiveSpreadsheet(),
  errorSheetUrl: 'https://docs.google.com/spreadsheets/d/1PRZJjLLi8KE73fka0cpP5hxB-452xpZ7VMjx8sdF9gw/edit',
  _studentsSheetHandler: null, //keys: time, name, grade, homeroomName, email
  _requestsSheetHandler: null, //keys: time, email, name, destination, purpose, status
  _teachersSheetHandler: null, //keys: time, email, name, homeroomName, subject

  getStudentsSheetHandler: function()  {
    if (!this._studentsSheetHandler)
    {
      this._studentsSheetHandler = new SheetHandler(Spreadsheets.studentsSheetName, {'time':0, 'email':1, 'name':2, 'grade':3, 'homeroomName':4});
    }
    return this._studentsSheetHandler;
  },

  getTeachersSheetHandler: function()  {
    if (!this._teachersSheetHandler)
    {
      this._teachersSheetHandler = new SheetHandler(Spreadsheets.teachersSheetName, {'time':0, 'email':1, 'name':2, 'homeroomName':3, 'subject':4, 'location':5});
    }
    return this._teachersSheetHandler;
  },

  getRequestsSheetHandler: function()  {
    if (!this._requestsSheetHandler)
    {
      this._requestsSheetHandler = new SheetHandler(Spreadsheets.requestsSheetName, {'time':0, 'email':1, 'name':2, 'destination':3, 'purpose':4, 'status':5});
    }
    return this._requestsSheetHandler;
  },

  studentsSheetName: "Students",//"Students",
  teachersSheetName: "Teachers",//"Teachers",
  requestsSheetName: "Form Requests",
  approvalSheetsColDict: {'email':0, 'name':1, 'purpose':2, 'status':3}
}
//real:
/*
Spreadsheets.studentsSheet = new SheetHandler(Spreadsheets.studentsSheetName, {'time':0, 'email':1, 'name':2, 'grade':3, 'homeroomName':4}); //has student info, read-only
Spreadsheets.requestsSheet = new SheetHandler(Spreadsheets.requestsSheetName, {'time':0, 'email':1, 'name':2, 'destination':3, 'purpose':4, 'status':5}); //has student requests, writeable (we will edit this form),
Spreadsheets.teachersSheet = new SheetHandler(Spreadsheets.teachersSheetName, {'time':0, 'email':1, 'name':2, 'homeroomName':3, 'subject':4, 'location':5}); //has teacher & room information, read-only
*/

class SpreadsheetBuilder
{
  //all rows in SpreadsheetBuilder will grow to the same length (separate for cell and background values)
  constructor(cellRowCapacity, bgRowCapacity) //initial capacities
  {
    this._cellRowCapacity = cellRowCapacity;
    this._bgRowCapacity = bgRowCapacity;
    this.cellRowsData = [];
    this.bgRowsData = [];
    this.cellRowFiller = '';
    this.bgRowFiller = null;
  }

  _forceAllRowsLength(newLength, rowList, rowFiller) //force all rows in list to same length
  {
    for (var rowIdx = 0; rowIdx < rowList.length; i++)
    {
      row = rowList[rowIdx];
      const toAddCount = newLength - row.length;
      if (toAddCount > 0)
      {
        for (var i = 0; i < toAddCount; i++)
        {
          row.push(rowFiller);
        }
      }
      else if (toAddCount < 0)
      {
        rowList.splice(toAddCount, -1 * toAddCount); //remove last elements from row
      }
    }
  }

  forceAllCellRowsLength(newLength)
  {
    this._forceAllRowsLength(newLength, this.cellRowsData, this.cellRowFiller);
  }

  forceAllBgRowsLength(newLength)
  {
    this._forceAllRowsLength(newLength, this.bgRowsData, this.bgRowFiller);
  }

  _updateRowsData(rowList, rowCapacity, newRow, rowFiller) //force new row to same length as others or, if row is largest row, set other rows to this length
  {
    if (newRow === undefined) {newRow = []}
    if (newRow.length > rowCapacity)
    {
      rowCapacity = newRow.length;
      //make all rows the new highest length
      this._forceAllRowsLength(newRow.length, rowList, rowFiller);
      rowList.push(newRow);
    }
    else
    {
      //make new row the same length as the row capacity
      const toAddCount = rowCapacity - newRow.length;
      for (var i = 0; i < toAddCount; i++)
      {
        newRow.push(rowFiller);
      }
      rowList.push(newRow);
    }
    return rowCapacity; //new row capacity
  }

  addRow(newRow)
  {
    this._cellRowCapacity = this._updateRowsData(this.cellRowsData, this._cellRowCapacity, newRow, this.cellRowFiller);
  }

  addBackgroundRow(newRow)
  {
    this._bgRowCapacity = this._updateRowsData(this.bgRowsData, this._bgRowCapacity, newRow, this.bgRowFiller);
  }

  _fillRowsUntilLength(rowList, rowCapacity, rowFiller, newLength) //adds empty rows until the row list reaches a specified length
  {
    const toAdd = newLength - rowList.length;
    if (toAdd <= 0) return;

    //create empty row
    var emptyRow = [];
    for (var i = 0; i < rowCapacity; i++)
    {
      emptyRow.push(rowFiller);
    }
    //add empty rows to row list until it reaches specified length
    for (var i = 0; i < toAdd; i++)
    {
      rowList.push(emptyRow);
    }
  }

  fillCellRowsUntilLength(newLength)
  {
    this._fillRowsUntilLength(this.cellRowsData, this._cellRowCapacity, '', newLength);
  }

  fillBgRowsUntilLength(newLength)
  {
    this._fillRowsUntilLength(this.bgRowsData, this._bgRowCapacity, null, newLength);
  }

  get cellRowLength()
  {
    return this._cellRowCapacity;
  }

  get bgRowLength()
  {
    return this._bgRowCapacity;
  }
}

class ApprovalSheetReader
{
  /** 
   * @param {SpreadsheetApp.Sheet} sheet*/
  constructor(sheet, studentsDict)
  {
    this.sheet = sheet;
    this.sheetValues = sheet.getDataRange().getValues();

    this.studentsApprovedDict = {};
    this.studentsDeniedDict = {};

    for (var i = 1; i < this.sheetValues.length; i++)
    {
      const email = this.sheetValues[i][Spreadsheets.approvalSheetsColDict.email];
      //const name = this.sheetValues[i][Spreadsheets.approvalSheetsColDict.name];
      const status = this.sheetValues[i][Spreadsheets.approvalSheetsColDict.status];
      const student = studentsDict[email];
      if (status == "Approve")
      {
        this.studentsApprovedDict[student.email] = student;
      }
      else
      {
        this.studentsDeniedDict[student.email] = student;
      }
    }
  }
}

class Room
{
  constructor (name, location)
  {
    this.name = name;
    //dicts indexed by email
    this.studentsDict = Object.create(null);
    this.teachersDict = Object.create(null);
    this.studentsDepartedDict = Object.create(null);
    this.studentsArrivedDict = Object.create(null);
    //default occupancy
    this.location = location;
  }

  get studentCount()
  {
    return Object.keys(this.studentsDict).length;
  }

  get teacherCount()
  {
    return Object.keys(this.teachersDict).length;
  }

  get occupantCount()
  {
    return this.studentCount + this.teacherCount;
  }
}

class Student
{
  constructor (email, name, grade, homeroomName, purpose)
  {
    this.email = email;
    this.name = name;
    this.grade = grade;
    this.homeroomName = homeroomName; //string
    this.purpose = purpose;
    this.currentRoom; //Room object
    
    this.hasRequest = false;
    this.requestRowIndex;
  }
  
  moveToRoom(destinationRoom)
  {
    if (destinationRoom == this.currentRoom) return; //cannot move to same room
    if (this.currentRoom)
    {
      delete this.currentRoom.studentsDict[this.email]; //remove from current room if in room
      if (this.currentRoom.name == this.homeroomName && destinationRoom) this.currentRoom.studentsDepartedDict[this.email] = this; //record student departure for attendance sheet unless we are just deleting the student (when destinationRoom is undefined)
    }
    if (destinationRoom)
    {
      destinationRoom.studentsDict[this.email] = this; //add to new room
      if (destinationRoom.name != this.homeroomName) destinationRoom.studentsArrivedDict[this.email] = this; //only record student arrival for attendance sheet if this isn't their homeroom
    }
    Logger.log(this.name + ", " + (this.currentRoom? this.currentRoom.name : "No room") + " -> " + (destinationRoom? destinationRoom.name : "No room"));
    this.currentRoom = destinationRoom;
  }

  sendEmail(subject, body)
  {
    MailApp.sendEmail(this.email, subject, body);
  }

  /**sets the status row of the student's request on the Sheet Handler
   * Remember to use SheetHandler.uploadChanges()
   * @param {String} newStatus
  */
  setRequestStatus(newStatus)
  {
    if (!this.requestRowIndex)
    {
      throw 'attribute requestRowIndex for student ' + student.name + ' not found. Did you forget to set requestRowIndex?';
    }
    Spreadsheets.getRequestsSheetHandler().setAttribute('status', this.requestRowIndex, newStatus);
  }

  /**sets the background color of status row of the student's request on the Sheet Handler
   * Remember to use SheetHandler.uploadChanges()
   * @param {String} newColor
  */
  setRequestBackground(newColor)
  {
    if (!this.requestRowIndex)
    {
      throw 'attribute requestRowIndex for student ' + student.name + ' not found. Did you forget to set requestRowIndex?';
    }
    Spreadsheets.getRequestsSheetHandler().setRowBackground(this.requestRowIndex, newColor);
  }
}

class Teacher
{
  constructor (email, name, homeroomName, subject)
  {
    this.email = email;
    this.name = name;
    this.homeroomName = homeroomName; //string
    this.subject = subject;
    this.currentRoom = undefined; //Room object
  }

  moveToRoom(room)
  {
    if (room == this.currentRoom) return;
    if (this.currentRoom) delete this.currentRoom.teachersDict[this.email]; //remove from current room
    if (room) room.teachersDict[this.email] = this; //add to new room
    Logger.log(this.name + ", " + (this.currentRoom? this.currentRoom.name : "No room") + " -> " + (room? room.name : "No room"));
    this.currentRoom = room;
  }

  sendEmail(subject, body)
  {
    MailApp.sendEmail(this.email, subject, body);
  }
}

function createRoomsWithTeachers(roomsDict, teachersDict)
{
  for (var i = 1; i < Spreadsheets.getTeachersSheetHandler().sheetValues.length; i++) //populate rooms and teachers dicts by looping through sheet top to bottom
  {
    //use this row's values to create a teacher and assign it to the dict
    const name = Spreadsheets.getTeachersSheetHandler().getAttribute('name', i);
    const homeroomName = Spreadsheets.getTeachersSheetHandler().getAttribute('homeroomName', i);
    const subject = Spreadsheets.getTeachersSheetHandler().getAttribute('subject', i);
    const email = Spreadsheets.getTeachersSheetHandler().getAttribute('email', i);
    const location = Spreadsheets.getTeachersSheetHandler().getAttribute('location', i);
    const teacher = new Teacher(email, name, homeroomName, subject);
    teachersDict[email] = teacher;

    //add teacher to their homeroom
    var room = roomsDict[homeroomName];
    if (room == undefined) //room doesn't exist so create the room with homeroomName 
    {
      room = new Room(homeroomName, location);
      roomsDict[homeroomName] = room;
    }
    teacher.moveToRoom(room);
  }
}

function createStudentsInRooms(studentsDict, roomsDict) //only uses most recently submitted student info entry if duplicate emails exist
{
  for (var i = 1; i < Spreadsheets.getStudentsSheetHandler().sheetValues.length; i++) //populate dict by looping through sheet top to bottom
  {
    //use this row's values to create a student and assign it to the dict
    const email = Spreadsheets.getStudentsSheetHandler().getAttribute('email', i);
    const name = Spreadsheets.getStudentsSheetHandler().getAttribute('name', i);
    const grade = Spreadsheets.getStudentsSheetHandler().getAttribute('grade', i);
    const homeroomName = Spreadsheets.getStudentsSheetHandler().getAttribute('homeroomName', i);
    const student = new Student(email, name, grade, homeroomName);
    if (studentsDict[email]) //student with same email was already created; we need to get rid of that one and use this more recent one
    {
      studentsDict[email].moveToRoom(); //moving old student out of their current room , then reassignning the studentsDict email key to the new student will effectively get rid of the old one
    }
    studentsDict[email] = student;

    //add student to their homeroom
    var room = roomsDict[homeroomName]; //find homeroom
    if (room == undefined) //room doesn't exist (student homeroom not found)
    {
      Logger.log('WARNING: Student [' + name + "]'s homeroom " + homeroomName + " on row " + (i + 1) + " was not found - this shouldn't happen. This probably means that the students data collection form contains an invalid room or a room with no teachers.");
    }
    else
    {
      student.moveToRoom(room);
    }
  }
}

function markInvalidRequests(studentsDict)
{
  //we want to check for and remove older requests from the students
  //we also want to remove requests if the email is not in our student database
  //loop through requests data bottom to top (as the more recent requests should be the actual request, the older ones from that student are outdated)
  for (var i = Spreadsheets.getRequestsSheetHandler().sheetValues.length - 1; i > 0; i--)
  {
    const status = Spreadsheets.getRequestsSheetHandler().getAttribute('status', i);
    if (status != "Request")
    {
      continue;
    }
    const email = Spreadsheets.getRequestsSheetHandler().getAttribute('email', i);
    const student = studentsDict[email]; //get student who submitted the request
    if (student == undefined)
    {
      Logger.log("Student with email [" + email + "] at row " + (i + 1) + " in requests not found in student data sheet. Denying student request...");
      MailApp.sendEmail(email, 'Silver Slip Request Denied', 'Hello, your silver slip request was denied.\nReason: Student not in database.\nPlease fill out this form:\nhttps://docs.google.com/forms/d/e/1FAIpQLSeXbY519X--sQkTyqJmBshfLS35XIqsXxA-sUPmqgqo3Js5qg/viewform \n\nThen, fill out the silver slip request form again.');
      Spreadsheets.getRequestsSheetHandler().setAttribute('status', i, "Denied,NotInDatabase");
      Spreadsheets.getRequestsSheetHandler().setRowBackground(i, 'DeepPink');
    }
    else
    {
      if (student.hasRequest == true) //we've already encountered the student's request so this one is old, void
      {
        Spreadsheets.getRequestsSheetHandler().setAttribute('status', i, 'Overridden');
        Spreadsheets.getRequestsSheetHandler().setRowBackground(i, 'DimGrey');
        continue;
      }
      else
      {
        student.hasRequest = true;
        student.requestRowIndex = i;
      }
    }
  }
  Spreadsheets.getRequestsSheetHandler().uploadChanges();
  SpreadsheetApp.flush();
}

function createNewSpreadsheetFile(fileName, directory)
{
  const file = DriveApp.getFileById(SpreadsheetApp.create(fileName).getId());
  file.moveTo(directory);
  return file;
}

function getCurrentDirectory() //get parent folder of this (current directory that we are in)
{
  const directory = DriveApp.getFileById(Spreadsheets.activeSpreadsheet.getId()).getParents().next();
  if (directory == 'My Drive') throw 'Please place the Silver Slip Google sheets in a folder';
  return directory;
}

function getOrCreateFolderInDirectory(folderName, directory) //returns the folder if it exists; if the folder does not exist, a folder is created in the directory and returned
{
  Logger.log('Getting or creating folder ' + folderName + " in " + directory.getName());

  const folders = directory.getFoldersByName(folderName);
  var targetFolder;
  if (folders.hasNext())
  {
    targetFolder = folders.next();
  }
  else
  {
    targetFolder = directory.createFolder(folderName);
  }
  return targetFolder;
}

function getOrCreateSpreadsheetFileInDirectory(fileName, directory) //returns the first file with the name if it exists; if the file does not exist, a file is created in the directory and returned
{
  Logger.log('Getting or creating file ' + fileName + " in " + directory.getName());

  const files = directory.getFilesByName(fileName);
  var targetFile;
  if (files.hasNext())
  {
    targetFile = files.next();
  }
  else
  {
    targetFile = createNewSpreadsheetFile(fileName, directory);
  }
  return targetFile;
}

function getDictOfSpreadsheetFilesInFolder(folder) //key: file name, value: sheet file
{
  const filesDict = Object.create(null); //key: file name, value: File Object
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()){
    const file = files.next();
    filesDict[file.getName()] = file;
  }
  return filesDict;
}


function createAttendanceSheets(attendanceFolder, roomsDict) //creates an attendance spreadsheet for each room and shares it with the teacher of that room
{
  //create an attendance sheet for each room
  Logger.log('Creating attendance sheets for each room...')

  //create dict of all the files in the attendance folder
  const filesDict = getDictOfSpreadsheetFilesInFolder(attendanceFolder);

  //make attendance sheets by either overwriting current existing file (file with the same name) or create the file if it does not exist
  for (const room of Object.values(roomsDict))
  {
    Logger.log("Writing: " + room.name);
    //get the attendance sheet file
    var sheetFile = filesDict[room.name];
    if (!sheetFile)
    {
      sheetFile = createNewSpreadsheetFile(room.name, attendanceFolder);
    }

    const ssBuilder = new SpreadsheetBuilder(3, 0);
    ssBuilder.addRow(['','','']);
    ssBuilder.addRow([room.name + ' Embedded Time Attendance:', '', '']);
    ssBuilder.addRow(['', '', '']);
    
    //write students that joined
    ssBuilder.addRow(['', '', '']);
    ssBuilder.addRow(['Arriving Students:', '', '']);
    ssBuilder.addRow([Object.values(room.studentsArrivedDict).length + ' arrived.', '', '']);
    ssBuilder.addRow(['Grade', 'Name', 'From']);
    for (const student of Object.values(room.studentsArrivedDict))
    {
      ssBuilder.addRow([student.grade, student.name, student.homeroomName]);
    }
    ssBuilder.addRow(['', '', '']);
    
    //write students that left
    ssBuilder.addRow(['Departing Students:', '', '']);
    ssBuilder.addRow([Object.values(room.studentsDepartedDict).length + ' departed.', '', '']);
    ssBuilder.addRow(['Grade', 'Name', 'To']);
    for (const student of Object.values(room.studentsDepartedDict))
    {
      ssBuilder.addRow([student.grade, student.name, student.currentRoom.name]);
    }
    ssBuilder.addRow(['', '', '']);
    
    //write full student attendance to sheet
    /*
    ssBuilder.addRow(['Full Student Attendance:', '', '']);
    ssBuilder.addRow([room.studentCount + ' students.', '', '']);
    ssBuilder.addRow(['Grade', 'Name', '']);
    for (const student of Object.values(room.studentsDict))
    {
      ssBuilder.addRow([student.grade, student.name, '']);
    }
    ssBuilder.addRow(['', '', '']);
    */
    
    //write teachers to sheet, make them editors
    ssBuilder.addRow(['Teachers:', '', '']);
    ssBuilder.addRow([room.teacherCount + ' teachers.', '', '']);
    for (const teacher of Object.values(room.teachersDict))
    {
      ssBuilder.addRow([teacher.name, '', '']);
      sheetFile.addEditor(teacher.email);
      //teacher.sendEmail('Embedded Time Attendance Sheet', "The link to your Embedded Time classroom's attendance sheet is below:\n" + sheetFile.getUrl()); //prob unnecessary because adding editor automatically sends email
    }

    //write over the remaining data
    const sheet = SpreadsheetApp.open(sheetFile).getActiveSheet();
    const sheetNumRows = sheet.getDataRange().getNumRows();
    ssBuilder.fillCellRowsUntilLength(sheetNumRows);
    sheet.getRange(1, 1, ssBuilder.cellRowsData.length, ssBuilder.cellRowLength).setValues(ssBuilder.cellRowsData);
    SpreadsheetApp.flush();
  }
  Logger.log('Attendance sheets at folder ' + attendanceFolder.getName());
}


function resetRequestsSheet()
{
  const requestsSheet = Spreadsheets.activeSpreadsheet.getSheetByName(Spreadsheets.requestsSheetName);
  const fullDataRange = requestsSheet.getDataRange();
  var sheetValues = fullDataRange.getValues();
  var bgValues = fullDataRange.getBackgrounds();

  //make every value (except for first row) in every inner array of sheetValues an empty string to make sheet blank
  for (var i = 1; i < sheetValues.length; i++)
  {
    var row = sheetValues[i];
    for (var j = 0; j < row.length; j++)
    {
      row[j] = '';
    }
  }
  fullDataRange.setValues(sheetValues);
  //do the same thing for backgrounds
  for (var i = 1; i < bgValues.length; i++)
  {
    var row = bgValues[i];
    for (var j = 0; j < row.length; j++)
    {
      row[j] = null;
    }
  }
  fullDataRange.setBackgrounds(bgValues);

  requestsSheet.clearNotes();

  //remove all range protections: https://developers.google.com/apps-script/reference/spreadsheet/sheet#getprotectionstype
  var protections = requestsSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
  var protection = protections[i];
  if (protection.canEdit()) {
    protection.remove();
  }
  }
}

function saveAndClearRequestsSheet(saveFolder)
{
  const sheet = SpreadsheetApp.open(createNewSpreadsheetFile(new Date().toString(), saveFolder)).getActiveSheet();
  const sheetValues = Spreadsheets.getRequestsSheetHandler().sheetValues;
  const backgroundValues = Spreadsheets.getRequestsSheetHandler().backgroundValues;
  sheet.getRange(1, 1, sheetValues.length, sheetValues[0].length).setValues(sheetValues);
  sheet.getRange(1, 1, backgroundValues.length, backgroundValues[0].length).setBackgrounds(backgroundValues);
  sheet.autoResizeColumns(1, sheetValues[0].length);
  resetRequestsSheet();
}

function updateStatisticsSheet(statisticsFolderName, statisticsFileName, roomsDict)
{
  Logger.log("Writing to statistics sheet...");

  const currentDirectory = getCurrentDirectory();
  const statisticsFolder = getOrCreateFolderInDirectory(statisticsFolderName, currentDirectory);

  const timestamp = new Date().toDateString();
  const sheetFile = getOrCreateSpreadsheetFileInDirectory(statisticsFileName, statisticsFolder);
  const statisticsSheet = SpreadsheetApp.open(sheetFile).getActiveSheet();

  //create nested array to pass to Range.setValues
  var sheetValues = [];
  //all inner arrays must have the same number of elements
  if (statisticsSheet.getLastRow() == 0) sheetValues.push(['Timestamp', 'Room Name', 'default', 'joined', 'left']); //only put the title if there is no data
  for (const room of Object.values(roomsDict))
  {
    sheetValues.push([timestamp, room.name, room.defaultOccupancy, Object.keys(room.studentsArrivedDict).length, Object.keys(room.studentsDepartedDict).length]);
  }
  statisticsSheet.getRange(statisticsSheet.getLastRow() + 1, 1, sheetValues.length, sheetValues[0].length).setValues(sheetValues); //appends the data to the end of the sheet

  Logger.log("Statistics Sheet in folder " + statisticsFolder.getName());
}
