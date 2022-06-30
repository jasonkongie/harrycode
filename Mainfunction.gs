function updateApprovalSheets()
{
  Logger.log("Updating approval sheets...");
  const studentsDict = Object.create(null); //indexed by email
  const teachersDict = Object.create(null); //indexed by email
  const roomsDict = Object.create(null); //indexed by room name
  createRoomsWithTeachers(roomsDict, teachersDict);
  createStudentsInRooms(studentsDict, roomsDict);

  markInvalidRequests(studentsDict);

  //move students to their requested room by looking through requests
  for (var i = 1; i < Spreadsheets.getRequestsSheetHandler().sheetValues.length; i++)
  {
    const status = Spreadsheets.getRequestsSheetHandler().getAttribute('status', i);
    if (status != "Request") continue; //we are only processing requests

    const email = Spreadsheets.getRequestsSheetHandler().getAttribute('email', i);
    const student = studentsDict[email]; //get student who submitted the request
    student.requestRowIndex = i;

    const purpose = Spreadsheets.getRequestsSheetHandler().getAttribute('purpose', i);
    student.purpose = purpose;

    const destinationRoomName = Spreadsheets.getRequestsSheetHandler().getAttribute('destination', i);
    const destinationRoom = roomsDict[destinationRoomName];
    if (!destinationRoom)
    {
      Logger.log("WARNING: Destination " + destinationRoomName + " not found for student " + student.name);
      continue;
    }
    if (destinationRoomName == student.homeroomName)
    {
      Logger.log("Destination same as homeroom... ignoring request.");
      student.setRequestStatus('Ignored,DestinationSameAsHomeroom');
      student.setRequestBackground('Aquamarine');
      student.sendEmail('Embedded Time Request', 'Hello, ' + student.name + '. There has likely been a problem with your e-silverslip request.\nPlease follow the following steps IN ORDER:\n\nStep 1: Fill out this form with your REGULAR Embedded Time room:\nhttps://docs.google.com/forms/d/e/1FAIpQLSeXbY519X--sQkTyqJmBshfLS35XIqsXxA-sUPmqgqo3Js5qg/viewform\n\nStep 2: Fill out this form with the Embedded Time room you want to go to:\nhttps://docs.google.com/forms/d/e/1FAIpQLSdlyrqakU_EvQXIYfWDnb4MdXqXtznSwdNCVc-xpegAhH-7_g/viewform\n\nThank you!');
      continue;
    }
    student.moveToRoom(destinationRoom);
  }


  //get existing approval sheets

  const cd = getCurrentDirectory();
  const approvalFolder = getOrCreateFolderInDirectory("Approval Sheets", cd);

  //dict of all the files in the approval folder; key: file name, value: sheet file
  const filesDict = getDictOfSpreadsheetFilesInFolder(approvalFolder);

  //sheet listing urls of all teacher approval sheets
  const approvalUrlSheet = SpreadsheetApp.open(getOrCreateSpreadsheetFileInDirectory('Approval Sheet Urls', cd)).getActiveSheet();
  const approvalUrlSsBuilder = new SpreadsheetBuilder(3,0);
  approvalUrlSsBuilder.addRow(['Tip: Press Ctrl + f to search for your room more easily']);

  //write to approval sheets
  for (const room of Object.values(roomsDict))
  {
    Logger.log("Writing " + room.name);
    var spreadsheet = null;
    var initFile = true;
    if (filesDict[room.name])
    {
      spreadsheet = SpreadsheetApp.open(filesDict[room.name]);
      initFile = false;
    }
    else
    {
      //create new approval sheet
      spreadsheet = SpreadsheetApp.open(createNewSpreadsheetFile(room.name, approvalFolder));
      initFile = true;
    }
    approvalUrlSsBuilder.addRow([room.location, room.name, spreadsheet.getUrl()]);

    //write to approval sheet while sending denails to students who were denied and removing them from the sheet

    const sheet = spreadsheet.getActiveSheet();
    const reader = new ApprovalSheetReader(sheet, studentsDict);

    const fullDataRange = sheet.getDataRange();
    const ssBuilder = new SpreadsheetBuilder(fullDataRange.getNumColumns(), 0);
    ssBuilder.addRow(["Student Email:", "Student Name:", "Student Purpose:", "Approval:"])
    for (const student of Object.values(room.studentsArrivedDict)) //we want to process all student requests to this room
    {
      if (reader.studentsDeniedDict[student.email]) //if the request was denied, we want to do denial process
      {
        MailApp.sendEmail(student.email, 'Silver slip request denied', 'Your silver slip request for ' + room.name + ' was unfortunately denied. Please go to back to your regular Embedded Time homeroom for this embedded time.');
        student.setRequestStatus('DeniedByTeacher');
        student.setRequestBackground('Crimson');
      }
      else
      {
        ssBuilder.addRow([student.email, student.name, student.purpose, 'Approve']);
      }
    }
    ssBuilder.fillCellRowsUntilLength(fullDataRange.getNumRows());
    sheet.getRange(1,1, ssBuilder.cellRowsData.length, ssBuilder.cellRowLength).setValues(ssBuilder.cellRowsData);
    sheet.autoResizeColumns(1, ssBuilder.cellRowLength);
    const protection = sheet.protect();
    if (initFile)
    {
      Logger.log("New Sheet: Initializing...");
      //add teachers in the room as editors
      protection.removeEditors(protection.getEditors());
      spreadsheet.addEditors(Object.keys(room.teachersDict));
    }
    //unprotect last column (status/approval column) starting on the second row (where the data starts)
    protection.setUnprotectedRanges([sheet.getRange(2, ssBuilder.cellRowLength, ssBuilder.cellRowsData.length, 1)]);
  }
  //write to approval sheet list (where teachers can view the urls of their approval sheets)
  approvalUrlSsBuilder.fillCellRowsUntilLength(approvalUrlSheet.getDataRange().getNumRows());
  approvalUrlSheet.getRange(1, 1, approvalUrlSsBuilder.cellRowsData.length, approvalUrlSsBuilder.cellRowLength).setValues(approvalUrlSsBuilder.cellRowsData);
  approvalUrlSheet.autoResizeColumns(1,1);

  //upload changes to request sheet (status values might have changed if teacher denied)
  Spreadsheets.getRequestsSheetHandler().uploadChanges();

  Logger.log("Done updating approval sheets.");
  Logger.log("Flushing...");
  SpreadsheetApp.flush();
}

function processRequests()
{
  //spreadsheet readers
  //just use the approval sheets and send emails / finalize attendance through there
  
  //gotta create rooms teachers students
  const studentsDict = Object.create(null); //indexed by email
  const teachersDict = Object.create(null); //indexed by email
  const roomsDict = Object.create(null); //indexed by room name
  createRoomsWithTeachers(roomsDict, teachersDict);
  createStudentsInRooms(studentsDict, roomsDict);

  markInvalidRequests(studentsDict);

  //for statistics (amount of people in each room before any requests are processed)
  for (const room of Object.values(roomsDict))
  {
    room.defaultOccupancy = room.occupantCount;
  }

  const cd = getCurrentDirectory();
  const approvalFolder = getOrCreateFolderInDirectory("Approval Sheets", cd);
  const filesDict = getDictOfSpreadsheetFilesInFolder(approvalFolder);

  //move approved students / email denial emails for each room
  for (const room of Object.values(roomsDict))
  {
    Logger.log("Processing " + room.name);
    const approvalSheetFile = filesDict[room.name];
    if (!approvalSheetFile)
    {
      Logger.log("Approval Sheet for Room " + room.name + "Not Found... This should not happen. Was it deleted by accident? Or did a teacher only just recently make a room during Embedded Time? Ooooh, you messed up (maybe) ;). It should be all good for next embedded time, though. It's just that all requests for that room won't go through for this Embedded Time.");
      continue;
    }
    else
    {
      const sheet = SpreadsheetApp.open(approvalSheetFile);
      const reader = new ApprovalSheetReader(sheet, studentsDict);
      //approve / deny students in room
      for (const student of Object.values(reader.studentsApprovedDict))
      {
        student.moveToRoom(room);
        student.sendEmail('Silver Slip Request Approved', 'Hello, ' + student.name +  '. Your silver slip request has been approved.\nPlease arrive at your regular Embedded Time room for attendance and then go to ' + room.name + " at " + room.location + " this Embedded Time. Thank you.");
        student.setRequestStatus('Granted');
      }
      for (const student of Object.values(reader.studentsDeniedDict))
      {
        student.sendEmail('Silver slip request denied', 'Your silver slip request for ' + room.name + ' was unfortunately denied. Please go to back to your regular Embedded Time homeroom for this embedded time.');
        student.setRequestStatus('DeniedByTeacher');
      }
    }
  }
  Spreadsheets.getRequestsSheetHandler().uploadChanges();
  
  const attendanceFolder = getOrCreateFolderInDirectory('Attendance Sheets', cd);
  createAttendanceSheets(attendanceFolder, roomsDict);

  Logger.log("Done processing!");
  updateStatisticsSheet('Statistics', 'Statistics', roomsDict);
  
  const pastRequestsFolder = getOrCreateFolderInDirectory('Past Requests', cd);
  saveAndClearRequestsSheet(pastRequestsFolder);

  Logger.log("Flushing...");
  SpreadsheetApp.flush();
}

function manuallySaveAndClearRequestsSheet()
{
  const cd = getCurrentDirectory();
  const pastRequestsFolder = getOrCreateFolderInDirectory('Past Requests', cd);
  saveAndClearRequestsSheet(pastRequestsFolder);
}

/* To Do:
Add sheet names as property to Spreadsheets object (don't want to hard-code) - Done
Add studentsLeft and studentsJoined dicts to room, key = email, value = student - Done
Allow teachers to request students
Add a function to add teachers in teachers sheet as editors (collaborators) for the Google Sheet - Done
Remove everything related to id (not using id anymore) - Done, I think
Change Spreadsheets properties to get functions so that we can instantiate the Sheet Handler if the instance is null - Done
IMPORTANT: Take (ROOM LOCATION) out from the student registration - Done
IMPORTANT: Make sure that the parameters of the sheet are correct (rows corresponding to attribute name for SheetHandler instances) - Done but keep checking

Mail namespace:
Add emails to a queue, batch send emails at the end, handle exception for quota reached
Enum for mails to send

Add a sheet listing all the rooms and their approval sheets (should be ez, just use Spreadsheet.getUrl()) - Done
Send emails to the denied students in markInvalidRequests but skip the ones already marked denied. - Done
Make a list of the # of requests each room had and make it public.
Lock e-silver slip request form at 9am, unlock after lunch (remember to lock, unlock function?) - Done
Add student purpose to approval sheets. - Done, I think
Remember to delete existing approval sheets to generate new ones for the teachers (after changing the sheet building)
Notify students of denial on updateApprovalSheets - Done
Pretty Important - Make Lock so that form submits don't override each other
Pretty Extensive - Make a Sheet holding all Embedded Time times and a system that runs it during the time (maybe like nodes)
Unclear about if students go to their reg ET first, make email more clear
1. Teachers can see where they come from and where they're going to - Done
2. Put Registration Form link on the Request Form - nah
3. 
Save the attendance sheets, new one each week
*/
