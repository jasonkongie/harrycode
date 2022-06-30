
//we need a formsubmit trigger to copy the request form submission to our spreadsheet
function onRequestsFormSubmit(e)
{
  const formResponse = e.response;
  const timestamp = formResponse.getTimestamp();
  const email = formResponse.getRespondentEmail();
  const name = EmailToName[email];

  var answers = [];
  formResponse.getItemResponses().forEach(element => answers.push(element.getResponse().trim()));

  var row = [timestamp, email, name].concat(answers);
  Logger.log(row);

  var lock = LockService.getScriptLock();
  // Wait for up to 60 seconds for other processes to finish.
  lock.waitLock(60000);
  try
  {
    const destinationIndex = Spreadsheets.getRequestsSheetHandler().getAttributeColumnIndex('destination');
    var destination = row[destinationIndex];

    Logger.log('New Request: ' + row);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Spreadsheets.requestsSheetName);
    sheet.appendRow(row);
    SpreadsheetApp.flush();
  }
  catch (err)
  {
    try
    {
      const errorSheet = SpreadsheetApp.openByUrl(Spreadsheets.errorSheetUrl);
      errorSheet.appendRow(row);
      SpreadsheetApp.flush();
    }
    catch (err)
    {
      Logger.log(e);
      throw err;
    }
    throw err;
  }
  lock.releaseLock();
}

//we need a formsubmit trigger to copy the student form submission to our spreadsheet
function onStudentRegistrationFormSubmit(e)
{
  const formResponse = e.response;
  const timestamp = formResponse.getTimestamp();
  const email = formResponse.getRespondentEmail();
  const name = EmailToName[email];

  var answers = [];
  formResponse.getItemResponses().forEach(element => answers.push(element.getResponse().trim()));

  var row = [timestamp, email, name].concat(answers);
  Logger.log(row);

  var lock = LockService.getScriptLock();
  // Wait for up to 60 seconds for other processes to finish.
  lock.waitLock(60000);
  try
  {
    Logger.log('New Student Registration: ' + row);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Spreadsheets.studentsSheetName);
    sheet.appendRow(row);
    SpreadsheetApp.flush();
  }
  catch (err)
  {
    Logger.log(e);
    throw err;
  }
  lock.releaseLock();
}

//we need a formsubmit trigger to copy the student form submission to our spreadsheet
function onTeacherRegistrationFormSubmit(e)
{
  const formResponse = e.response;
  const timestamp = formResponse.getTimestamp();
  const email = formResponse.getRespondentEmail();

  var answers = [];
  formResponse.getItemResponses().forEach(element => answers.push(element.getResponse().trim()));

  var row = [timestamp, email].concat(answers);
  Logger.log(row);

  var lock = LockService.getScriptLock();
  // Wait for up to 60 seconds for other processes to finish.
  lock.waitLock(60000);
  try
  {
    Logger.log('New Teacher Registration: ' + row);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Spreadsheets.teachersSheetName);
    sheet.appendRow(row);
    SpreadsheetApp.flush();
  }
  catch (err)
  {
    Logger.log(e);
    throw err;
  }
  lock.releaseLock();
}

//no use
/*
function getTeacherEmailWithHomeroomName(homeroomName)
{
  const homeroomNameIndex = Spreadsheets.getTeachersSheetHandler().getAttributeColumnIndex('homeroomName');
  //find the row in the teachers sheet that created the homeroom
  for (var i = 0; i < Spreadsheets.getTeachersSheetHandler().sheetValues.length; i++) 
  {
    const row = Spreadsheets.getTeachersSheetHandler().sheetValues[i];
    if (row[homeroomNameIndex] == homeroomName) 
    {
      return row[Spreadsheets.getTeachersSheetHandler().getAttributeColumnIndex('email')]; //return the email
    }
  }
}
*/
