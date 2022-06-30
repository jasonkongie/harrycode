function approvalSheetsTimeTrigger() //runs every 15 mins
{
  var timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH");
  Logger.log("Hour: " + timeStamp);
  if (timeStamp >= 6 && timeStamp <= 22) //between 6 am and 11 pm
  {
    updateApprovalSheets();
  }
  else
  {
    Logger.log("Not a active time...");
  }
  return;
}

function lockRequestFormTimeTrigger() //runs every Tuesday and Thursday at 9:15am +- 15 min
{
  Logger.log("Locking request form...");
  const requestForm = FormApp.openByUrl(FormUrls.studentRequestForm);
  requestForm.setAcceptingResponses(false);
}

function unlockRequestFormTimeTrigger() //runs every Tuesday and Thursday at 12:45 +- 15 min
{
  Logger.log("Unlocking request form...");
  const requestForm = FormApp.openByUrl(FormUrls.studentRequestForm);
  requestForm.setAcceptingResponses(true);
}

function processRequestsTimeTrigger() //runs every Tuesday and Thursday at 10:45am +- 15 min
{
  processRequests();
}
