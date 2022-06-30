const FormUrls =
{
  studentRegistrationForm: 'https://docs.google.com/forms/d/19Gh9zSWby3x6xKKofmlXWe10D23Atl5Nozv8JC_kU2c/edit',
  studentRequestForm: 'https://docs.google.com/forms/d/1BvFAL_fOLa2fUJXoPCkKMatVQ51EhKpH5xepu-7sd3E/edit',
  teacherRegistrationForm: 'https://docs.google.com/forms/d/1IbMONRdTu1YGY_lBMDuHNysHehCvNNtBMBibqMUBW30/edit',
}

function addHomeroomDropdownForStudentRegistrationForm()
{
  const studentRegistrationForm = FormApp.openByUrl(FormUrls.studentRegistrationForm);

  var dropdown = studentRegistrationForm.addListItem();

  var dropdownChoices = [];
  for (var i = 1; i < Spreadsheets.teachersSheet.sheetValues.length; i++)
  {
    const choiceText = Spreadsheets.teachersSheet.getAttribute('homeroomName', i) + " (" + Spreadsheets.teachersSheet.getAttribute('location', i) + ")";
    dropdownChoices.push(dropdown.createChoice(choiceText));
  }

  dropdown.setTitle('Your Embedded Time Class')
    .setChoices(dropdownChoices)
    .setRequired(true);
}

function addHomeroomDropdownForStudentRequestForm()
{
  const studentRequestForm = FormApp.openByUrl(FormUrls.studentRequestForm);

  var dropdown = studentRequestForm.addListItem();

  var dropdownChoices = [];
  for (var i = 1; i < Spreadsheets.teachersSheet.sheetValues.length; i++)
  {
    const choiceText = Spreadsheets.teachersSheet.getAttribute('homeroomName', i) + " (" + Spreadsheets.teachersSheet.getAttribute('location', i) + ")";
    dropdownChoices.push(dropdown.createChoice(choiceText));
  }

  dropdown.setTitle('Which room would you like to go to?')
    .setChoices(dropdownChoices)
    .setRequired(true);
}
