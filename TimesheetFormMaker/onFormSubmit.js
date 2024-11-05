function onFormSubmit(e) {
  var responses = e.namedValues;

  var title = responses['Title of the timesheet to be made'][0];
  var startDate = new Date(responses['Start date of the timesheet'][0]);
  var endDate = new Date(responses['End date of the timesheet'][0]);
  var companyName = responses['Name of the company'][0];
  var rotationLayout = responses['Rotation layout'][0];
  var emailAddress = responses['Your Email Address'][0];

  // Create a new Google Doc
  var doc = DocumentApp.create(title);
  var body = doc.getBody();

  // Set the document orientation
  if (rotationLayout === 'Landscape') {
    body.setPageOrientation(DocumentApp.PageOrientation.LANDSCAPE);
  }

  // Get weeks in the range
  var weeks = getWeeksInRange(startDate, endDate);

  weeks.forEach(function(weekStartDate, index) {
    // Add a page break if not the first page
    if (index > 0) {
      body.appendPageBreak();
    }

    // Add Company Name
    body.appendParagraph(companyName).setHeading(DocumentApp.ParagraphHeading.HEADING1);

    // Add Week Start Date
    body.appendParagraph('Week Commencing: ' + formatDate(weekStartDate)).setHeading(DocumentApp.ParagraphHeading.HEADING2);

    // Add Name field
    body.appendParagraph('Name: ____________________');

    // Add a table with headings
    var table = body.appendTable();
    var headerRow = table.appendTableRow();
    var headers = ['Date', 'Worksite Address', 'Start Time', 'End Time', 'Lunch Duration', 'Basic hrs', 'O/T 1.5', 'O/T 2.0'];
    headers.forEach(function(header) {
      headerRow.appendTableCell(header);
    });

    // Get days of the week
    var daysOfWeek = getDaysOfWeek(weekStartDate);

    daysOfWeek.forEach(function(date) {
      if (date > endDate) return; // Stop if date exceeds end date
      var row = table.appendTableRow();
      var dateStr = formatDayDate(date);
      row.appendTableCell(dateStr);
      // Append empty cells for other columns
      for (var i = 1; i < headers.length; i++) {
        row.appendTableCell('');
      }
    });
  });

  // Send email with the document link
  var docUrl = doc.getUrl();
  var subject = 'Your Generated Timesheet Document';
  var message = 'Hello,\n\nYour timesheet document has been created. You can access it using the following link:\n' + docUrl + '\n\nBest regards.';
  MailApp.sendEmail(emailAddress, subject, message);

  Logger.log('Document URL: ' + docUrl);
}

function getWeeksInRange(startDate, endDate) {
  var weeks = [];
  var currentDate = new Date(startDate);

  // Adjust to the start of the week (Monday)
  var day = currentDate.getDay();
  var diff = (day === 0) ? -6 : 1 - day;
  currentDate.setDate(currentDate.getDate() + diff);

  while (currentDate <= endDate) {
    weeks.push(new Date(currentDate));
    currentDate.setDate(currentDate.getDate() + 7);
  }

  return weeks;
}

function getDaysOfWeek(weekStartDate) {
  var days = [];
  for (var i = 0; i < 7; i++) {
    var date = new Date(weekStartDate);
    date.setDate(date.getDate() + i);
    days.push(date);
  }
  return days;
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMMM dd, yyyy');
}

function formatDayDate(date) {
  var dayName = Utilities.formatDate(date, Session.getScriptTimeZone(), 'EEE');
  var day = date.getDate();
  var suffix = getOrdinalSuffix(day);
  return dayName + ' ' + day + suffix;
}

function getOrdinalSuffix(day) {
  if (day > 3 && day < 21) return 'th';
  switch (day % 10) {
    case 1: return 'st';
    case 2: return 'nd';
    case 3: return 'rd';
    default: return 'th';
  }
}