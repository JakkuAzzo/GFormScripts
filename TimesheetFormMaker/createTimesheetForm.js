function createTimesheetForm() {
  // Create a new form
  var form = FormApp.create('Timesheet Configuration Form');
  
  // Add a title question
  form.addTextItem()
      .setTitle('Title of the timesheet to be made')
      .setRequired(true);
  
  // Add a start date question
  form.addDateItem()
      .setTitle('Start date of the timesheet')
      .setRequired(true);
  
  // Add an end date question
  form.addDateItem()
      .setTitle('End date of the timesheet')
      .setRequired(true);
  
  // Add a company name question
  form.addTextItem()
      .setTitle('Name of the company')
      .setRequired(true);
  
  // Add a rotation layout question
  var item = form.addMultipleChoiceItem();
  item.setTitle('Rotation layout')
      .setChoices([
        item.createChoice('Portrait'),
        item.createChoice('Landscape')
      ])
      .setRequired(true);

  // Add an email address question
  form.addTextItem()
      .setTitle('Your Email Address')
      .setRequired(true)
      .setValidation(
        FormApp.createTextValidation()
          .requireTextIsEmail()
          .setHelpText('Please enter a valid email address.')
          .build()
      );
  
  Logger.log('Form URL: ' + form.getPublishedUrl());
}