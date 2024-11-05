function createForm() {
    // Create a new form
    var form = FormApp.create('New Form');
    
    // Add a multiple-choice question
    var item = form.addMultipleChoiceItem();
    item.setTitle('Is this a form?')
        .setChoices([
          item.createChoice('Yes'),
          item.createChoice('No')
        ]);
    
    // Add a text question
    form.addTextItem()
        .setTitle('Why?');
    
    Logger.log('Form URL: ' + form.getPublishedUrl());
  }