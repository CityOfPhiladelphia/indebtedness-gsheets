var dataSheetName = "Indebtedness Check"
  ,templateSheetName = "Template"
  ,templateRange = "A1"
  ,departmentColumn = "Hiring Department"
  ,departmentAllValue = "All Departments"
  ,letterRequiredColumn = "Need Letter?"
  ,letterRequiredValue = "Yes"
  ,dateLetterSentColumn = "Date Letter Sent"

/**
 * Prompt user for department and other fields before running the merge
 */
function prompt() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
    ,dataSheet = ss.getSheetByName(dataSheetName)
    ,templateSheet = ss.getSheetByName(templateSheetName)
    ,template = templateSheet.getRange(templateRange).getValue()
    ,app = UiApp.createApplication().setTitle("Merge Details")
    ,form = app.createFormPanel()
    ,flow = app.createFlowPanel()
    ,i
    ,elements = [
      ,app.createLabel("Select Department")
      ,createListBoxFromRange(app, ss.getRangeByName("Departments")).setName("department").addItem(departmentAllValue)
      ,app.createLabel("Template")
      ,app.createTextArea().setText(template).setSize("100%", "180px").setName("template")
      ,app.createSubmitButton("Submit").setStyleAttribute("display", "block")
    ];
  for(i in elements) {
    flow.add(elements[i]);
  }
  form.add(flow);
  app.add(form);
  ss.show(app);
}

/**
 * Called when user presses submit on prompt
 * Execute the merge
 */
function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
    ,department = e.parameter.department
    ,dataSheet = ss.getSheetByName(dataSheetName)
    ,dataRange = dataSheet.getRange(2, 1, dataSheet.getMaxRows() - 1, dataSheet.getMaxColumns())
    ,template = e.parameter.template
    ,objects = getRowsData(dataSheet, dataRange) // Create one object per row of data
    ,docTitle = "Merge (" + new Date().toString() + ")"
    ,doc = DocumentApp.create(docTitle)
    ,docBody = doc.getBody()
    ,normalizedColumns = {
      department: normalizeHeader(departmentColumn)
      ,letterRequired: normalizeHeader(letterRequiredColumn)
      ,dateLetterSent: normalizeHeader(dateLetterSentColumn)
    }
    ,i, contents;
  
  // For every row object, create a page in the new document
  Logger.log("Reviewing " + objects.length + " rows!");
  Logger.log(normalizedColumns);
  for(i = 0; i < objects.length; i++) {
    // Correct department, letter required, and date letter sent empty?
    if((department == departmentAllValue || objects[i][normalizedColumns.department] == department)
        && objects[i][normalizedColumns.letterRequired] == letterRequiredValue
        && ! objects[i][normalizedColumns.dateLetterSent]) {
      contents = fillTemplate(template, objects[i]);
      Logger.log(contents);
      docBody.appendParagraph(contents).appendPageBreak();
    } else Logger.log(objects[i][normalizedColumns.department] + " != " + department);
  }
  
  // Save & close document
  doc.saveAndClose();
  
  // When completed, alert the user of the link to the new document
  ss.show(confirmation(doc.getUrl(), docTitle));
}

/**
 * Show confirmation box with URL to new document
 */
function confirmation(url, title) {
  var app = UiApp.getActiveApplication().setTitle("Running Merge");
  app.add(app.createAnchor(title, url));
  return app;
}

/**
 * Add menu button
 */
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Run Merge",
    functionName : "prompt"
  }];
  ss.addMenu("Merge", entries);
};
