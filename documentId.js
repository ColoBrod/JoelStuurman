function contactId() {
  return ({
    
    openForm: function (rowIndex, colIndex) {
      let sheet = ss.getSheetByName(TAB.CONTACTS);
      // let colIndex = getColumnIndex(TAB.CONTACTS, COLUMN.CONTACTS.ID_ADD);
      let chbox = sheet.getRange(rowIndex, colIndex);
      let urlColIndex = getColumnIndex(TAB.CONTACTS, COLUMN.CONTACTS.ID_LINK);
      let idUrl = sheet.getRange(rowIndex, urlColIndex).getValue();
      // Opening an HTML form that uploads image
      let template = HtmlService.createTemplateFromFile("UploadId.html");
      template.rowIndex = rowIndex;
      template.idUrl = idUrl;
      let html = template.evaluate();
      html.setWidth(800).setHeight(400);
      ui.showModalDialog(html, "Upload ID");
      // At the end of execution restore button style
      chbox.setValue("FALSE");
    },

    updateLink: function (rowIndex, url) {
      let sheet = ss.getSheetByName(TAB.CONTACTS);
      let colIndex = getColumnIndex(TAB.CONTACTS, COLUMN.CONTACTS.ID_LINK);
      let range = sheet.getRange(rowIndex, colIndex);
      range.setValue(url);
    },
    
  })
}

function updateIdLink(rowIndex, url) {
  contactId().updateLink(rowIndex, url);
}