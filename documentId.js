function contactId() {
  return ({
    openForm: function (rowIndex, colIndex) {
      let sheet = ss.getSheetByName(TAB.CONTACTS);
      // let colIndex = getColumnIndex(TAB.CONTACTS, COLUMN.CONTACTS.ID_ADD);
      let chbox = sheet.getRange(rowIndex, colIndex);
      
      // At the end of execution restore button style
      chbox.setValue("FALSE");
    },
    updateLink: function () {

    },
    
  })
}