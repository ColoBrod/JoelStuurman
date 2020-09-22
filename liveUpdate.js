function onEditInstallable(e) {
  let { range, value, oldValue } = e;
  let page = ss.getActiveSheet();
  let pageName = page.getName();
  let cell = page.getActiveCell();
  let colIndex = cell.getColumn();
  let colLetter = columnToLetter(colIndex);
  let rowIndex = cell.getRow();
  let colName = page.getRange(1,colIndex).getValue();
  let lastRow = page.getLastRow();
  let lastColumn = page.getLastColumn();
  let columnRange = page.getRange(`${colLetter}2:${colLetter}${lastRow}`);
  let columnValues = columnRange.getValues();
  let duplNumber = 0;
  
  // Create new contract
  if ( pageName == TAB.CONTRACTS 
    && colName == COLUMN.CONTRACTS.CREATE_NEW_CONTRACT
  ) {
    let titles, row, contract = {}; 
    // Titles of the TAB Contracts
    titles = page.getRange(1,1,1,lastColumn).getValues()[0];
    // Current row
    row = page.getRange(rowIndex,1,1,lastColumn).getValues()[0];
    // Contract forms from this object:
    for (let i in titles) {
      switch (titles[i]) {
        case COLUMN.CONTRACTS.REFERENCE:
          contract.fileName = row[i];
          break;
        case COLUMN.CONTRACTS.PROPERTY:
          contract.property = row[i];
          break;
        case COLUMN.CONTRACTS.PROPERTY_CONTACT:
          contract.lessorFullName = row[i];
          break;
        case COLUMN.CONTRACTS.CLIENT:
          contract.lesseeFullName = row[i];
          break;
        case COLUMN.CONTRACTS.PROPERTY_TYPE:
          contract.propertyType = row[i];
          break;
        case COLUMN.CONTRACTS.CONTRACT_START_DATE:
          contract.leaseStart = row[i];
          break;
        case COLUMN.CONTRACTS.CONTRACT_END_DATE:
          contract.leaseEnd = row[i];
          break;
        case COLUMN.CONTRACTS.CONTRACT_DURATION:
          contract.contractDuration = row[i];
          break;
        case COLUMN.CONTRACTS.INVENTORY_LIST:
          contract.inventoryList = row[i];
          break;
        case COLUMN.CONTRACTS.PARKING_SLOT_NO:
          contract.parkingSlot = row[i];
          break;
        case COLUMN.CONTRACTS.ADVANCE_APPLICABLE_ON:
          contract.advanceApplicableOn = row[i];
          break;
        case COLUMN.CONTRACTS.DEPOSIT:
          contract.deposit = row[i];
          break;
        case COLUMN.CONTRACTS.REMAINING_PAYMENT:
          contract.remainingPayment = row[i];
          break;
        case COLUMN.CONTRACTS.PET_CLAUSE:
          contract.petClause = row[i];
          break;
        case COLUMN.CONTRACTS.ASSOCIATION_DUES:
          contract.associationDues = row[i];
          break;
        case COLUMN.CONTRACTS.GROSS_SALES_RATE:
          contract.leaseRate = row[i];
          break;
      }
    }
    let propertyTab = ss.getSheetByName(TAB.PROPERTIES);
    let propertyLastColumn = propertyTab.getLastColumn();
    let propertyLastRow = propertyTab.getLastRow();
    let nameAutoFillIndex = 
      getColumnIndex(TAB.PROPERTIES, COLUMN.PROPERTIES.NAME_AUTOFILL);
    let nameAutoFill = propertyTab
      .getRange(2,nameAutoFillIndex,propertyLastRow-1,1)
      .getValues();
    let propertyRowIndex;
    for (let i in nameAutoFill) {
      if (nameAutoFill[i][0] == contract.property) {
        propertyRowIndex = parseInt(i)+2;
        break;
      }
    }
    let propertyTitles = propertyTab
      .getRange(1,1,1,propertyLastColumn)
      .getValues()[0];
    let propertyRow = propertyTab
      .getRange(propertyRowIndex, 1, 1, propertyLastColumn)
      .getValues()[0];
    for (let i in propertyTitles) {
      switch (propertyTitles[i]) {
        case COLUMN.PROPERTIES.BUILDING_VILLAGE:
          contract.propertyName = propertyRow[i];
          break;
        case COLUMN.PROPERTIES.UNIT:
          contract.propertyNo = propertyRow[i];
          break;
        case COLUMN.PROPERTIES.TOWER_STREET:
          contract.propertySection = propertyRow[i];
          break;
        case COLUMN.PROPERTIES.PROPERTY_STREET:
          contract.propertyStreet = propertyRow[i];
          break;
        case COLUMN.PROPERTIES.PROPERTY_ADDRESS:
          contract.propertyAddress = propertyRow[i];
          break;
      }
    }
    let contactTab = ss.getSheetByName(TAB.CONTACTS);
    let contactLastColumn = contactTab.getLastColumn();
    let contactLastRow = contactTab.getLastRow();
    let contactFullNameIndex = 
      getColumnIndex(TAB.CONTACTS, COLUMN.CONTACTS.FULL_NAME);
    let contactFullName = contactTab
      .getRange(2,contactFullNameIndex,contactLastRow-1,1)
      .getValues();
    let lessorRowIndex;
    let lesseeRowIndex;
    for (let i in contactFullName) {
      if (contactFullName[i][0] == contract.lessorFullName)
        lessorRowIndex = parseInt(i)+2;
      else if (contactFullName[i][0] == contract.lesseeFullName)
        lesseeRowIndex = parseInt(i)+2;
      if (lessorRowIndex && lesseeRowIndex) break;
    }
    let contactTitles = contactTab
      .getRange(1,1,1,contactLastColumn)
      .getValues()[0];
    let lessorRow = contactTab
      .getRange(lessorRowIndex, 1, 1, contactLastColumn)
      .getValues()[0];
    let lesseeRow = contactTab
      .getRange(lesseeRowIndex, 1, 1, contactLastColumn)
      .getValues()[0];
    for (let i in contactTitles) {
      switch (contactTitles[i]) {
        case COLUMN.CONTACTS.NATIONALITY:
          contract.lessorNationality = lessorRow[i];
          contract.lesseeNationality = lesseeRow[i];
          break;
        case COLUMN.CONTACTS.ADDRESS:
          contract.lessorAddress = lessorRow[i];
          contract.lesseeAddress = lesseeRow[i];
          break;
        case COLUMN.CONTACTS.ID_LINK:
          contract.lessorIdLink = lessorRow[i];
          contract.lesseeIdLink = lesseeRow[i];
          break;
      }
    }

    // CREATE FILE AND BUILD NEW CONTRACT:
    const contractDoc = new ContractDocument(contract.fileName);
    let url = contractDoc.build(
      contract.lessorFullName,
      contract.lessorNationality,
      contract.lessorAddress,
      contract.lessorIdLink,
      contract.lesseeFullName,
      contract.lesseeNationality,
      contract.lesseeAddress,
      contract.lesseeIdLink,
      contract.propertyNo,
      contract.propertySection,
      contract.propertyName,
      contract.propertyStreet,
      contract.propertyAddress,
      contract.propertyType,
      contract.leaseStart,
      contract.leaseEnd,
      contract.contractDuration,
      contract.inventoryList,
      contract.leaseRate,
      contract.parkingSlot,
      contract.advanceApplicableOn,
      contract.deposit,
      contract.petClause,
      contract.associationDues,
      contract.remainingPayment
    );
    // CHECKBOX:
    let chbox = page.getRange(rowIndex, colIndex);
    chbox.setValue("FALSE");
    // Update contract URL:
    let contractUrl = { rowIndex }
    contractUrl.colIndex = 
      getColumnIndex(pageName, COLUMN.CONTRACTS.CONTRACT_URL)
    contractUrl.range = 
      page.getRange(contractUrl.rowIndex, contractUrl.colIndex);
    contractUrl.range.setValue(url);
  }

  // Update Contact ID
  if (pageName == TAB.CONTACTS && colName == COLUMN.CONTACTS.ID_ADD) {
    contactId().openForm(rowIndex, colIndex);
  }

  // Multiple choice for following columns:
  if (  (
          pageName == TAB.CONTRACTS && 
          [COLUMN.CONTRACTS.KAM, COLUMN.CONTRACTS.LANDLORD, COLUMN.CONTRACTS.PROPERTY, COLUMN.CONTRACTS.CLIENT, COLUMN.CONTRACTS.PROPERTY_CONTACT]
            .includes(colName)
        )
    ||  (
          pageName == TAB.PROPERTIES && 
          [COLUMN.PROPERTIES.PROPERTY_CONTACT, COLUMN.PROPERTIES.LANDLORD]
            .includes(colName)
        ) 
  ) {
    let arr1;
    let arr2;
    let newArr = [];
    if (oldValue) arr1 = oldValue.trim().split(/\s*\;\s*/);
    else arr1 = [];
    if (value) arr2 = value.trim().split(/\s*\;\s*/);
    else arr2 = [];
    if (!value) cell.setValue("");
    else {
      newArr = [...arr1];
      for (let i in arr2) {
        if (newArr.includes(arr2[i])) continue;
        newArr.push(arr2[i]);
      }
      cell.setValue(newArr.join("; "));
    }
    return;
  }

  // Concatenate 3 cells and update Property name:
  let colsToConcat = [COLUMN.PROPERTIES.BUILDING_VILLAGE, COLUMN.PROPERTIES.TOWER_STREET, COLUMN.PROPERTIES.UNIT];
  let foundIndex = colsToConcat.indexOf(colName);
  if ( pageName == TAB.PROPERTIES && foundIndex > -1 ) {
    let oldConcatenated = "";
    let newConcatenated = "";
    for (let i in colsToConcat) {
      if (i == foundIndex) {
        oldConcatenated += ( oldValue ? oldValue : "" );
        newConcatenated += ( value ? value : "" )
        appendSpace();
        continue;
      }
      let colContent = getColumnContent(pageName, colsToConcat[i]);
      let cellContent = colContent[rowIndex-2];
      oldConcatenated += cellContent;
      newConcatenated += cellContent;
      appendSpace();
      function appendSpace() { 
        if (i != colsToConcat.length - 1) {
          oldConcatenated += " "; 
          newConcatenated += " ";
        }
      }
    }
    // Browser.msgBox(`${oldConcatenated} | ${newConcatenated}`)
    updateName({ sheet: TAB.CONTRACTS, column: COLUMN.CONTRACTS.PROPERTY }, oldConcatenated, newConcatenated)
  }

  // Update all obsolete entries
  if ( colName
    && link[pageName] 
    && Object.keys(link[pageName]).length > 0
    && link[pageName][colName]
  ) {
    // Update 
    if (oldValue && value) {
      for (let update of link[pageName][colName]) {
        updateName(update, oldValue, value);
      }
    }
    // Check for DDL duplicates
    for (let i = 0; i < columnValues.length && duplNumber < 2; i++) {
      if (columnValues[i][0] == value) duplNumber++;
    }
    if (duplNumber > 1) {
      duplicateFound(range, value, oldValue);
      return;
    }
  }

}

function duplicateFound(cell, newValue, oldValue) {
  cell.setValue(oldValue);
  Browser.msgBox(`Record "${newValue}" already exists in column.`);
}

function updateName(update, oldValue, newValue) {
  let from = update.from ? update.from : null;
  // TODO: create a method to update from another column
  let sheet = ss.getSheetByName(update.sheet);
  let lastColumn = sheet.getLastColumn();
  let titleRange = sheet.getRange(1, 1, 1, lastColumn).getValues();
  let titles = titleRange[0];
  let colIndex = titles.indexOf(update.column) + 1;
  let colLetter = columnToLetter(colIndex);
  let lastRow = sheet.getLastRow();
  let range = sheet.getRange(2, colIndex, lastRow-1, 1);
  let value = range.getValues();
  let arr = [];

  for (let i in value) {
    let content = value[i][0];
    let regex = new RegExp(`^((?:\\s*[^;]*\\s*;\\s*)*)${oldValue}((?:\\s*;\\s*[^;]*\\s*)*)$`);
    let regexStr = regex.toString()
    // Browser.msgBox(regexStr)
    if (content.match(regex)) {
      let newStr = content.replace(regex, `$1${newValue}$2`);
      arr[i] = [newStr]
    }
    else {
      arr[i] = [content];
    }
  }
  range.setValues(arr);
}

function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function letterToColumn(letter) {
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

function getColumnIndex(sheetName, header) {
  let sheet = ss.getSheetByName(sheetName)
  let lastColumn = sheet.getLastColumn();
  let range = sheet.getRange(1,1,1,lastColumn);
  let values = range.getValues()[0];
  let index = values.indexOf(header) + 1;
  return index;
}
