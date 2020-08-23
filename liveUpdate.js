
function onEdit(e) {
  let { range, value, oldValue } = e;
  let page = ss.getActiveSheet();
  let pageName = page.getName();
  let cell = page.getActiveCell();
  let colIndex = cell.getColumn();
  let colLetter = columnToLetter(colIndex);
  let rowIndex = cell.getRow();
  let colName = page.getRange(1,colIndex).getValue();
  let lastRow = page.getLastRow();
  let columnRange = page.getRange(`${colLetter}2:${colLetter}${lastRow}`);
  let columnValues = columnRange.getValues();
  let duplNumber = 0;
  
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

