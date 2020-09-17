const ss = SpreadsheetApp.getActiveSpreadsheet();
const ui = SpreadsheetApp.getUi();
const TAB = {
  CONTRACTS:  "Contracts",
  PAYMENTS:   "Payments",
  PROPERTIES: "Properties",
  CONTACTS:   "Contacts",
  DATA:       "Data",
};
const COLUMN = {
  CONTRACTS: {
    REFERENCE:            "Reference",
    ID:                   "ID",
    KAM:                  "KAM",
    TRANSACTION_TYPE:     "Transaction Type *Data*",
    CONTRACT_START_DATE:  "Contract Start Date",
    CONTRACT_END_DATE:    "Contract End Date",
    CONTRACT_DURATION:    "Contract Duration",
    LANDLORD:             "Landlord *Contacts*",
    GROSS_SALES_RATE:     "Gross Sales Rate",
    CONTRACT_STATUS:      "Contract Status",
    PROPERTY:             "Property *Properties*",
    CLIENT:               "Client *Contacts*",
    PROPERTY_CONTACT:     "Property contact *Contacts*",
    CONTRACT_URL:         "Contract URL",
    CREATE_NEW_CONTRACT:  "[+] Create new Contract",
    PET_CLAUSE:           "Pet Clause",
    INVENTORY_LIST:       "Inventory List",
    PARKING_SLOT_NO:      "Parking Slot No",
    ADVANCE_APPLICABLE_ON:"Advance applicable on",
    DEPOSIT:              "Deposit",
    REMAINING_PAYMENT:    "Remaining Payment",
    PROPERTY_TYPE:        "Property Type",
    ASSOCIATION_DUES:     "Association Dues",
  },
  PAYMENTS: {
    REFERENCE:            "Reference",
    ID:                   "ID",
    KAM:                  "KAM",
    TRANSACTION_TYPE:     "Transaction Type",
    CONTRACT_START_DATE:  "Contract Start Date",
    CONTRACT_END_DATE:    "Contract End Date",
    CONTRACT_DURATION:    "Contract Duration",
    PAYER:                "Payer",
    GROSS_SALES_RATE:     "Gross Sales Rate",
    DUES_PHP:             "Dues (Php)",
    NET_SALES_RATE:       "Net Sales Rate",
    COM_RATE_PERCENT:     "Com. Rate (%)",
    GROSS_COM:            "Gross Com.",
    DISCOUNT_PERCENT:     "Discount (%)",
    DISCOUNT_PHP:         "Discount (Php)",
    NET_COM:              "Net Com.",
    EXTERNAL_AGENT_GCI_PHP: "External Agent GCI (Php)",
    EXTERNAL_AGENT_GCI_PERCENT: "External Agent GCI (%)",
    EXTERNAL_AGENT_NAME:  "External Agent (Name)",
    PROPEL_GCI_PERCENT:   "Propel GCI (%)",
    PROPEL_GCI_PHP:       "Propel GCI (Php)",
    INTERNAL_AGENT_PHP:   "Internal Agent (Php)",
    INTERNAL_AGENT_NAME:  "Internal Agent (Name)",
    INTERNAL_AGENT_PERCENT: "Internal Agent (%)",
    COST_TYPE:            "Cost Type",
    PAYMENT_DUE_DATE:     "Payment Due Date",
    PAYMENT_AMOUNT_DUE:   "Payment Amount Due",
    PAYMENT_AMOUNT_RECEIVED: "Payment Amount Received",
    PAYMENT_RECORD_1_DATE: "Payment Record 1 (Date)",
    PAYMENT_STATUS:       "Payment Status",
  },
  PROPERTIES: {
    NAME_AUTOFILL:        "Name (Autofill)",
    TRANSACTION_TYPE:     "Transaction Type *Data*",
    KAM:                  "KAM *Data*",
    PROPERTY_CONNECTION:  "Property Connection *Data*",
    BUILDING_VILLAGE:     "Building / Village *Data*",
    TOWER_STREET:         "Tower / Street",
    UNIT:                 "Unit",
    UNIT_TYPE:            "Unit Type",
    SQM:                  "Sqm",
    FURNISHED:            "Furnished *Data*",
    RENT_MONTH:           "Rent/month (asking)",
    SALES_PRICE:          "Sales Price",
    PETFRIENDLY:          "Petfriendly",
    BALCONY:              "Balcony",
    PARKING:              "Parking",
    MAIDSROOM:            "Maidsroom",
    FACING_VIEW:          "Facing/View",
    COMMENT:              "Comment",
    URL_TO_PICTURES:      "URL to Pictures",
    PROPERTY_CONTACT:     "Property contact",
    LANDLORD:             "Landlord",
    PROPERTY_STREET:      "Property Street",
    PROPERTY_ADDRESS:     "Property Address",
  },
  CONTACTS: {
    FULL_NAME_AUTOFILL:   "FULL NAME (Autofill)",
    FULL_NAME:            "FULL NAME",
    KAM:                  "KAM *Data*",
    CONTACT_TYPE:         "Contact Type *Data*",
    NATIONALITY:          "Nationality",
    EMPLOYER:             "Emloyer",
    ADDRESS:              "Address",
    MOBILE:               "Mobile (+639xx with predial...)*",
    EMAIL:                "Email",
    ID_PICTURE:           "ID - picture",
    ID_LINK:              "ID - link",
    ID_ADD:               "[+] add ID",
  },
  DATA: {
    KAM:                  "KAM",
    TRANSACTION_TYPE:     "Transaction Type",
    CONTRACT_STATUS:      "Contract Status",
    PROPERTY_CONNECTION:  "Property Connection",
    BUILDING:             "Building",
    FURNISHING:           "Furnishing",
    CONTACT_TYPE:         "Contact Type",
    COST_TYPE:            "Cost Type",
    PAYABLE_ACCOUNT:      "Payable Account",
    REMAINING_PAYMENT:    "Remaining Payment",
    PROPERTY_TYPE:        "Property Type",
  },
};
const formHeader = {
  [TAB.CONTRACTS]: { 
    kam:                COLUMN.CONTRACTS.KAM,
    transactionType:    COLUMN.CONTRACTS.TRANSACTION_TYPE,
    contractStartDate:  COLUMN.CONTRACTS.CONTRACT_START_DATE,
    contractEndDate:    COLUMN.CONTRACTS.CONTRACT_END_DATE,
    landlord:           COLUMN.CONTRACTS.LANDLORD,
    grossSalesRate:     COLUMN.CONTRACTS.GROSS_SALES_RATE,
    contractStatus:     COLUMN.CONTRACTS.CONTRACT_STATUS,
    property:           COLUMN.CONTRACTS.PROPERTY,
    client:             COLUMN.CONTRACTS.CLIENT,
    propertyContact:    COLUMN.CONTRACTS.PROPERTY_CONTACT,
    contractUrl:        COLUMN.CONTRACTS.CONTRACT_URL,
    petClause:          COLUMN.CONTRACTS.PET_CLAUSE,
    inventoryList:      COLUMN.CONTRACTS.INVENTORY_LIST,
    parkingSlotNo:      COLUMN.CONTRACTS.PARKING_SLOT_NO,
    advanceApplicableOn: COLUMN.CONTRACTS.ADVANCE_APPLICABLE_ON,
    deposit:            COLUMN.CONTRACTS.DEPOSIT,
    remainingPayment:   COLUMN.CONTRACTS.REMAINING_PAYMENT,
    propertyType:       COLUMN.CONTRACTS.PROPERTY_TYPE,
  },
  [TAB.PAYMENTS]: { 
    reference:          COLUMN.PAYMENTS.REFERENCE,
    duesPhp:            COLUMN.PAYMENTS.DUES_PHP,
    netSalesRate:       COLUMN.PAYMENTS.NET_SALES_RATE,
    comRatePercent:     COLUMN.PAYMENTS.COM_RATE_PERCENT,
    grossCom:           COLUMN.PAYMENTS.GROSS_COM,
    discountPercent:    COLUMN.PAYMENTS.DISCOUNT_PERCENT,
    discountPhp:        COLUMN.PAYMENTS.DISCOUNT_PHP,
    netCom:             COLUMN.PAYMENTS.NET_COM,
    externalAgentGciPhp: COLUMN.PAYMENTS.EXTERNAL_AGENT_GCI_PHP,
    externalAgentGciPercent: COLUMN.PAYMENTS.EXTERNAL_AGENT_GCI_PERCENT,
    externalAgentName:  COLUMN.PAYMENTS.EXTERNAL_AGENT_NAME,
    propelGciPercent:   COLUMN.PAYMENTS.PROPEL_GCI_PERCENT,
    propelGciPhp:       COLUMN.PAYMENTS.PROPEL_GCI_PHP,
    internalAgentPhp:   COLUMN.PAYMENTS.INTERNAL_AGENT_PHP,
    internalAgentName:  COLUMN.PAYMENTS.INTERNAL_AGENT_NAME,
    internalAgentPercent: COLUMN.PAYMENTS.INTERNAL_AGENT_PERCENT,
    costType:           COLUMN.PAYMENTS.COST_TYPE,
    paymentDueDate:     COLUMN.PAYMENTS.PAYMENT_DUE_DATE,
    paymentRecord1Date: COLUMN.PAYMENTS.PAYMENT_RECORD_1_DATE,
  },
  [TAB.PROPERTIES]: { 
    transactionType:    COLUMN.PROPERTIES.TRANSACTION_TYPE,
    kam:                COLUMN.PROPERTIES.KAM,
    propertyConnection: COLUMN.PROPERTIES.PROPERTY_CONNECTION,
    buildingVillage:    COLUMN.PROPERTIES.BUILDING_VILLAGE,
    towerStreet:        COLUMN.PROPERTIES.TOWER_STREET,
    unit:               COLUMN.PROPERTIES.UNIT,
    unitType:           COLUMN.PROPERTIES.UNIT_TYPE,
    sqm:                COLUMN.PROPERTIES.SQM,
    furnished:          COLUMN.PROPERTIES.FURNISHED,
    rentMonth:          COLUMN.PROPERTIES.RENT_MONTH,
    salesPrice:         COLUMN.PROPERTIES.SALES_PRICE,
    petfriendly:        COLUMN.PROPERTIES.PETFRIENDLY,
    balcony:            COLUMN.PROPERTIES.BALCONY,
    parking:            COLUMN.PROPERTIES.PARKING,
    maidsroom:          COLUMN.PROPERTIES.MAIDSROOM,
    facingView:         COLUMN.PROPERTIES.FACING_VIEW,
    comment:            COLUMN.PROPERTIES.COMMENT,
    urlToPictures:      COLUMN.PROPERTIES.URL_TO_PICTURES,
    propertyContact:    COLUMN.PROPERTIES.PROPERTY_CONTACT,
    landlord:           COLUMN.PROPERTIES.LANDLORD,
    propertyStreet:     COLUMN.PROPERTIES.PROPERTY_STREET,
    propertyAddress:    COLUMN.PROPERTIES.PROPERTY_ADDRESS,
  },
  [TAB.CONTACTS]: { 
    fullName:           COLUMN.CONTACTS.FULL_NAME,
    kam:                COLUMN.CONTACTS.KAM,
    contactType:        COLUMN.CONTACTS.CONTACT_TYPE,
    nationality:        COLUMN.CONTACTS.NATIONALITY,
    employer:           COLUMN.CONTACTS.EMPLOYER,
    address:            COLUMN.CONTACTS.ADDRESS,
    mobile:             COLUMN.CONTACTS.MOBILE,
    email:              COLUMN.CONTACTS.EMAIL,
    idLink:             COLUMN.CONTACTS.ID_LINK,
  },
}

const link = {
  [TAB.CONTRACTS]: {
    [COLUMN.CONTRACTS.REFERENCE]: { sheet: TAB.PAYMENTS, column: COLUMN.PAYMENTS.REFERENCE, multiChoice: false },
  },
  [TAB.PAYMENTS]: {

  },
  [TAB.PROPERTIES]: {

  },
  [TAB.CONTACTS]: {
    [COLUMN.CONTACTS.FULL_NAME]: [
      { sheet: TAB.CONTRACTS, column: COLUMN.CONTRACTS.LANDLORD, multiChoice: true },
      { sheet: TAB.CONTRACTS, column: COLUMN.CONTRACTS.CLIENT, multiChoice: true },
      { sheet: TAB.CONTRACTS, column: COLUMN.CONTRACTS.PROPERTY_CONTACT, multiChoice: true },
      { sheet: TAB.PROPERTIES, column: COLUMN.PROPERTIES.PROPERTY_CONTACT, multiChoice: true },
      { sheet: TAB.PROPERTIES, column: COLUMN.PROPERTIES.LANDLORD, multiChoice: true },
      { sheet: TAB.PAYMENTS, column: COLUMN.PAYMENTS.EXTERNAL_AGENT_NAME, multiChoice: false },
      { sheet: TAB.PAYMENTS, column: COLUMN.PAYMENTS.INTERNAL_AGENT_NAME, multiChoice: false },
    ],
  },
  [TAB.DATA]: {
    [COLUMN.DATA.COST_TYPE]: [
      { sheet: TAB.PAYMENTS, column: COLUMN.PAYMENTS.COST_TYPE, multiChoice: false },
    ],
    [COLUMN.DATA.KAM]: [
      { sheet: TAB.CONTRACTS, column: COLUMN.CONTRACTS.KAM, multiChoice: true },
      { sheet: TAB.PROPERTIES, column: COLUMN.PROPERTIES.KAM, multiChoice: false },
      { sheet: TAB.CONTACTS, column: COLUMN.CONTACTS.KAM, multiChoice: false },
    ],
    [COLUMN.DATA.TRANSACTION_TYPE]: [
      { sheet: TAB.CONTRACTS, column: COLUMN.CONTRACTS.TRANSACTION_TYPE, multiChoice: false },
      { sheet: TAB.PROPERTIES, column: COLUMN.PROPERTIES.TRANSACTION_TYPE, multiChoice: false },
    ],
    [COLUMN.DATA.CONTRACT_STATUS]: [
      { sheet: TAB.CONTRACTS, column: COLUMN.CONTRACTS.CONTRACT_STATUS, multiChoice: false },
    ],
    [COLUMN.DATA.PROPERTY_CONNECTION]: [
      { sheet: TAB.PROPERTIES, column: COLUMN.PROPERTIES.PROPERTY_CONNECTION, multiChoice: false },
    ],
    [COLUMN.DATA.BUILDING]: [
      { sheet: TAB.PROPERTIES, column: COLUMN.PROPERTIES.BUILDING_VILLAGE, multiChoice: false },
    ],
    [COLUMN.DATA.FURNISHING]: [
      { sheet: TAB.PROPERTIES, column: COLUMN.PROPERTIES.FURNISHED, multiChoice: false },
    ],
    [COLUMN.DATA.CONTACT_TYPE]: [
      { sheet: TAB.CONTACTS, column: COLUMN.CONTACTS.CONTACT_TYPE, multiChoice: false },
    ],
    [COLUMN.DATA.REMAINING_PAYMENT]: [
      { sheet: TAB.CONTRACTS, column: COLUMN.CONTRACTS.REMAINING_PAYMENT },
    ],
    [COLUMN.DATA.PROPERTY_TYPE]: [
      { sheet: TAB.CONTRACTS, column: COLUMN.CONTRACTS.PROPERTY_TYPE },
    ],
  },
};

// Functions, that open sidebar with html-forms
const addContract = (clearHistory = true) => 
  openSidebar("Add a new contract", "Contracts", clearHistory);
const addPayment = (clearHistory = true) =>  
  openSidebar("Add a new payment", "Payments", clearHistory);
const addProperty = (clearHistory = true) => 
  openSidebar("Add a new property", "Properties", clearHistory);
const addContact = (clearHistory = true) =>  
  openSidebar("Add a new contact", "Contacts", clearHistory);

// Creating menu in the navigation bar
function onOpen(e) {
  createMenu();
}

// This function is required for including files to the sidebar
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function createMenu() {
  ui.createMenu("_SIDEBAR_FORM_")
    .addItem("Add Contract", "addContract")
    .addItem("Add Payment", "addPayment")
    .addItem("Add Property", "addProperty")
    .addItem("Add Contact", "addContact")
    .addToUi();
}

function openDialog(sheetName, value = "") {
  let dialogTitle;
  switch (sheetName) {
    case TAB.CONTRACTS:
      dialogTitle = "Edit contract";
      break;
    case TAB.PAYMENTS:
      dialogTitle = "Edit payment";
      break;
    case TAB.PROPERTIES:
      dialogTitle = "Edit property";
      break;
    case TAB.CONTACTS:
      dialogTitle = "Edit contact";
      break;
    default: 
      return;
  }
  let dialog = new Dialog("ModalWindow.html", dialogTitle, sheetName, value);
  dialog.show();
}

function openSidebar(sidebarTitle, headTitle, clearHistory = true) {
  let sidebar = new Sidebar("Form.html", sidebarTitle, headTitle, clearHistory);
  sidebar.show();
}

function addRecord(sheetName, content) {
  let sheet = ss.getSheetByName(sheetName);
  let lastColumn = sheet.getLastColumn();
  let lastRow = sheet.getLastRow();
  let newRow = lastRow+1;
  let titles = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  let arr = [];
  if (sheetName == TAB.CONTRACTS) content.ID = generateUniqueId();
  for (let i in titles) {
    let title = titles[i];
    let value = content[title] ? content[title] : "";
    arr.push(value);
  }
  sheet.appendRow(arr);
  // Columns which should be autofilled:
  let autoFill;
  switch (sheetName) {
    case TAB.CONTRACTS:
      autoFill = 
        [COLUMN.CONTRACTS.REFERENCE, COLUMN.CONTRACTS.CONTRACT_DURATION];
      break;
    case TAB.PAYMENTS:
      autoFill = [];
      break;
    case TAB.PROPERTIES:
      autoFill = [COLUMN.PROPERTIES.NAME_AUTOFILL];
      break;
    case TAB.CONTACTS:
      autoFill = [COLUMN.CONTACTS.FULL_NAME_AUTOFILL];
      break;
  }
  // Formatting:
  let range = sheet.getRange(2,1,1,lastColumn);
  range.copyFormatToRange(sheet,1,lastColumn,newRow,newRow);
  // Autofill:
  for (let col of autoFill) {
    let colIndex = getColumnIndex(sheetName, col);
    let range = sheet.getRange(2,colIndex,lastRow-1,1);
    let dest = sheet.getRange(2,colIndex,newRow-1,1);
    range.autoFill(dest, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  }
  // Display a msgbox notifying about successful execution
  let txt = "";
  for (let key in content) txt += `${key}: ${content[key]}\\n`;
  Browser.msgBox(
    `The record was successfully added to the ${sheetName} tab.\\n\\n${txt}`
  );
}

function generateUniqueId() {
  let existingIds = getColumnContent(TAB.CONTRACTS, COLUMN.CONTRACTS.ID);
  let characters =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  let newId;
  do {
    newId = "";
    for (let i = 0; i < 5; i++)
      newId += characters.charAt(Math.floor(Math.random() * characters.length));
  } while (existingIds.includes(newId));
  return newId;
}

function getColumnContent(sheetName, columnName) {
  let sheet = ss.getSheetByName(sheetName);
  let lastColumn = sheet.getLastColumn();
  let lastRow = sheet.getLastRow();
  let titles = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  let columnIndex = titles.indexOf(columnName) + 1;
  let range = sheet.getRange(2, columnIndex, lastRow - 1, 1);
  let value = range.getValues();
  let arr = [];
  for (let element of value) arr.push(element[0]);
  return arr;
}

function searchRow(sheetName, value) {
  if (!value) return null;
  let sheet = ss.getSheetByName(sheetName);
  let searchInColumn;
  switch (sheetName) {
    case TAB.CONTRACTS:
      searchInColumn = COLUMN.CONTRACTS.REFERENCE;
      break;
    case TAB.PAYMENTS:
      searchInColumn = COLUMN.PAYMENTS.REFERENCE;
      break;
    case TAB.PROPERTIES:
      searchInColumn = COLUMN.PROPERTIES.NAME_AUTOFILL;
      break;
    case TAB.CONTACTS:
      searchInColumn = COLUMN.CONTACTS.FULL_NAME;
      break;
  }
  // Searching within the column range
  let colIndex = getColumnIndex(sheetName, searchInColumn);
  let lastRow = sheet.getLastRow();
  let range = sheet.getRange(2,colIndex,lastRow-1,1)
  let colContent = range.getValues();
  let i;
  for (i = 0; i < colContent.length; i++)
    if (colContent[i][0] == value) break;
  if (i >= colContent.length) return -1;
  return i;
}

function getRowObject(sheetName, rowIndex) {
  let sheet = ss.getSheetByName(sheetName);
  let lastColumn = sheet.getLastColumn();
  let titles = sheet.getRange(1,1,1,lastColumn).getValues()[0];
  let row = sheet.getRange(rowIndex+2,1,1,lastColumn).getValues()[0];
  let obj = {};
  for (let k in titles) {
    let key = titles[k];
    let value = row[k];
    obj[key] = value;
  }
  return obj;
}

function updateRowContent(sheetName, rowIndex, content) {
  let sheet = ss.getSheetByName(sheetName);
  let lastColumn = sheet.getLastColumn();
  let titles = sheet.getRange(1,1,1,lastColumn).getValues()[0];
  let row = sheet.getRange(rowIndex+2,1,1,lastColumn).getValues()[0];
  let arr = [];
  for (let k in titles) {
    let key = titles[k];
    let value = row[k];
    if (Object.keys(content).includes(key)) {
      // Browser.msgBox (`ROW: ${rowIndex+2}, COLUMN: ${k+1}`)
      let r = sheet.getRange(rowIndex+2,parseInt(k)+1);
      r.setValue(content[key]);
    }
  }
  return true;
}

function uploadImgToDrive(obj) {
  const blob = Utilities.newBlob(
    Utilities.base64Decode(obj.data), 
    obj.mimeType, 
    obj.fileName
  );
  let file = DriveApp.createFile(blob);
  let id = file.getId();
  let folder = DriveApp.getFolderById("1aamye4GREcKd9MJY6F4as5wjDIausYNp");
  file.moveTo(folder);
  return id;
}

function getTabNames()    { return TAB; }
function getColumnNames() { return COLUMN; }
