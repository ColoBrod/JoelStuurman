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
    KAM:                  "KAM *Data*",
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
const addContract = () => 
  openSidebar("Add a new contract", "Contracts");
const addPayment = () =>  
  openSidebar("Add a new payment", "Payments");
const addProperty = () => 
  openSidebar("Add a new property", "Properties");
const addContact = () =>  
  openSidebar("Add a new contact", "Contacts");

// Testing dialog
const testDialog = () => 
  openDialog("Contacts");

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
  let dialog = new Dialog("ModalWindow.html", dialogTitle, sheetName);
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
    case "Contracts":
      autoFill = 
        [COLUMN.CONTRACTS.REFERENCE, COLUMN.CONTRACTS.CONTRACT_DURATION];
      break;
    case "Payments":
      autoFill = [];
      break;
    case "Properties":
      autoFill = [COLUMN.PROPERTIES.NAME_AUTOFILL];
      break;
    case "Contacts":
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

function getTabNames()    { return TAB; }
function getColumnNames() { return COLUMN; }
