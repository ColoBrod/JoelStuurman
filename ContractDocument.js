class ContractDocument {
  
  static get templateId() { 
    return "1ZLw3azqFocvoRNzyF-B5aKGH58I_wx7OezuoiOR-8qI"; 
  }

  static get folderId() {
    return "1iIPqXxHQLQWZg9OlJRG7Jd8hT8ZfqBK0";
  }

  constructor(fileName) {
    this.fileName = fileName;
    this.folder = DriveApp.getFolderById(ContractDocument.folderId);
    const templateFile = DriveApp.getFileById(ContractDocument.templateId);
    const copiedFile = templateFile.makeCopy(fileName, this.folder);
    this.id = copiedFile.getId();
    this.file = DocumentApp.openById(this.id);
    this.url = this.file.getUrl();
    this.body = this.file.getBody();
    this.paragraphs = this.body.getParagraphs();
    this.p = {};
    for (let i = 0; i < this.paragraphs.length; i++) {
      let p = this.paragraphs[i];
      let txt = p.getText();
      if (txt.match(/Condo$/)) this.p.condo = p;
      else if (txt.match(/House$/)) this.p.house = p;
      else if (txt.match(/\(Option A\)/)) this.p.optionA = p;
      else if (txt.match(/\(Option B\)/)) this.p.optionB = p;
      else if (txt.match("Should the LESSEE fail or default "))
        this.p.underTable = p;
      else if (txt.match(/^ASSOCIATION DUES: /)) this.p.associationDues = p;
      else if (txt.match(/^POLICY ON PETS: /)) this.p.policyOnPets = p;
      else if (txt.match(/^ANNEX A: Furnishings/)) this.p.inventoryList = p;
    }
    this.p.condoP = this.p.condo.getNextSibling();
    this.p.houseP = this.p.house.getNextSibling();
    this.p.optionAP = this.p.optionA.getNextSibling();
    this.p.optionBP = this.p.optionB.getNextSibling();
    this.table = this.body.getTables();
    this.paymentTable = this.table[0];

    // this.body = this.file.getBody();
    // const template = DocumentApp.openById(ContractDocument.templateId);
    // const templateBody = template.getBody();
    // Browser.msgBox(ContractDocument.templateId);
  }

  build(
    lessorFullName,
    lessorNationality,
    lessorAddress,
    lesseeFullName,
    lesseeNationality,
    lesseeAddress,
    propertyNo,
    propertySection,
    propertyStreet,
    propertyAddress,
    propertyType,
    leaseStart,
    leaseEnd,
    contractDuration,
    inventoryList,
    leaseRate,
    parkingSlot,
    advanceApplicableOn,
    deposit,
    petClause,
    associationDues,
    remainingPayment
  ) {
    // Lessor data
    if (lessorFullName)
      this.body.replaceText("{{LessorFullName}}", lessorFullName);
    if (lessorNationality)
      this.body.replaceText("{{LessorNationality}}", lessorNationality);
    if (lessorAddress)
      this.body.replaceText("{{LessorAddress}}", lessorAddress);
    // Lessee data
    if (lesseeFullName)
      this.body.replaceText("{{LesseeFullName}}", lesseeFullName);
    if (lesseeNationality)
      this.body.replaceText("{{LesseeNationality}}", lesseeNationality);
    if (lesseeAddress)
      this.body.replaceText("{{LesseeAddress}}", lesseeAddress);
    // Property data
    if (propertyNo)
      this.body.replaceText("{{PropertyNo}}", propertyNo);
    if (propertySection)
      this.body.replaceText("{{PropertySection}}", propertySection);
    if (propertyStreet)
      this.body.replaceText("{{PropertyStreet}}", propertyStreet);
    if (propertyAddress)
      this.body.replaceText("{{PropertyAddress}}", propertyAddress);
    // CONDOMONIUM OR TOWNHOUSE
    // Remove B. House and sibling paragraph. Also remove A. Condo item list
    // Inventory list also affects this section
    this.p.condo.removeFromParent();
    this.p.house.removeFromParent();
    if (propertyType == "Condominium")
      this.p.houseP.removeFromParent();
    else if (propertyType == "Townhouse")
      this.p.condoP.removeFromParent();
    if (!inventoryList) {
      let txt = "inclusive of the furniture, appliances, equipment and fixtures found therein and listed on the Annex A of this contract \\(collectively, the “Furnishings”\\), "
      this.body.replaceText(txt, "");
      this.p.inventoryList.removeFromParent();
      // this.removePages(1);
    }
    
    let advanceArr = advanceApplicableOn.trim().split(/\s*,\s*/);
    for (let i in advanceArr) advanceArr[i] = parseInt(advanceArr[i]);
    let notPaid = [];
    for (let i = 1; i <= contractDuration; i++)
      if (advanceArr.indexOf(i) === -1) notPaid.push(i);
    let advancePeriod = "";
    for (let i = 0; i < advanceArr.length; i++) {
      let x = advanceArr[i];
      let delta = x-1;
      let startDate = new Date(leaseStart);
      if (delta != 0)
        startDate.setMonth(startDate.getMonth() + delta);
      let nextMonth = new Date(startDate);
      nextMonth.setMonth(nextMonth.getMonth() + 1);
      let endDate = new Date(nextMonth);
      endDate.setDate(endDate.getDate() - 1);
      let startDateFormatted = formatDate(startDate, false);
      let endDateFormatted = formatDate(endDate, false);
      let rentalPeriod = `${startDateFormatted} to ${endDateFormatted}`;
      if (i != 0) advancePeriod += " and ";
      advancePeriod += rentalPeriod;
    }

    let fillPaymentTable = () => {
      for (let i = 0; i < notPaid.length; i++) {
        let row = this.paymentTable.appendTableRow();
        let x = notPaid[i];
        let delta = x-1;
        let startDate = new Date(leaseStart);
        if (delta != 0)
          startDate.setMonth(startDate.getMonth() + delta);
        let nextMonth = new Date(startDate);
        nextMonth.setMonth(nextMonth.getMonth() + 1);
        let endDate = new Date(nextMonth);
        endDate.setDate(endDate.getDate() - 1);
        // Filling table:
        let startDateFormatted = formatDate(startDate, true);
        let endDateFormatted = formatDate(endDate, true);
        let rentalPeriod = `${startDateFormatted} - ${endDateFormatted}`;
        let amount = `PHP${numberWithCommas(leaseRate)}`;
        row.appendTableCell(startDateFormatted);
        row.appendTableCell(amount);
        row.appendTableCell(rentalPeriod);
      }
    }
    // Remaining Payment affects (Option A) and (Option B)
    this.p.optionA.removeFromParent();
    this.p.optionB.removeFromParent();
    switch (remainingPayment) {
      case "by deposit":
        this.p.optionAP.removeFromParent();
        fillPaymentTable();
        break;
      case "by check":
        this.p.optionBP.removeFromParent();
        this.paymentTable.removeFromParent();
        this.p.underTable.removeFromParent();
        break;
      default:
        this.p.optionB.removeFromParent();
        this.p.optionBP.removeFromParent();
        this.paymentTable.removeFromParent();
        this.p.underTable.removeFromParent();
    }

    // Additional data
    if (leaseStart) {
      leaseStart = formatDate(leaseStart);
      this.body.replaceText("{{LeaseStart}}", leaseStart);
    }
    if (leaseEnd) {
      leaseEnd = formatDate(leaseEnd);
      this.body.replaceText("{{LeaseEnd}}", leaseEnd);
    }
    if (contractDuration) {
      contractDuration = `${humanize(contractDuration)} (${contractDuration})`;
      this.body.replaceText("{{ContractDuration}}", contractDuration);
    }
    if (leaseRate) {
      let txt = numberWithCommas(leaseRate);
      this.body.replaceText("{{LeaseRate}}", txt);
      let advance = advanceArr.length
      if (advance) {
        txt = `${humanize(leaseRate*advance).toUpperCase()} (PHP${numberWithCommas(leaseRate*advance)})`;
        this.body.replaceText("{{LeaseRate\\*Advance}}", txt);
        txt = `${humanize(advance)} (${advance})`;
        this.body.replaceText("{{Advance}}", txt);
        this.body.replaceText("{{AdvancePeriod}}", advancePeriod);
      }
    }
    if (associationDues) {
      let txt = associationDues.toLowerCase();
      this.body.replaceText("{{AssociationDues}}", txt);
      if (associationDues === "Exclusive")
        this.p.associationDues.replaceText("LESSOR", "LESSEE");
    }
    console.log("Parking:");
    console.log(parkingSlot);
    if (parkingSlot) {
      this.body.replaceText("{{ParkingSlot}}", parkingSlot);
    } 
    else {
      let regex = ", inclusive rental for the parking slot Number \\{\\{ParkingSlot\\}\\}\\.";
      this.body.replaceText(regex, ".");
    }

    // if (advance) {
    //   let txt = `${humanize(advance)} (${advance})`;
    //   this.body.replaceText("{{Advance}}", txt);
    // }

    if (deposit) {
      let txt = `${humanize(deposit)} (${deposit})`
      this.body.replaceText("{{Deposit}}", txt);
      txt = `${humanize(leaseRate*deposit).toUpperCase()} (PHP${numberWithCommas(leaseRate*deposit)})`;
      this.body.replaceText("{{LeaseRate\\*Deposit}}", txt);
    }

    if (!petClause) this.p.policyOnPets.removeFromParent();
    
    // Count number of pages in contract

    return this.url;

  }

  removePages(n) {
    let paragraph = this.paragraphs
    let counter = 0;
    for (let i = paragraph.length - 1; i >= 0; i--) {
      for (var j = 0; j < paragraph[i].getNumChildren(); j++) 
        if (paragraph[i].getChild(j).getType() == DocumentApp.ElementType.PAGE_BREAK) counter++;
      paragraph[i].clear();
      if (counter == n) break;
    }
  }

  removeParagraph(regex) {

  }

}