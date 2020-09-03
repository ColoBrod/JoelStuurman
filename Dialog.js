class Dialog {
  constructor(src, dialogTitle, headTitle) {
    this.src = src;
    this.headTitle = headTitle;
    let template = HtmlService.createTemplateFromFile(this.src);
    template.headTitle = headTitle;
    template.form = this.generateFormHtml();
    this.html = template.evaluate();
    this.html.setWidth(800).setHeight(400);
    this.title = dialogTitle;
  }

  show() {
    ui.showModalDialog(this.html, this.title);
  }

  updateTitle(newTitle) {
    this.title = newTitle;
    this.html.setTitle(newTitle);
  }

  generateFormHtml() {
    let template = HtmlService
      .createTemplateFromFile(`dialog-${this.headTitle.toLowerCase()}.html`);
    switch (this.headTitle) {
      case "Contracts":
        Object.assign(template, { 
          kam:                COLUMN.CONTRACTS.KAM,
          transactionType:    COLUMN.CONTRACTS.TRANSACTION_TYPE,
          contractStartDate:  COLUMN.CONTRACTS.CONTRACT_START_DATE,
          contractEndDate:    COLUMN.CONTRACTS.CONTRACT_END_DATE,
          contractDuration:   COLUMN.CONTRACTS.CONTRACT_DURATION,
          landlord:           COLUMN.CONTRACTS.LANDLORD,
          grossSalesRate:     COLUMN.CONTRACTS.GROSS_SALES_RATE,
          contractStatus:     COLUMN.CONTRACTS.CONTRACT_STATUS,
          property:           COLUMN.CONTRACTS.PROPERTY,
          client:             COLUMN.CONTRACTS.CLIENT,
          propertyContact:    COLUMN.CONTRACTS.PROPERTY_CONTACT,
          contractUrl:        COLUMN.CONTRACTS.CONTRACT_URL,
        });
        break;
      case "Payments":
        Object.assign(template, { 
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
        });
        break;
      case "Properties":
        Object.assign(template, { 
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
        });
        break;
      case "Contacts":
        Object.assign(template, { 
          fullName:           COLUMN.CONTACTS.FULL_NAME,
          kam:                COLUMN.CONTACTS.KAM,
          contactType:        COLUMN.CONTACTS.CONTACT_TYPE,
          nationality:        COLUMN.CONTACTS.NATIONALITY,
          employer:           COLUMN.CONTACTS.EMPLOYER,
          address:            COLUMN.CONTACTS.ADDRESS,
          mobile:             COLUMN.CONTACTS.MOBILE,
          email:              COLUMN.CONTACTS.EMAIL,
          idLink:             COLUMN.CONTACTS.ID_LINK,
        });
        break;
    }
    return template.evaluate().getContent();
  }

}
