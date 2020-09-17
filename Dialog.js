class Dialog {
  constructor(src, dialogTitle, headTitle, toSearch) {
    this.src = src;
    this.headTitle = headTitle;
    let template = HtmlService.createTemplateFromFile(this.src);
    template.headTitle = headTitle;
    template.form = this.generateFormHtml();
    template.toSearch = toSearch;
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
    Object.assign(template, formHeader[this.headTitle]);
    return template.evaluate().getContent();
  }

}
