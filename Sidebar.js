class Sidebar {
  constructor(src, sidebarTitle, headTitle, clearHistory) {
    this.src = src;
    this.headTitle = headTitle;
    let template = HtmlService.createTemplateFromFile(this.src);
    template.headTitle = headTitle;
    template.form = this.generateFormHtml();
    template.clearHistory = clearHistory;
    this.html = template.evaluate();
    this.title = sidebarTitle;
  }

  show() {
    this.html.setTitle(this.title);
    ui.showSidebar(this.html);
  }

  updateTitle(newTitle) {
    this.title = newTitle;
    this.html.setTitle(newTitle);
  }

  generateFormHtml() {
    let template = HtmlService
      .createTemplateFromFile(`form-${this.headTitle.toLowerCase()}.html`);
    Object.assign(template, formHeader[this.headTitle]);
    return template.evaluate().getContent();
  }

}
