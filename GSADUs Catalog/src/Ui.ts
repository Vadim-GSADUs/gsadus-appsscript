function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GSADUs')
    .addItem('Refresh (Master)', 'GS.Update.runAll')
    .addSeparator()
    .addItem('Refresh Image Registry', 'GS.Registry.refresh')
    .addSubMenu(
      ui.createMenu('Publish')
        .addItem('Publish to Production', 'GS.Publish.publishCatalog')
    )
    .addToUi();
}
