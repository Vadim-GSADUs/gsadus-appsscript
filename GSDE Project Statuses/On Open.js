/**
 * GSADUs Smart Sync - Menu Configuration
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('GSADUs Smart Sync')
    // Discover & Sync
    .addItem('Discover & Sync All (Recommended)', 'discoverAndSyncManual')
    .addSeparator()
    .addItem('Discover New Projects (from Drive)', 'discoverNewProjectsManual')
    .addItem('Sync Statuses (from Statuses tab)', 'syncStatusesManual')
    .addSeparator()
    // Push (Sheet -> Drive)
    .addItem('Push Active Projects (Sheet -> Drive)', 'pushActiveProjects')
    .addItem('Push ALL Projects (Sheet -> Drive)', 'pushAllProjects')
    .addSeparator()
    // New GSDE Master Sync
    .addItem('Sync to GSDE (Master CSV)', 'mirrorFoldersToGSDE')
    .addSeparator()
    // Pull (Drive -> Sheet)
    .addItem('Pull Active Projects (Drive -> Sheet)', 'pullActiveProjects')
    .addItem('Pull ALL Projects (Drive -> Sheet)', 'pullAllProjects')
    .addSeparator()
    // Sync (Pull then Push)
    .addItem('Sync Active (Pull then Push)', 'syncActiveProjects')
    .addItem('Sync ALL (Pull then Push)', 'syncAllProjects')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Advanced / Repair')
        .addItem('Fix Phantom Links (Repair Missing File IDs)', 'repairPhantomLinks')
        .addItem('Batch Update Missing Addresses', 'fillMissingAddresses')
        .addItem('Reset Config Tab (Rebuild Headers)', 'initializeConfigTab')
    )
    .addToUi();
}