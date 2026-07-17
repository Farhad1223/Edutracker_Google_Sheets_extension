/**
 * @OnlyCurrentDoc
 * This is the master onOpen function that runs when the spreadsheet is opened.
 * It initializes the menus for all extensions in this project.
 */
function onOpen() {
  // Call the function to create the menu for the Cell Notes extension.
  createCellNotesMenu(); 
  
  // Call the function to create the menu for your old extension.
  createOldExtensionMenu(); 
}