/**
 * @OnlyCurrentDoc
 */

const NOTES_SHEET_NAME = 'NotesData';
const IMAGE_FOLDER_NAME = 'Spreadsheet Cell Notes Images';

function createCellNotesMenu() {
  SpreadsheetApp.getUi()
      .createMenu('Cell Notes')
      .addItem('Show Note Dialog', 'showNoteDialogForActiveCell')
      .addToUi();
}

function showNoteDialogForActiveCell() {
  const activeCell = SpreadsheetApp.getActiveRange();
  if (activeCell.getNumRows() === 1 && activeCell.getNumColumns() === 1) {
    showNoteDialog(activeCell);
  } else {
    SpreadsheetApp.getUi().alert('Please select a single cell to manage its note.');
  }
}

// REPLACE this function
function showNoteDialog(range) {
  const a1Notation = range.getA1Notation();
  const sheetName = range.getSheet().getSheetName();
  
  // This is the new, unique key we will use for saving and loading
  const uniqueKey = sheetName + '!' + a1Notation;

  const htmlTemplate = HtmlService.createTemplateFromFile('dialog');
  
  // Pass BOTH the uniqueKey (for saving) and a1Notation (for display)
  htmlTemplate.uniqueKey = uniqueKey;
  htmlTemplate.a1Notation = a1Notation; // For the dialog title
  
  const htmlOutput = htmlTemplate.evaluate().setWidth(550).setHeight(650);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, `Note for ${sheetName}!${a1Notation}`);
}

// REPLACE this function
function saveNote(uniqueKey, noteContent) {
  try {
    // We now use the uniqueKey (e.g., "Sheet1!C5")
    PropertiesService.getUserProperties().setProperty(uniqueKey, noteContent);
  } catch (e) {
    Logger.log(`Error saving note for ${uniqueKey}: ${e.message}`);
  }
}

// REPLACE this function
function getNote(uniqueKey) {
  try {
    // We now use the uniqueKey (e.g., "Sheet1!C5")
    return PropertiesService.getUserProperties().getProperty(uniqueKey) || '';
  } catch (e) {
    Logger.log(`Error getting note for ${uniqueKey}: ${e.message}`);
    return ''; // Return empty string on error
  }
}

function getImagesFolder() {
  const folders = DriveApp.getFoldersByName(IMAGE_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(IMAGE_FOLDER_NAME);
}

/**
       * NEW FUNCTION
       * Intercepts paste events, checks for images, and uploads them
       * using the same logic as the 'Upload' button.
       */
      function handlePaste(e) {
        // Get the clipboard items
        const items = (e.clipboardData || e.originalEvent.clipboardData).items;
        let foundImage = false;

        for (let i = 0; i < items.length; i++) {
          const item = items[i];
          if (item.type.indexOf('image') !== -1) {
            // We found an image!
            foundImage = true;
            e.preventDefault(); // IMPORTANT: Stop the browser's default paste

            const file = item.getAsFile(); // Get the image as a file
            const reader = new FileReader();

            reader.onload = function(event) {
              // This is the *exact same logic* from your handleFileSelect function
              statusDiv.textContent = "Uploading pasted image...";
              google.script.run.withSuccessHandler(fileId => {
                if (fileId && !fileId.startsWith("Error")) {
                  // Insert the correct, Drive-backed image tag
                  const imgHtml = `<img src="" data-file-id="${fileId}" alt="Loading..." style="max-width: 95%;"/>`;
                  document.execCommand("insertHTML", false, imgHtml);
                  
                  loadImages(); // Load the image we just inserted
                  handleTyping(); // Trigger an autosave
                } else {
                  statusDiv.textContent = "Error: " + fileId; // Show the error
                }
              }).uploadImageAndGetId(event.target.result, file.type, file.name || 'pasted-image.png');
            };

            reader.readAsDataURL(file); // Read the file to trigger the 'onload'
            break; // Stop looping once we've handled the image
          }
        }
        // If no image was found (foundImage === false),
        // we don't call e.preventDefault(), so the browser 
        // pastes the text normally.
      }
      
/**
 * Uploads an image and returns its Google Drive File ID.
 */
function uploadImageAndGetId(base64Data, mimeType, fileName) {
  try {
    const folder = getImagesFolder();
    const data = Utilities.base64Decode(base64Data.split(',')[1]);
    const blob = Utilities.newBlob(data, mimeType, fileName || 'untitled.png');
    const file = folder.createFile(blob);
    // No need to set sharing or sleep anymore. We just need the ID.
    return file.getId();
  } catch (e) {
    return `Error: ${e.toString()}`;
  }
}

/**
 * **NEW FUNCTION**
 * Gets a file from Drive by its ID and returns it as a Base64 data string.
 * This is the key to reliable image rendering.
 * @param {string} fileId The ID of the file in Google Drive.
 * @return {string} The Base64 encoded data URI for the image.
 */
function getImageAsBase64(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const base64Data = Utilities.base64Encode(blob.getBytes());
    return `data:${blob.getContentType()};base64,${base64Data}`;
  } catch (e) {
    return `Error fetching image: ${e.toString()}`;
  }
}