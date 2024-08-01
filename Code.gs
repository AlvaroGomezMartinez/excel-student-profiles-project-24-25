/*
The purpose of the function below called processStudentImages is to import
the hyperlinks of student pictures that are stored in a drive folder.

After processStudentImages is run, the Autocrat Chrome Extension runs and
merges the data into a google docs template.

When setting up for the new school year do the following:
     1. Create a folder where the picture images will be saved
     2. Create a folder where the teachers will see the profile pages
     3. Update the folder ID of the imageFolder variable in the processStudentImages function below (line 45)
     4. Update the folder ID of the folder variable in the uploadJpgToDrive function below (line 126)
     This Screencastify video shows how to do this: https://watch.screencastify.com/v/RQ27kmPDe48vBCuG3aP7 

Point of contact: Alvaro Gomez, Special Campuses Academic Technology Coach, 210-363-1577
*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Get Picture Hyperlinks')
      .addItem('Import Picture Hyperlinks', 'processStudentImages')
      .addSeparator()
      .addItem('References', 'openReferencesDialog')
      .addToUi();
}

function openReferencesDialog() {
  var htmlOutput = HtmlService.createHtmlOutput('<a href="https://watch.screencastify.com/v/RQ27kmPDe48vBCuG3aP7" target="_blank">1. How to update the drive folders</a><br><br>' + '<a href="https://watch.screencastify.com/v/PaafWij3FQSfBhexDu8p" target="_blank">2. How to run Autocrat to make the student profiles</a>')
      .setWidth(400)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'References');
}

function processStudentImages() {
  // showStatusDialog('Getting images and adding the hyperlinks...');
  // Gets the current spreadsheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Gets all the data in columns B, D, and L (excluding the header)
  var studentNames = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  var entryDates = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues();
  var imageLinks = sheet.getRange(2,12, sheet.getLastRow()-1,1).getValues();

  // Gets the folder where the .bmp images are stored
  var imageFolder = DriveApp.getFolderById('1lBd67o8T7loqLnxRhXp9j5HHWdtarYrC'); // Replace the text in red with the picture folder ID
  
  // Loop through the student names, entry dates, and image links from the google sheet
  for (var i = 0; i < studentNames.length; i++) {
    var studentName = studentNames[i][0];
    var entryDate = entryDates[i][0];
    var existingImageLink = imageLinks[i][0];
    
    // Check if existingImageLink is null or empty
    if (existingImageLink == null || existingImageLink === "") {
      // Continue processing for this student

      // Formats the date as "mm-dd-yy" with leading zeros
      var formattedDate = formatDateWithLeadingZeros(entryDate);
      
      // Concatenates student name and formatted entry date with a space in between. The purpose of this is to match the
      // name of the student in the sheet with the name of the .bmp file in the picture folder.
      var bmpFileName = studentName + ' ' + formattedDate + '.bmp';
      
      // Search for the .bmp file in the picture folder
      var bmpFile = findFileInFolder(imageFolder, bmpFileName);
      
      if (!bmpFile) {
        // If not found, it will try searching without leading zeros in the date
        formattedDate = formatDateWithoutLeadingZeros(entryDate);
        bmpFileName = studentName + ' ' + formattedDate + '.bmp';
        bmpFile = findFileInFolder(imageFolder, bmpFileName);
      }
      
      if (bmpFile) {
        // Converts .bmp to .jpg so the file can work with the google template
        var jpgBlob = convertBmpToJpg(bmpFile);
        
        if (jpgBlob) {
          // Get the URL of the converted .jpg
          var jpgUrl = uploadJpgToDrive(jpgBlob, studentName);
          
          if (jpgUrl) {
            // Adds the URL to column L in the spreadsheet
            sheet.getRange(i + 2, 12).setValue(jpgUrl);
            
            // Logs a message for successful processing
            Logger.log('Processed: ' + studentName);
          } else {
            // Logs an error message for URL retrieval failure
            Logger.log('Error getting URL for: ' + studentName);
          }
        } else {
          // Logs an error message for file conversion failure
          Logger.log('Error converting file for: ' + studentName);
        }
      } else {
        // Logs an error message for files not found
        Logger.log('File not found for: ' + studentName + ' (' + bmpFileName + ')');
      }
    } else {
      // Skip processing for this student and move to the next one
      Logger.log('Skipping processing for: ' + studentName + ' (Existing Image Link: ' + existingImageLink + ')');
      continue;
    }
  } // The for loop
}

// Below are five helper functions to help the processStudentImages function above
function formatDateWithLeadingZeros(date) {
  var month = ('0' + (date.getMonth() + 1)).slice(-2); // Add leading zero to month
  var day = ('0' + date.getDate()).slice(-2); // Add leading zero to day
  var year = date.getYear() - 100; // Get the last two digits of the year
  return month + '-' + day + '-' + year;
}

function formatDateWithoutLeadingZeros(date) {
  var month = date.getMonth() + 1; // Get month without leading zero
  var day = date.getDate(); // Get day without leading zero
  var year = date.getYear() - 100; // Get the last two digits of the year
  return month + '-' + day + '-' + year;
}

function findFileInFolder(folder, fileName) {
  var files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    return files.next();
  }
  return null;
}

function convertBmpToJpg(bmpFile) {
  try {
    // Convert BMP to JPG
    var jpgBlob = bmpFile.getBlob().setName(bmpFile.getName().replace(/\.bmp$/, '.jpg'));
    return jpgBlob;
  } catch (e) {
    return null; // Return null on conversion error
  }
}

function uploadJpgToDrive(jpgBlob, studentName) {
  try {
    var folder = DriveApp.getFolderById('1aHoho-tcELww5pQtJgM7YBEzJaZZgRZ6'); // Replace with your destination folder ID
    var jpgFile = folder.createFile(jpgBlob);
    return jpgFile.getUrl();
  } catch (e) {
    return null; // Return null on upload error
  }
}