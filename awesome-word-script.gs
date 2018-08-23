/**
Awesome* Word Script
(*maybe not great)
*/

var selectSpredsheet = SpreadsheetApp.getActive(); // Select active spredsheet

var indexSheet = 'Trenazer' // Select index sheet
var dataSheet = '01_Czlowiek' // Select data sheet
var cacheSheet = 'Pamięć podręczna' // Select cache sheet
var historySheet = 'Historia' // Select history sheet

// Index sheet
var selectSheetDestination = selectSpredsheet.getSheetByName(indexSheet);
// Data sheet
var selectSheet = selectSpredsheet.getSheetByName(dataSheet);
// Cache sheet
var selectSheetCache = selectSpredsheet.getSheetByName(cacheSheet);
// History sheet
var selectSheetHistory = selectSpredsheet.getSheetByName(historySheet);

function onEdit() {
  var getCellValue = selectSheetDestination.getRange('B2').getValue();
  var startLesson = selectSheetDestination.getRange('G7').getValue();
  
  // Get cell coordinates, for example 1, 1. It means A1
  var cell = selectSheetDestination.getActiveRange();
  var row = cell.getRow();
  var column = cell.getColumn();
  var cellCoordinates = column + ", " + row;
  
  if (cellCoordinates == '7, 7') {
    selectSheetDestination.getRange('G6').setValue('0');
    selectSheetHistory.getRange("A2:C").clearContent(); // Clear selected sheet range
    selectSheetCache.getRange("A2:C").clearContent(); // Clear selected sheet range
    
    var startLesson = selectSheetDestination.getRange('G7').getValue();
    
    if (startLesson != 0) {
      selectSheetCache.getRange("A2:C").clearContent(); // Clear cache before starting lesson
      selectSheetHistory.getRange("A2:C").clearContent(); // Clear histore before starting lesson
      startAwesomeLesson();
    }
  }
  
  // Check if word was written correctly
  if (getCellValue === 'Yup') {
    // Check if Awesome Lesson isn't active
    if (startLesson === 0) {
      checkCache();
      getAwesomeWord();
      selectSheetDestination.getRange('B8').clearContent();
    } else {
      startAwesomeLesson();
      getAwesomeWord();
      selectSheetDestination.getRange('B8').clearContent();
    }
  } else {
    var addToHistory = selectSheetDestination.getRange('D1:F1').getValues();
    selectSheetHistory.getRange(selectSheetHistory.getLastRow()+1, 1,1,3).setValues(addToHistory);
  }
}

// Restart cache if all data were extracted or increment drawing range
function checkCache() {
  // Get data sheet rows max range
  var getSheetDataRowSpread = selectSheet.getRange('A2:A').getValues();
  var getSheetDataRowRange = getSheetDataRowSpread.filter(String).length;
  
  // Get cache sheet rows max range
  var getSheetRowSpread = selectSheetCache.getRange('A1:A').getValues();
  var getSheetRowRange = getSheetRowSpread.filter(String).length;
  var whereStartExtraction = selectSheetCache.getRange('E1').getValue();
  
  // Get history sheet rows max range
  var getSheetHistoryRowSpread = selectSheetHistory.getRange('A1:A').getValues();
  var getSheetHistoryRowRange = getSheetHistoryRowSpread.filter(String).length;
 
    // Restart cache if all data were extracted
  if (getSheetRowRange <= whereStartExtraction) {
    if (getSheetHistoryRowRange <= 1) {
      selectSheetCache.getRange("A2:C").clearContent(); // Clear selected sheet range\
       copyInto(dataSheet, cacheSheet, 'A2:C'); // Copy data from selected sheets
    } else {
      selectSheetCache.getRange("A2:C").clearContent(); // Clear selected sheet range
      copyInto(historySheet, cacheSheet, 'A2:C'); // Copy data from selected sheets
      selectSheetCache.getRange("E1").setValue(2); // Row where start extraction
      removeDuplicates(cacheSheet); // remove non-unique values from selected sheet
      selectSheetHistory.getRange("A2:C").clearContent(); // Clear selected sheet range
    }
    
    /*
    var getSheetRowSpread = selectSheetCache.getRange('A2:A').getValues();
    var getSheetRowRange = getSheetRowSpread.filter(String).length;
    var lastDrawedRange = selectSheetCache.getRange('E1').getValue();
    var drawingRange = lastDrawedRange;
    */
    
    selectSheetCache.getRange("E1").setValue(2); // Row where start extraction
    selectSheetDestination.getRange('B8').clearContent();
    shuffleSheet(cacheSheet);
  } else {
    incrementDrawingRange();
  }
}

function incrementDrawingRange() {
  var drawingRange = selectSheetCache.getRange('E1').getValue();
    drawingRange++;
    selectSheetCache.getRange('E1').setValue(drawingRange); // Cache/Overwrite a last generated number
}

function getAwesomeWord() {    
  var lastDrawedRange = selectSheetCache.getRange('E1').getValue();
  var drawingRange = lastDrawedRange;
  
  var setDrawedRange = 'A' + drawingRange + ':' + 'B' + drawingRange;
  
  // Break data down into individual cells
  var cellFirst = 'A' + drawingRange;
  var cellSecond = 'B' + drawingRange;
  
  // Assign the range you want to copy
  // Set range
  var cellFirstRange = selectSheetCache.getRange(cellFirst);
  var cellSecondRange = selectSheetCache.getRange(cellSecond);
  var fullRange = selectSheet.getRange(setDrawedRange);
  
  // Downlad data from set range 
  var fullData = fullRange.getValues();
  var cellFirstData = cellFirstRange.getValue();
  var cellSecondData = cellSecondRange.getValue();
  
  // Write/Overwrite random word to Index sheet
  selectSheetDestination.getRange('B4').setValue(cellFirstData);
  selectSheetDestination.getRange('B5').setValue(cellSecondData);
  
  // Write random word to History sheet, uncomment to use
  //selectSheetHistory.getRange(selectSheetHistory.getLastRow()+1, 1,1,2).setValues(fullData);
}

function startAwesomeLesson() {
  var numberOfWordsPerLesson = selectSheetDestination.getRange('G7').getValue(); // Amount of words to copy
  var numberOfLessons = selectSheetDestination.getRange('G5').getValue();
  var currentLesson = selectSheetDestination.getRange('G6').getValue();
  // Not allow divide by zero
  if (currentLesson === 0) {
    var fakeLesson = selectSheetDestination.getRange('G6').getValue() + 1;
  }
  
  var currentLessonWordRange = (numberOfWordsPerLesson * fakeLesson) - 1; // Count the falling range per lesson
  var currentLessonWordRangeStart = (currentLessonWordRange + 3) - numberOfWordsPerLesson; // Count the falling range per lesson
  var copyIntoLessonRange = 'A' + currentLessonWordRangeStart + ':C' + (currentLessonWordRange + 3); // Make a range of words to copy
  
  // Get cache sheet rows max range
  var getSheetRowSpread = selectSheetCache.getRange('A1:A').getValues();
  var getSheetRowRange = getSheetRowSpread.filter(String).length;
  var whereStartExtraction = selectSheetCache.getRange('E1').getValue();
  
  // Get history sheet rows max range
  var getSheetHistoryRowSpread = selectSheetHistory.getRange('A1:A').getValues();
  var getSheetHistoryRowRange = getSheetHistoryRowSpread.filter(String).length;
  
  // Restart cache if all data were extracted
  if (getSheetRowRange <= whereStartExtraction) {
    if (getSheetHistoryRowRange <= 1) {
      var incrementCurrentLesson = currentLesson + 1;
      
      selectSheetDestination.getRange('G6').setValue(incrementCurrentLesson); // Increment currentLesson
      
      var numberOfLessons = selectSheetDestination.getRange('G5').getValue();
      var currentLesson = selectSheetDestination.getRange('G6').getValue();
      
      if (numberOfLessons < currentLesson) {
        // End lessons
        selectSheetCache.getRange("A2:C").clearContent(); // Clear selected sheet range
        selectSheetDestination.getRange('G6').setValue(0);
        selectSheetDestination.getRange('G7').setValue(0);
        checkCache();
        getAwesomeWord();
      } else {
        var currentLessonWordRange = (numberOfWordsPerLesson * currentLesson) - 1; // Count the falling range per lesson
        var currentLessonWordRangeStart = (currentLessonWordRange + 3) - numberOfWordsPerLesson; // Count the falling range per lesson
        var copyIntoLessonRange = 'A' + currentLessonWordRangeStart + ':C' + (currentLessonWordRange + 3); // Make a range of words to copy
        
        selectSheetCache.getRange("A2:C").clearContent(); // Clear selected sheet range
        copyInto(dataSheet, cacheSheet, copyIntoLessonRange); // Copy data from selected sheets
      }
    } else {
      selectSheetCache.getRange("A2:C").clearContent(); // Clear selected sheet range
      copyInto(historySheet, cacheSheet, 'A2:C'); // Copy data from selected sheets
      selectSheetCache.getRange("E1").setValue(2); // Row where start extraction
      removeDuplicates(cacheSheet); // remove non-unique values from selected sheet
      selectSheetHistory.getRange("A2:C").clearContent(); // Clear selected sheet range
    }
    
    /*
    var getSheetRowSpread = selectSheetCache.getRange('A1:A').getValues();
    var getSheetRowRange = getSheetRowSpread.filter(String).length;
    var lastDrawedRange = selectSheetCache.getRange('E1').getValue();
    var drawingRange = lastDrawedRange;
    */
    
    selectSheetCache.getRange("E1").setValue(2); // Row where start extraction
    selectSheetDestination.getRange('B8').clearContent();
    shuffleSheet(cacheSheet);
    getAwesomeWord();
  } else {
    incrementDrawingRange();
  }
}

// Cmpare characters 
// Custom formula from https://mashe.hawksey.info/2013/07/google-spreadsheet-how-to-compare-two-strings-and-highlight-the-differences/
// This formula is licensed under a Creative Commons Attribution 3.0 Unported License. CC-BY mhawksey 
function stringComparison(s1, s2) {
  // lets test both variables are the same object type if not throw an error
  if (Object.prototype.toString.call(s1) !== Object.prototype.toString.call(s2)){
    throw("Both values need to be an array of cells or individual cells")
  }
  // if we are looking at two arrays of cells make sure the sizes match and only one column wide
  if( Object.prototype.toString.call(s1) === '[object Array]' ) {
    if (s1.length != s2.length || s1[0].length > 1 || s2[0].length > 1){
      throw("Arrays of cells need to be same size and 1 column wide");
    }
    // since we are working with an array intialise the return
    var out = [];
    for (r in s1){ // loop over the rows and find differences using diff sub function
      out.push([diff(s1[r][0], s2[r][0])]);
    }
    return out; // return response
  } else { // we are working with two cells so return diff
    return diff(s1, s2)
  }
}
function diff (s1, s2){
  var out = "[ ";
  var notid = false;
  // loop to match each character
  for (var n = 0; n < s1.length; n++){
    if (s1.charAt(n) == s2.charAt(n)){
      // I changed: "out += "–";" too "out += s2.charAt(n);"
      out += s2.charAt(n);
    } else {
      // I changed: "out += s2.charAt(n);" too "out += s2.charAt(n);"
      out += "–";
      notid = true;
    }
out += " ";
  }
  out += " ]"
  return (notid) ? out :  "[ id. ]"; // if notid(entical) return output or [id.]
}

// Make a data copy from specific sheet
function copyInto(sheetToCopy, sheetToPaste, rangeToCopy) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var copySheet = ss.getSheetByName(sheetToCopy);
  var pasteSheet = ss.getSheetByName(sheetToPaste);

  // get source range
  var source = copySheet.getRange(rangeToCopy);
  // get destination range
  var destination = pasteSheet.getRange(pasteSheet.getLastRow()+1, 1,1,2);

  // copy values to destination range
  source.copyTo(destination);
}

// Shuffle sheet
function shuffleSheet(shuffleSheetName) {
  var sheet = selectSpredsheet.getSheetByName(shuffleSheetName);
  var getFromSheetLastRow = sheet.getLastRow();
  var makeRange = 'A2' + ':' + 'C' + getFromSheetLastRow;
  var range = sheet.getRange(makeRange);
  range.setValues(shuffleArray(range.getValues()));    
}    

function shuffleArray(array) {
  var i, j, temp;
  for (i = array.length - 1; i > 0; i--) {
    j = Math.floor(Math.random() * (i + 1));
    temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
  return array;
}

// Remove only duplicates, non-unique values from selected sheet
function removeDuplicates(removeDuplicatesFromSheet) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(removeDuplicatesFromSheet);
  var data = sheet.getDataRange().getValues();
  var newData = new Array();
  for(i in data){
    var row = data[i];
    var duplicate = false;
    for(j in newData){
      if(row.join() == newData[j].join()){
        duplicate = true;
      }
    }
    if(!duplicate){
      newData.push(row);
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length)
      .setValues(newData);
}
