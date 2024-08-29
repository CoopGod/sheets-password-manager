// Created by: Cooper Goddard
// Date: 2024-08-28
// This program helps to maintain a logbook of sites, usernames and 
// passwords in a google sheet
const initSiteCol = 'D';
const initSiteRow = 3;
const initLinksCol = 'B';
const initLinksRow = 2;

// AssignLinks manages a set of the alphabet that has links to the sites
// the begin with that particular letter of the alphabet for ease of 
// access when seaching for the site
function AssignLinks() {
  // Get sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];

  // Gather links to first letter of each site
  var linkLocations = createLinks(sheet,initSiteCol, initSiteRow);

  // Apply links connecting to the first letter of each site
  applyLinks(sheet, linkLocations);
}


// createLinks moves through the sites searching for each one with a new
// beginning letter of the alphabet to make a link
function createLinks(sheet, initSiteCol, initSiteRow) {
  const sheetId = sheet.getSheetId();
  const spreadsheetUrl = sheet.getParent().getUrl();
  var linkLocations = {};
  var currentLetter = '';
  // Get range of sites
  var siteRange = sheet.getRange(initSiteCol + initSiteRow.toString());
  var topSiteCell = siteRange.activateAsCurrentCell();
  var siteRange = sheet.getRange(`${topSiteCell.getA1Notation()}:${topSiteCell.getNextDataCell(SpreadsheetApp.Direction.DOWN).getA1Notation()}`);
  var sites = siteRange.getValues();
  for (let i = 0; i < sites.length; i++) {
    // check to if first letter is new, add to dict if so
    var firstLetter = sites[i][0][0].toUpperCase();
    if (currentLetter != firstLetter) {
      currentLetter = firstLetter;
      currentRow = initSiteCol + (initSiteRow + i).toString();
      linkLocations[firstLetter] = `${spreadsheetUrl}#gid=${sheetId}&range=${currentRow}`;
    }
  }

  return linkLocations;
}


// applyLinks uses a dictionary of linkLocations and assigns them
// to each letter of the alphabet accordingly
function applyLinks(sheet, linkLocations) {
    Object.keys(linkLocations).forEach(function(key, index) {
    Logger.log(key);
    var currentRange = sheet.getRange(initLinksCol + (initLinksRow + (key.charCodeAt(0) - 65)).toString());
    var currentCell = currentRange.activateAsCurrentCell()
    
    // build rich text with new url
    var richValue = SpreadsheetApp.newRichTextValue()
      .setText(key)
      .setLinkUrl(linkLocations[key])
      .build();
    Logger.log(richValue);
    currentCell.setRichTextValue(richValue);
  });
}


function addMenu() {
  var menu = SpreadsheetApp.getUi().createMenu('Link Assigner');
  menu.addItem('Assign Links', 'AssignLinks');
  menu.addToUi();
}


function onOpen(e) {
  addMenu();
}