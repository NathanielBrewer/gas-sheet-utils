function getSubscriptions() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const subscriptions = scriptProperties.getProperty('SheetUtils.subscriptions');
  return JSON.parse(subscriptions) ?? [];
}

function checkRequirements () {
  if(!executeHandlerByName) {
    console.error(`[SheetUtils.checkRequirements()] missing requirements error. SheetUtils optionally requires the parent GS script to implement the executeHandlerByName function if using subscriptions. See https://github.com/NathanielBrewer?tab=repositories for more information`);
  };
  if(!SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Apps Script Config')) {
    throw new Error('[SheetUtils.checkRequirements()] error. No "Apps Script Config" sheet found. Add the sheet and run setup() again. See https://github.com/NathanielBrewer?tab=repositories for more information');
  }
}

function setup() {
  checkRequirements();
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteAllProperties();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Apps Script Config');
  
  const FIRST_SPREADSHEET_SETTING_ROW = 2;
  const FIRST_SHEET_SETTING_ROW = 5;
  const spreadsheetSettingsRange = sheet.getRange(FIRST_SPREADSHEET_SETTING_ROW, 1, 3, sheet.getLastColumn());
  const spreadsheetSettings = spreadsheetSettingsRange.getValues();

  const getProcessedSettings = (headers, values, prefix) => {
    const processedSettings = {};
    headers.forEach((header, index) => {
      processedSettings[`${prefix}.${header}`] = values[index];
    });
    return processedSettings;
  };

  const processedSpreadsheetSettings = getProcessedSettings(spreadsheetSettings[0], spreadsheetSettings[1], 'spreadsheet');
  scriptProperties.setProperties(processedSpreadsheetSettings);
  const processedSheetSettings = [];

  const sheetSettings = sheet.getRange(FIRST_SHEET_SETTING_ROW, 1, sheet.getLastRow() - (FIRST_SHEET_SETTING_ROW - 1), sheet.getLastColumn()).getValues();
  const sheetHeadings = sheetSettings.shift();
  const indexOfSheetName = sheetHeadings.findIndex((heading) => heading == 'sheetName');
  const sheetNames = [];
  for(let i = 0; i < sheetSettings.length; i++) {
    const sheetSettingsValues = sheetSettings[i];
    const sheetName = `${sheetSettingsValues[indexOfSheetName]}`;
    processedSheetSettings.push(getProcessedSettings(sheetHeadings, sheetSettingsValues, `${sheetName}`));
    sheetNames.push(sheetName);
  }
  processedSheetSettings.forEach((settings) => scriptProperties.setProperties(settings));
  scriptProperties.setProperty('spreadsheet.configuredSheetNames', JSON.stringify(sheetNames));
}

function addSubscription( subscription) {
  const subscriptions = getSubscriptions();
  subscriptions.push(subscription);
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('SheetUtils.subscriptions', JSON.stringify(subscriptions));
}

function onSheetEvent(event) {
  const trigger = GasToolkit.getTriggerById(event.triggerUid);
  const eventType = trigger.getEventType();
  const subscriptions = getSubscriptions();
  subscriptions.forEach((subscription) => {
    if(
      eventType == subscription.eventType
      && event.source.getSheetName() == subscription.sheetName
      && isEventInRange(event.range, subscription.range)
    ) {
      try {
        executeHandlerByName(subscription.handler, event);
      } catch(error) {
        console.error('[SheetUtils.onSheetEvent(event)] error. The error is probably because an `executeHandlerByName(functionName, event)` method has not been implemented in parent GS script.', error);
      }
    }
  });
}

function isEventInRange(eventRange, subscriptionRange) {
  let eventRow = eventRange.getRow();
  if (eventRow < subscriptionRange.top && eventRow > subscriptionRange.bottom) {
    return false;
  }
  let eventCol = eventRange.getColumn();
  if (eventCol < subscriptionRange.left || eventCol > subscriptionRange.right) {
    return false;
  }
  return true;
}

function getColumnsWithHeader(header, headerRow, sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const matchedHeaders = [];
  headers.forEach((h, index) => {
    if (header.toLowerCase() == h.toLowerCase()) {
      matchedHeaders.push(index + 1); // Convert zero-based index to one-based column number
    }
  });
  return matchedHeaders;
}

class Range {
  constructor(top, right, bottom, left) {
    this.top = top;
    this.right = right;
    this.bottom = bottom;
    this.left = left;
  }
}

class Subscription{
  constructor(sheetName, range, handler, eventType){
    this.sheetName = sheetName;
    this.range = range;
    this.handler = handler;
    this.eventType = eventType;
  }
}