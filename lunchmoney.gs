// Lunchmoney Google Apps Script
//Copyright (c) 2023 - Robert McLellan

//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:

//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.

//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Lunchmoney');
  
  createConfigSheetIfNotExists();
  menu.addItem('Refresh All', 'refreshAllEndpoints');
  menu.addSeparator();
  menu.addItem('Refresh Categories', 'refreshCategories');
  menu.addItem('Refresh Transactions', 'refreshTransactions');
  menu.addItem('Refresh Budgets', 'refreshBudgets');
  menu.addItem('Refresh Tags', 'refreshTags');
  menu.addItem('Refresh Plaid Accounts', 'refreshPlaidAccounts');
  menu.addItem('Refresh Assets', 'refreshAssets');
  menu.addItem('Refresh Crypto', 'refreshCrypto');
  menu.addItem('Refresh Recurring Expenses', 'refreshRecurringExpenses');
  menu.addToUi();
  
  hideOtherSheets(); // Add this line to call the hideOtherSheets function
}

// Add this new function to hide other sheets
function hideOtherSheets() {
  const visibleSheets = ['transactions', 'budget', 'configuration'];
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  sheets.forEach(sheet => {
    const sheetName = sheet.getName().toLowerCase();
    if (!visibleSheets.includes(sheetName)) {
      sheet.hideSheet();
    }
  });
}


// Base URL for Lunch Money API
const BASE_URL = 'https://dev.lunchmoney.app/v1';
const CONFIG_SHEET_NAME = "Lunchmoney Configuration";
const API_KEY = 'abcdefghijklmnopqrstuvwxyz';
const TRANSACTION_START_DATE = '2022-01-01'
const TRANSACTION_END_DATE = '2023-01-01'
const BUDGETS_START_DATE = '2022-01-01'
const BUDGETS_END_DATE = '2023-01-01'
const RECURRING_START_DATE = '2023-01-01'

function createConfigSheetIfNotExists() {
  try {
    Logger.log('Attempting to create config sheet if it does not exist.');
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let configSheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);

    if (!configSheet) {
      configSheet = spreadsheet.insertSheet(CONFIG_SHEET_NAME);
      configSheet.appendRow(["Variable", "Value"]);
      configSheet.appendRow(["API_Key", API_KEY]);
      configSheet.appendRow(["TRANSACTION_START_DATE", TRANSACTION_START_DATE]);
      configSheet.appendRow(["TRANSACTION_END_DATE", TRANSACTION_END_DATE]);
      configSheet.appendRow(["BUDGETS_START_DATE", BUDGETS_START_DATE]);
      configSheet.appendRow(["BUDGETS_END_DATE", BUDGETS_END_DATE]);
      configSheet.appendRow(["RECURRING_START_DATE", RECURRING_START_DATE]);
      configSheet.autoResizeColumns(1, 2);
    }

    return configSheet;
  } catch (error) {
    Logger.log(`Error: Failed to create config sheet. ${error}`);
    return null;
  }
}

function getConfigValue(variable) {
  try {
    Logger.log(`Attempting to get config value for "${variable}".`);
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
    const lastRow = configSheet.getLastRow();

    for (let i = 2; i <= lastRow; i++) {
      const key = configSheet.getRange(i, 1).getValue();
      const value = configSheet.getRange(i, 2).getValue();

      if (key === variable) {
        Logger.log(`Found config value for "${variable}": ${value}`);
        return value;
      }
    }

    Logger.log(`Error: Variable "${variable}" not found in config sheet.`);
    return null;
  } catch (error) {
    Logger.log(`Error: Failed to get config value for "${variable}". ${error}`);
    return null;
  }
}

function createSheetIfNotExists(SHEET_NAME) {
  try {
    Logger.log(`Attempting to create sheet "${SHEET_NAME}" if it does not exist.`);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
    }

    return sheet;
  } catch (error) {
    Logger.log(`Error: Failed to create sheet "${SHEET_NAME}". ${error}`);
    return null;
  }
}

function toQueryParams(obj) {
  try {
    Logger.log(`Converting object to query string: ${JSON.stringify(obj)}`);
    return Object.keys(obj)
      .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(obj[key]))
      .join('&');
  } catch (error) {
    Logger.log(`Error: Failed to convert object to query string. ${error}`);
    return null;
  }
}

function htmlDecode(input) {
  if (typeof input !== 'string') {
    return input;
  }

  const entities = {
    '&amp;': '&',
    '&lt;': '<',
    '&gt;': '>',
    '&quot;': '"',
    '&#39;': "'",
    '&#x2F;': '/',
    '&#x27;': "'", // Add this line to handle &#x27;
  };

  return input.replace(/&[^;]+;/g, function (entity) {
    return entities[entity] || entity;
  });
}

function updateSheet(sheetName, fetchDataFunction, apiKey) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log(`Error: Sheet "${sheetName}" not found.`);
      return;
    }
  
    Logger.log(`Clearing the contents of "${sheetName}" sheet.`);
    sheet.clearContents();
    
    Logger.log(`Fetching data for "${sheetName}" sheet.`);
    const data = fetchDataFunction(apiKey);
    if (!data || data.length === 0) {
      Logger.log(`Error: No data found for "${sheetName}" sheet.`);
      return;
    }
    
    // Decode the HTML entities in the 'name' and 'description' fields
    data.forEach(tag => {
      Object.keys(tag).forEach(key => {
        if (tag[key] !== undefined) {
          tag[key] = htmlDecode(tag[key]);
        }
      });
    });
    
    Logger.log(`Updating the "${sheetName}" sheet with data.`);
    const headers = Object.keys(data[0]);
    sheet.appendRow(headers);
    const numRows = data.length;
    const numCols = headers.length;

    sheet.getRange(2, 1, numRows, numCols).setValues(data.map(row => headers.map(header => row[header])));

    Logger.log(`Successfully updated the "${sheetName}" sheet.`);
  } catch (error) {
    Logger.log(`Error: Failed to update sheet "${sheetName}". ${error}`);
  }
}

function apiRequest(endpoint, method, data, apiKey, queryParams = {}) {
  try {
    Logger.log(`Sending API request: ${method} ${endpoint}`);
    const options = {
      method: method,
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + apiKey
      },
      muteHttpExceptions: true // Add this line to avoid throwing exceptions for non-200 response codes
    };

    if (data) {
      options.payload = JSON.stringify(data);
    }

    const queryString = toQueryParams(queryParams);
    const url = BASE_URL + endpoint + (queryString ? '?' + queryString : '');
    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
      const errorDetails = JSON.parse(response.getContentText());
      throw new Error(`Error ${response.getResponseCode()}: ${errorDetails.error_message}`);
    }

    Logger.log(`API request successful: ${method} ${endpoint}`);
    const responseData = JSON.parse(response.getContentText());

    // Check if the response contains an array and extract it
    for (const key in responseData) {
      if (Array.isArray(responseData[key])) {
        return responseData[key];
      }
    }

    return responseData;
  } catch (error) {
    Logger.log(`Error: API request failed. ${error}`);
    return null;
  }
}

function getEndpoints() {
  return {
    '/categories': CATEGORIES,
    '/transactions': TRANSACTIONS,
    '/budgets': BUDGETS,
    '/tags': TAGS,
    '/plaid_accounts': PLAID_ACCOUNTS,
    '/assets': ASSETS,
    '/crypto': CRYPTO,
    '/recurring_expenses': RECURRING_EXPENSES,
  };
}

function tags(apiKey) {
  return apiRequest('/tags', 'GET', null, apiKey);
}

function categories(apiKey) {
  return apiRequest('/categories', 'GET', null, apiKey);
}

function transactions(apiKey) {
  const params = {
    start_date: TRANSACTION_START_DATE,
    end_date: TRANSACTION_END_DATE,
  };
  return apiRequest('/transactions', 'GET', null, apiKey, params);
}

function budgets(apiKey) {
  const params = {
    start_date: BUDGETS_START_DATE,
    end_date: BUDGETS_END_DATE,
  };
  return apiRequest('/budgets', 'GET', null, apiKey, params);
}

function plaidaccounts(apiKey) {
  return apiRequest('/plaid_accounts', 'GET', null, apiKey);
}

function crypto(apiKey) {
  return apiRequest('/crypto', 'GET', null, apiKey);
}

function recurringexpenses(apiKey) {
  const params = {
    start_date: RECURRING_START_DATE,
  };
  return apiRequest('/recurring_expenses', 'GET', null, apiKey, params);
}

function assets(apiKey) {
  return apiRequest('/assets', 'GET', null, apiKey);
}

function refreshCategories() {
  const sheetName = 'Lunchmoney /categories';
  createSheetIfNotExists(sheetName);
  updateSheet(sheetName, categories, getConfigValue('API_Key'));
}

function refreshTransactions() {
  const sheetName = 'Lunchmoney /transactions';
  createSheetIfNotExists(sheetName);
  updateSheet(sheetName, transactions, getConfigValue('API_Key'));
}

function refreshBudgets() {
  const sheetName = 'Lunchmoney /budgets';
  createSheetIfNotExists(sheetName);
  updateSheet(sheetName, budgets, getConfigValue('API_Key'));
}

function refreshTags() {
  const sheetName = 'Lunchmoney /tags';
  createSheetIfNotExists(sheetName);
  updateSheet(sheetName, tags, getConfigValue('API_Key'));
}

function refreshPlaidAccounts() {
  const sheetName = 'Lunchmoney /plaid_accounts';
  createSheetIfNotExists(sheetName);
  updateSheet(sheetName, plaidaccounts, getConfigValue('API_Key'));
}

function refreshAssets() {
  const sheetName = 'Lunchmoney /assets';
  createSheetIfNotExists(sheetName);
  updateSheet(sheetName, assets, getConfigValue('API_Key'));
}

function refreshCrypto() {
  const sheetName = 'Lunchmoney /crypto';
  createSheetIfNotExists(sheetName);
  updateSheet(sheetName, crypto, getConfigValue('API_Key'));
}

function refreshRecurringExpenses() {
  const sheetName = 'Lunchmoney /recurring_expenses';
  createSheetIfNotExists(sheetName);
  updateSheet(sheetName, recurringexpenses, getConfigValue('API_Key'));
}

function refreshAllEndpoints() {
  refreshCategories();
  refreshTransactions();
  refreshBudgets();
  refreshTags();
  refreshPlaidAccounts();
  refreshAssets();
  refreshCrypto();
  refreshRecurringExpenses();
}
