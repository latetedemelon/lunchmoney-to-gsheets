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
const API_KEY = '604d25abd1c770cdb9ddf8f8d8ddd2ef462649d1dab3074cc9';
const TRANSACTION_START_DATE = '2023-01-01'
const TRANSACTION_END_DATE = '2023-04-30'
const BUDGETS_START_DATE = '2023-01-01'
const BUDGETS_END_DATE = '2023-04-30'
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

function updateSheet(sheetName, fetchDataFunction, apiKey, startDate, endDate) {
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
    const data = fetchDataFunction(apiKey, startDate, endDate);
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

function transactions(apiKey, startDate, endDate) {
  const params = {
    start_date: startDate,
    end_date: endDate,
  };
  return apiRequest('/transactions', 'GET', null, apiKey, params);
}

function budgets(apiKey, startDate, endDate) {
  const params = {
    start_date: startDate,
    end_date: endDate,
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
  updateSheet(sheetName, transactions, getConfigValue('API_Key'), TRANSACTION_START_DATE, TRANSACTION_END_DATE);
}

function refreshBudgets() {
  const sheetName = 'Lunchmoney /budgets';
  createSheetIfNotExists(sheetName);
  updateSheet(sheetName, budgets, getConfigValue('API_Key'), BUDGETS_START_DATE, BUDGETS_END_DATE);
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




function getExpensesForCategory(transactions, categoryId, startDate, endDate) {
  let totalExpenses = 0;

  transactions.forEach(transaction => {
    const transactionDate = new Date(transaction.date);

    if (transaction.category_id === categoryId && transactionDate >= startDate && transactionDate <= endDate) {
      totalExpenses += transaction.to_base;
    }
  });

  return totalExpenses.toFixed(2);
}

function getBudgetForCategory(budgets, categoryId, startDate, endDate) {
  let totalBudget = 0;

  budgets.forEach(budget => {
    if (budget.category_id === categoryId) {
      const dataKeys = Object.keys(budget.data);
      dataKeys.forEach(date => {
        const budgetDate = new Date(date);

        if (budgetDate >= startDate && budgetDate <= endDate) {
          totalBudget += budget.data[date].budget_to_base;
        }
      });
    }
  });

  return totalBudget.toFixed(2);
}

function getCategoryNameById(categories, categoryId) {
  const category = categories.find(category => category.id === categoryId);

  return category ? category.name : '';
}

function compareExpensesWithBudget(transactions, categories, budgets, startDate, endDate) {
  const results = [];

  categories.forEach(category => {
    const categoryId = category.id;
    const categoryName = getCategoryNameById(categories, categoryId);
    const totalExpenses = getExpensesForCategory(transactions, categoryId, startDate, endDate);
    const totalBudget = getBudgetForCategory(budgets, categoryId, startDate, endDate);

    results.push({
      category_id: categoryId,
      category_name: categoryName,
      total_expenses: totalExpenses,
      total_budget: totalBudget,
      difference: (totalExpenses - totalBudget).toFixed(2)
    });
  });

  return results;
}


function isDateWithinRange(date, startDate, endDate) {
  return date >= startDate && date <= endDate;
}



async function fetchTransactions(apiKey, startDate, endDate) {
  return await transactions(apiKey, startDate, endDate);
}

async function fetchCategories(apiKey) {
  return await categories(apiKey);
}

async function fetchBudgets(apiKey, startDate, endDate) {
  return await budgets(apiKey, startDate, endDate);
}

async function fetchRecurringExpenses(apiKey, startDate, endDate) {
  return await recurringexpenses(apiKey, startDate, endDate);
}



async function getTotalIncomeAndExpenses(apiKey, startDate, endDate) {
  const transactions = await fetchTransactions(apiKey, startDate, endDate);
  const categories = await fetchCategories(apiKey);
  let totalIncome = 0;
  let totalExpenses = 0;

  // Convert startDate and endDate to Date objects
  startDate = new Date(startDate);
  endDate = new Date(endDate);

  // Helper function to find category by its ID
  function findCategoryById(categoryId) {
    return categories.find(category => category.id === categoryId);
  }

  // Helper function to compare dates without time components
  function isDateWithinRange(date, startDate, endDate) {
    const year = date.getFullYear();
    const month = date.getMonth();
    const day = date.getDate();

    const startYear = startDate.getFullYear();
    const startMonth = startDate.getMonth();
    const startDay = startDate.getDate();

    const endYear = endDate.getFullYear();
    const endMonth = endDate.getMonth();
    const endDay = endDate.getDate();

    return (
      (year > startYear || (year === startYear && (month > startMonth || (month === startMonth && day >= startDay)))) &&
      (year < endYear || (year === endYear && (month < endMonth || (month === endMonth && day <= endDay))))
    );
  }

  transactions.forEach(transaction => {
    const transactionDate = new Date(transaction.date);

    if (isDateWithinRange(transactionDate, startDate, endDate)) {
      const category = findCategoryById(transaction.category_id);

      if (category && !category.exclude_from_budget) {
        if (category.is_income) {
          totalIncome += transaction.to_base;
        } else {
          totalExpenses += transaction.to_base;
        }
      }
    }
  });

  return {
    total_income: totalIncome.toFixed(2),
    total_expenses: totalExpenses.toFixed(2),
  };
}





async function getTotalBudgetAndSpend(apiKey, startDate, endDate) {
  const budgets = await fetchBudgets(apiKey, startDate, endDate);
  const transactions = await fetchTransactions(apiKey, startDate, endDate);

  let totalBudget = 0;
  let totalSpend = 0;
  let totalTransactionSpend = 0;

  const excludedCategories = [239188, 127507, 239190, 126309, 211808, 210116, 127500, 138568, 127506];

  const monthsInRange = getMonthsInRange(startDate, endDate);
  console.log(monthsInRange);

  budgets.forEach(budget => {
    const budgetData = budget.data;

    if (budget.is_group !== true && !excludedCategories.includes(budget.category_id)) {
      for (const date in budgetData) {
        if (monthsInRange.includes(date)) {
          totalBudget += parseFloat(budgetData[date].budget_to_base || 0);
          totalSpend += parseFloat(budgetData[date].spending_to_base || 0);
        }
      }
    }
  });

  transactions.forEach(transaction => {
    const transactionDate = new Date(transaction.date);
    if (
      transactionDate >= startDate &&
      transactionDate <= endDate &&
      transaction.category_id &&
      transaction.is_group !== true &&
      !excludedCategories.includes(transaction.category_id)
    ) {
      totalTransactionSpend += transaction.to_base;
    }
  });

  return {
    total_budget: totalBudget.toFixed(2),
    total_spend: totalSpend.toFixed(2),
    total_transaction_spend: totalTransactionSpend.toFixed(2),
  };
}


async function getBudgetSpendAndTransactionsByCategory(apiKey, startDate, endDate) {
  const[startDateSom, startDateEom] = getStartAndEndDatesOfMonth(startDate);
  const[endDateSom, endDateEom] = getStartAndEndDatesOfMonth(endDate);

  const budgets = await fetchBudgets(apiKey, startDateSom, endDateEom);
  const transactions = await fetchTransactions(apiKey, startDateSom, endDateEom);
  const monthsInRange = getMonthsInRange(startDateSom, endDateEom);

  console.log(monthsInRange);

  const result = {};

  budgets.forEach(budget => {
    const budgetData = budget.data;

    if (budget.category_id === 126299 && budget.is_group !== true) {
      result[budget.category_name] = {
        transactions: {},
      };

      for (const date in budgetData) {
        if (monthsInRange.includes(date)) {
          if (!result[budget.category_name][date]) {
            result[budget.category_name][date] = {
              budget: 0,
              spend: 0,
              transactions: {},
            };
          }
          result[budget.category_name][date].budget += parseFloat(budgetData[date].budget_to_base);
          result[budget.category_name][date].spend += parseFloat(budgetData[date].spending_to_base);
        }
      }
    }
  });

  transactions.forEach(transaction => {
    const transactionDate = transaction.date;

    if (transaction.category_id === 126299 && transactionDate >= startDateSom && transactionDate <= endDateEom) {
      const categoryName = budgets.find(budget => budget.category_id === transaction.category_id).category_name;
      const payee = transaction.payee;

      // Get the transaction month (yyyy-mm format)
      const [year, month] = transactionDate.split("-");
      const transactionMonth = `${year}-${month}-01`;

      if (result[categoryName] && result[categoryName][transactionMonth]) {
        if (!result[categoryName][transactionMonth].transactions[payee]) {
          result[categoryName][transactionMonth].transactions[payee] = 0;
        }
        result[categoryName][transactionMonth].transactions[payee] += parseFloat(transaction.amount);
      }
    }
  });

  console.log('Category data:', JSON.stringify(result, null, 2));
  return result;
}



function getMonthsInRange(startDate, endDate) {
  const monthsInRange = [];
  console.log(startDate, endDate);
  const parseDate = (dateString) => {
    const [year, month, day] = dateString.split("-");
    return new Date(year, month - 1, day);
  };

  const startDateFom = parseDate(startDate);
  const endDateFom = parseDate(endDate);

  while (startDateFom <= endDateFom) {
    monthsInRange.push(
      `${startDateFom.getFullYear()}-${(startDateFom.getMonth() + 1)
        .toString()
        .padStart(2, "0")}-01`
    );

    if (startDateFom.getMonth() === 11) {
      startDateFom.setFullYear(startDateFom.getFullYear() + 1);
      startDateFom.setMonth(0);
    } else {
      startDateFom.setMonth(startDateFom.getMonth() + 1);
    }
  }

  return monthsInRange;
}


function getStartAndEndDatesOfMonth(dateString) {
  const parseDate = (dateString) => {
    const [year, month, day] = dateString.split("-");
    return new Date(year, month - 1, day);
  };

  const formatDate = (date) => {
    const year = date.getFullYear();
    const month = (date.getMonth() + 1).toString().padStart(2, "0");
    const day = date.getDate().toString().padStart(2, "0");
    return `${year}-${month}-${day}`;
  };

  const inputDate = parseDate(dateString);

  const startOfMonth = new Date(inputDate.getFullYear(), inputDate.getMonth(), 1);
  const endOfMonth = new Date(inputDate.getFullYear(), inputDate.getMonth() + 1, 0);

  return [formatDate(startOfMonth),formatDate(endOfMonth)];
}



async function testReport() {
  try {
    console.log("Starting test report...");
    const apiKey = getConfigValue('API_Key');
    const startDate = '2023-01-02';
    const endDate = '2023-04-30';

    // Get total income and expenses
    console.log("Getting total income and expenses...");
    const incomeAndExpenses = await getTotalIncomeAndExpenses(apiKey, startDate, endDate);
    console.log("Total income and expenses retrieved successfully:", incomeAndExpenses);

    // Get total budget and spend
    //console.log("Getting total budget and spend...");
    //const budgetAndSpend = await getTotalBudgetAndSpend(apiKey, startDate, endDate);
    //console.log("Total budget and spend retrieved successfully:", budgetAndSpend);

    // Get budget and spend by category
    console.log("Getting budget and spend by category...");
    const budgetAndSpendByCategory = await getBudgetSpendAndTransactionsByCategory(apiKey, startDate, endDate);
    console.log("Budget and spend by category retrieved successfully:", JSON.stringify(budgetAndSpendByCategory, null, 2));

    //const dateString = "2023-02-15";
    //const result = getStartAndEndDatesOfMonth(dateString);
    //console.log(result); // {start: "2023-02-01", end: "2023-02-28"}

    // Get planned and actual spend
    //console.log("Getting planned and actual spend...");
    // const plannedAndActualSpend = await getPlannedAndActualSpend(apiKey, startDate, endDate);
    //console.log("Planned and actual spend retrieved successfully:", plannedAndActualSpend);

  } catch (error) {
    console.error("An error occurred during test report execution:", error);
    console.log("Please check the logs for more information.");
  }
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
