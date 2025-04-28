const BASE_URL = 'https://chudoshkolaumnica.s20.online/v2api';
const BASE_URL_WITH_BRANCH = BASE_URL + '/1';
let authToken = '';
let leadStatuses = [];
let customerStatuses = [];
let leadRejectReasons = [];
// branches like 
// [{"id":2,"name":"Верхняя Первомайская 59\\35-2","subject_ids":[1,2,3,4,5,6,7,10,12,14,15,16,17,19,20,21,23,24,25,27,28,29,31,33,35,36,39,40,41,42,43,44,45,46,47,48,49,50,51,52,54,55,56,57,58,59,60,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101,102,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121,122,123,124,125,126,127,128,129,130,131,132,133,134,135,136,137,138,139,142],"is_active":1,"weight":1},{"id":5,"name":"Монтессори-парк","subject_ids":[1,2,3,4,5,6,7,10,12,14,15,16,17,18,19,20,21,22,23,24,25,27,28,29,31,33,36,41,42,43,44,45,46,47,48,49,54,55,56,57,58,59,60,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101,102,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121,122,123,124,125,126,127,128,129,130,131,132,133,134,135,136,137,138,139,141,142],"is_active":1,"weight":2},{"id":4,"name":"Первомайская 42к1","subject_ids":[1,3,4,5,6,7,10,12,13,14,15,16,17,18,19,20,21,23,24,27,28,29,31,33,36,41,42,43,44,45,46,47,48,49,54,55,56,57,58,59,60,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101,102,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121,122,123,124,125,126,127,128,129,130,131,132,133,134,135,136,137,138,139,140,142],"is_active":1,"weight":3},{"id":3,"name":"Сиреневый бульвар 62-1","subject_ids":[1,2,3,4,5,6,7,10,12,14,15,16,17,19,20,21,23,24,25,27,28,29,31,32,33,36,39,41,42,43,44,45,46,47,48,49,53,54,55,56,57,58,59,60,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101,102,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121,122,123,124,125,126,127,128,129,130,131,132,133,134,135,136,137,138,139,142],"is_active":1,"weight":4},{"id":6,"name":"АРТ","subject_ids":[2,4,14,25,26,32,39,47,103,116,124,128,143,144,145],"is_active":1,"weight":5},{"id":1,"name":"ОТДЕЛ ПРОДАЖ","subject_ids":[1,2,5,7,12,16,24,25,26,28,31,32,33,39,60],"is_active":1,"weight":6}]
let branches = [];
let leadSources = [];
let lastSyncCell = null;
let prevSyncDateString = undefined;
let lastSyncDate = undefined;

const authData = {
  email: 'E.m.belotserkovskaya@gmail.com',
  api_key: '75d8dd82-2d7d-11ef-b9b8-3cecefbdd1ae'
}

const PAGE_SIZE = 50;


// TODO:
// Statuses, reasons and etc - are not updated
// Run on schedule
// TS: https://developers.google.com/apps-script/guides/typescript


function _getCustomerFinalStatus(customer) {

  if (customer.lead_reject_name && customer.lead_reject_name !== 'No reason') {
    return customer.lead_reject_name
  }

  if (customer.lead_reject_name && customer.lead_reject_name === 'No reason') {
    if (['Активен', 'Активен/Взрослые', 'Не занимается', 'Временно на паузе', 'Планирует возобновить'].includes(customer.study_status_name)) {
      return 'Купил абонемент'
    }

    if (customer.study_status_name === "No status") {
      return customer.lead_status_name
    }

  }

  return '! Необработанный случай'

}


function loadAllData() {

  updateAuthToken();
  // console.log(authToken);

  loadBranches();
  loadLeadStatuses();
  loadCustomerStatuses();
  loadRejectReasons();
  loadLeadSources();


  loadPrevSyncDate();
  lastSyncDate = new Date(Date.now());
  loadCustomers();
  // after successfull sync, save last sync date
  saveLastSyncDate();
}



function loadCustomers() {
  const customers = [];


  branches.forEach(branch => {
    let currentPage = 0;
    let isCustomersLeft = true;
    while (isCustomersLeft) {
      Logger.log(`Fetching branch ${branch.id} customers page ${currentPage}`)
      const result = _fetchCustomers(currentPage, branch.id);
      Logger.log(`After fetching, branch ${branch.id} page ${currentPage}, result.page ${result.page}, result.count ${result.count}, result.total ${result.total}`);
      customers.push(...result.items);
      if (result.total <= (result.page + 1) * PAGE_SIZE) {
        // No data left to fetch
        isCustomersLeft = false;
        break;
      }
      currentPage++;
    }

    Logger.log(`Fetched ${customers.length} customers updated since ${prevSyncDateString}`);
  })


  // Inject status in customers
  customers.forEach(customer => {

    customer.lead_status_name = customer.lead_status_id ? _getLeadStatusTitle(customer.lead_status_id) : 'No status';
    customer.study_status_name = customer.study_status_id ? _getCustomerStatusTitle(customer.study_status_id) : 'No status';
    customer.lead_reject_name = customer.lead_reject_id ? _getLeadRejectReason(customer.lead_reject_id) : 'No reason';
    customer.lead_source_name = customer.lead_source_id ? _getLeadSource(customer.lead_source_id) : '';
    customer.lead_final_status = _getCustomerFinalStatus(customer);
    customer.branches = customer.branch_ids.join(',');
    customer.lead_branch = _getCustomerBranchName(customer);
  })

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('customers');
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // Get headers and data (assuming the first row is headers)
  const headers = values[0];
  const data = values.slice(1);

  // Find the index of important columns
  const idIndex = headers.indexOf('id');
  const updatedAtIndex = headers.indexOf('updated_at');

  // Create a map of existing customers by ID for easy lookup
  const existingCustomersMap = {};
  data.forEach(row => {
    const id = row[idIndex];
    if (id) {
      existingCustomersMap[id] = row;
    }
  });

  // Prepare arrays for new and updated rows
  const updatedRows = [];
  const newRows = [];

  customers.forEach(customer => {
    const existingCustomer = existingCustomersMap[customer.id];
    if (existingCustomer) {
      // Parse existing and new `updated_at` timestamps
      const existingUpdatedAt = new Date(existingCustomer[updatedAtIndex]);
      const newUpdatedAt = new Date(customer.updated_at);

      // Update the row if the new data is more recent
      if (newUpdatedAt > existingUpdatedAt) {
        const updatedRow = headers.map(header => customer[header] || '');
        updatedRows.push({ id: customer.id, row: updatedRow });
      }
    } else {
      // Prepare new customer rows for insertion
      const newRow = headers.map(header => customer[header] || '');
      newRows.push(newRow);
    }
  });

  // Update rows with newer data
  updatedRows.forEach(({ id, row }) => {
    const rowIndex = data.findIndex(row => row[idIndex] === id) + 2; // +2 for header row and 0-index adjustment
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  });

  // Append new customer rows
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }

  Logger.log('Customers updated successfully!');

}


function _fetchCustomers(pageNumber, branch, attempt = 1) {
  const url = BASE_URL + `/${branch}/customer/index`;
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify()
  }

  try {
    const response = UrlFetchApp.fetch(url, {
      'method': 'post',
      'contentType': 'application/json',
      'headers': {
        'X-ALFACRM-TOKEN': authToken
      },
      "muteHttpExceptions": true,
      'payload': JSON.stringify({
        "is_study": 2,
        "updated_at_from": prevSyncDateString,
        "removed": 1,
        "page": pageNumber
      })
    })
    const responseCode = response.getResponseCode();
    // console.log(`response coce ${responseCode}, ${typeof responseCode}, attempt ${attempt}`)

    if (responseCode === 200) {
      return JSON.parse(response.getContentText());

    } else if (responseCode === 401 && attempt === 1) {
      Logger.log(`Fetching customers with 401. Refetching with token update...`);
      updateAuthToken();
      return _fetchCustomers(pageNumber, branch, 2);
    }
    else {
      Logger.log(`Fetching error or 2nd time 401. responseCode: ${responseCode}. rethrow`);
      throw new Error('Cannot fetch customers');
    }


  } catch (e) {
    Logger.log(`Got error during customer fetch: ${e}`)
    throw e;
  }
}

// Load lead statuses
function loadLeadStatuses() {
  const url = BASE_URL_WITH_BRANCH + `/lead-status/index`;
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify()
  }
  const response = UrlFetchApp.fetch(url, {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'X-ALFACRM-TOKEN': authToken
    }
  })

  leadStatuses = JSON.parse(response.getContentText()).items;

  // const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('lead_statuses');
  // const dataRange = sheet.getDataRange();
  // const values = dataRange.getValues();

  // // Get headers and data (assuming the first row is headers)
  // const headers = values[0];
  // const data = values.slice(1);

  // // Find the index of important columns
  // const idIndex = headers.indexOf('id');
  // const updatedAtIndex = headers.indexOf('updated_at');

  // // Create a map of existing customers by ID for easy lookup
  // const existingStatusesMap = {};
  // data.forEach(row => {
  //   const id = row[idIndex];
  //   if (id) {
  //     existingStatusesMap[id] = row;
  //   }
  // });

  // // Prepare arrays for new and updated rows
  // const updatedRows = [];
  // const newRows = [];

  // leadStatuses.forEach(leadStatus => {
  //   const existingStatus = existingStatusesMap[leadStatus.id];
  //   if (existingStatus) {
  //     // Parse existing and new `updated_at` timestamps
  //     const existingUpdatedAt = new Date(existingStatus[updatedAtIndex]);
  //     const newUpdatedAt = new Date(leadStatus.updated_at);

  //     // Update the row if the new data is more recent
  //     if (newUpdatedAt > existingUpdatedAt) {
  //       const updatedRow = headers.map(header => leadStatus[header] || '');
  //       updatedRows.push({ id: leadStatus.id, row: updatedRow });
  //     }
  //   } else {
  //     // Prepare new customer rows for insertion
  //     const newRow = headers.map(header => leadStatus[header] || '');
  //     newRows.push(newRow);
  //   }
  // });

  // // Update rows with newer data
  // updatedRows.forEach(({ id, row }) => {
  //   const rowIndex = data.findIndex(row => row[idIndex] === id) + 2; // +2 for header row and 0-index adjustment
  //   sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  // });

  // // Append new customer rows
  // if (newRows.length > 0) {
  //   sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  // }

  Logger.log('leadStatuses updated successfully!');

}

// load customer statuses
function loadCustomerStatuses() {
  const url = BASE_URL_WITH_BRANCH + `/study-status/index`;
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify()
  }
  const response = UrlFetchApp.fetch(url, {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'X-ALFACRM-TOKEN': authToken
    }
  })

  customerStatuses = JSON.parse(response.getContentText()).items;

  // const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('customer_statuses');
  // const dataRange = sheet.getDataRange();
  // const values = dataRange.getValues();

  // // Get headers and data (assuming the first row is headers)
  // const headers = values[0];
  // const data = values.slice(1);

  // // Find the index of important columns
  // const idIndex = headers.indexOf('id');
  // const updatedAtIndex = headers.indexOf('updated_at');

  // // Create a map of existing customers by ID for easy lookup
  // const existingStatusesMap = {};
  // data.forEach(row => {
  //   const id = row[idIndex];
  //   if (id) {
  //     existingStatusesMap[id] = row;
  //   }
  // });

  // // Prepare arrays for new and updated rows
  // const updatedRows = [];
  // const newRows = [];

  // customerStatuses.forEach(leadStatus => {
  //   const existingStatus = existingStatusesMap[leadStatus.id];
  //   if (existingStatus) {
  //     // Parse existing and new `updated_at` timestamps
  //     const existingUpdatedAt = new Date(existingStatus[updatedAtIndex]);
  //     const newUpdatedAt = new Date(leadStatus.updated_at);

  //     // Update the row if the new data is more recent
  //     if (newUpdatedAt > existingUpdatedAt) {
  //       const updatedRow = headers.map(header => leadStatus[header] || '');
  //       updatedRows.push({ id: leadStatus.id, row: updatedRow });
  //     }
  //   } else {
  //     // Prepare new customer rows for insertion
  //     const newRow = headers.map(header => leadStatus[header] || '');
  //     newRows.push(newRow);
  //   }
  // });

  // // Update rows with newer data
  // updatedRows.forEach(({ id, row }) => {
  //   const rowIndex = data.findIndex(row => row[idIndex] === id) + 2; // +2 for header row and 0-index adjustment
  //   sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  // });

  // // Append new customer rows
  // if (newRows.length > 0) {
  //   sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  // }

  Logger.log('Customer Statuses updated successfully!');

}


// Load lead reject reasons
function loadRejectReasons() {
  const url = BASE_URL_WITH_BRANCH + `/lead-reject/index`;
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify()
  }
  const response = UrlFetchApp.fetch(url, {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'X-ALFACRM-TOKEN': authToken
    }
  })

  leadRejectReasons = JSON.parse(response.getContentText()).items;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('lead_reject');
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // Get headers and data (assuming the first row is headers)
  const headers = values[0];
  const data = values.slice(1);

  // Find the index of important columns
  const idIndex = headers.indexOf('id');
  const updatedAtIndex = headers.indexOf('updated_at');

  // Create a map of existing customers by ID for easy lookup
  const existingStatusesMap = {};
  data.forEach(row => {
    const id = row[idIndex];
    if (id) {
      existingStatusesMap[id] = row;
    }
  });

  // Prepare arrays for new and updated rows
  const updatedRows = [];
  const newRows = [];

  leadRejectReasons.forEach(reason => {
    const existingStatus = existingStatusesMap[reason.id];
    if (existingStatus) {
      // Parse existing and new `updated_at` timestamps
      const existingUpdatedAt = new Date(existingStatus[updatedAtIndex]);
      const newUpdatedAt = new Date(reason.updated_at);

      // Update the row if the new data is more recent
      if (newUpdatedAt > existingUpdatedAt) {
        const updatedRow = headers.map(header => reason[header] || '');
        updatedRows.push({ id: reason.id, row: updatedRow });
      }
    } else {
      // Prepare new customer rows for insertion
      const newRow = headers.map(header => reason[header] || '');
      newRows.push(newRow);
    }
  });

  // Update rows with newer data
  updatedRows.forEach(({ id, row }) => {
    const rowIndex = data.findIndex(row => row[idIndex] === id) + 2; // +2 for header row and 0-index adjustment
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  });

  // Append new customer rows
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }

  Logger.log('LeadRejectReasons updated successfully!');

}


// Load lead sources
function loadLeadSources() {
  const url = BASE_URL_WITH_BRANCH + `/lead-source/index`;
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify()
  }
  const response = UrlFetchApp.fetch(url, {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'X-ALFACRM-TOKEN': authToken
    }
  })

  leadSources = JSON.parse(response.getContentText()).items;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('lead_source');
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // Get headers and data (assuming the first row is headers)
  const headers = values[0];
  const data = values.slice(1);

  // Find the index of important columns
  const idIndex = headers.indexOf('id');
  const updatedAtIndex = headers.indexOf('updated_at');

  // Create a map of existing customers by ID for easy lookup
  const existingStatusesMap = {};
  data.forEach(row => {
    const id = row[idIndex];
    if (id) {
      existingStatusesMap[id] = row;
    }
  });

  // Prepare arrays for new and updated rows
  const updatedRows = [];
  const newRows = [];

  leadSources.forEach(source => {
    const existingStatus = existingStatusesMap[source.id];
    if (existingStatus) {
      // Parse existing and new `updated_at` timestamps
      const existingUpdatedAt = new Date(existingStatus[updatedAtIndex]);
      const newUpdatedAt = new Date(source.updated_at);

      // Update the row if the new data is more recent
      if (newUpdatedAt > existingUpdatedAt) {
        const updatedRow = headers.map(header => source[header] || '');
        updatedRows.push({ id: source.id, row: updatedRow });
      }
    } else {
      // Prepare new customer rows for insertion
      const newRow = headers.map(header => source[header] || '');
      newRows.push(newRow);
    }
  });

  // Update rows with newer data
  updatedRows.forEach(({ id, row }) => {
    const rowIndex = data.findIndex(row => row[idIndex] === id) + 2; // +2 for header row and 0-index adjustment
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  });

  // Append new customer rows
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }

  Logger.log('LeadSources updated successfully!');

}

// Load branches
function loadBranches() {
  const url = BASE_URL_WITH_BRANCH + `/branch/index`;
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify()
  }
  const response = UrlFetchApp.fetch(url, {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'X-ALFACRM-TOKEN': authToken
    }
  })

  branches = JSON.parse(response.getContentText()).items;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('branches');
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // Get headers and data (assuming the first row is headers)
  const headers = values[0];
  const data = values.slice(1);

  // Find the index of important columns
  const idIndex = headers.indexOf('id');
  const updatedAtIndex = headers.indexOf('updated_at');

  // Create a map of existing customers by ID for easy lookup
  const existingBranches = {};
  data.forEach(row => {
    const id = row[idIndex];
    if (id) {
      existingBranches[id] = row;
    }
  });

  // Prepare arrays for new and updated rows
  const updatedRows = [];
  const newRows = [];

  branches.forEach(source => {
    const isExistingBranch = existingBranches[source.id];
    if (isExistingBranch) {
      // Parse existing and new `updated_at` timestamps
      const existingUpdatedAt = new Date(isExistingBranch[updatedAtIndex]);
      const newUpdatedAt = new Date(source.updated_at);

      // Update the row if the new data is more recent
      if (newUpdatedAt > existingUpdatedAt) {
        const updatedRow = headers.map(header => source[header] || '');
        updatedRows.push({ id: source.id, row: updatedRow });
      }
    } else {
      // Prepare new branch rows for insertion
      const newRow = headers.map(header => source[header] || '');
      newRows.push(newRow);
    }
  });

  // Update rows with newer data
  updatedRows.forEach(({ id, row }) => {
    const rowIndex = data.findIndex(row => row[idIndex] === id) + 2; // +2 for header row and 0-index adjustment
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  });

  // Append new branch rows
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }

  Logger.log('Branches updated successfully!');

}

// Update auth token
function updateAuthToken() {

  const url = BASE_URL_WITH_BRANCH + `/auth/login`;
  const response = UrlFetchApp.fetch(url, {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(authData)
  })

  authToken = JSON.parse(response.getContentText()).token;

  // const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('token');
  // console.log(data);
  //sheet.appendRow(JSON.stringify(data));

}


function loadPrevSyncDate() {
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('settings');
  lastSyncCell = settingsSheet.getRange('B1'); // Cell where the last sync date is stored
  const lastSyncDate = lastSyncCell.getValue() ? new Date(lastSyncCell.getValue()) : new Date(Date.now());

  prevSyncDateString = lastSyncDate.toISOString();

  console.log(`Loaded prev sync date ${prevSyncDateString}`);

}

function saveLastSyncDate() {
  lastSyncCell.setValue(lastSyncDate.toISOString());
  console.log(`Saved last sync date ${lastSyncDate.toISOString()}`);
}


function _getLeadStatusTitle(targetId) {
  const statuses = leadStatuses.filter(status => {
    return status.id === targetId
  });
  if (statuses.length === 0) { return 'Unknown status' }
  return statuses[0].name;
}

function _getCustomerStatusTitle(targetId) {
  const statuses = customerStatuses.filter(status => {
    return status.id === targetId
  });
  if (statuses.length === 0) { return 'Unknown status' }
  return statuses[0].name;
}

function _getLeadRejectReason(targetId) {
  const statuses = leadRejectReasons.filter(status => {
    return status.id === targetId
  });
  if (statuses.length === 0) { return 'Unknown reason' }
  return statuses[0].name;
}

function _getLeadSource(targetId) {
  const statuses = leadSources.filter(status => {
    return status.id === targetId
  });
  if (statuses.length === 0) { return 'Unknown source' }
  return statuses[0].name;
}


function _getCustomerBranchName(customer) {
  if (customer && customer.branch_ids && customer.branch_ids.length > 0) {
    const foundBranches = branches.filter(branch => {
      return branch.id === customer.branch_ids[0]
    });
    if (foundBranches.length === 0) {
      return 'UNKNOWN BRANCH'
    }
    return foundBranches[0].name
  }
  return '';
}