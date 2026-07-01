function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('NS Investment Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDropdownData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName('Investment Workspace');
  var data = ws.getDataRange().getValues();

  // Asset Classes (Column C, rows 3-6)
  var assetClasses = [];
  for (var i = 2; i <= 5; i++) {
    if (data[i] && data[i][2]) assetClasses.push(data[i][2]);
  }

  // Class Selections (Row 12 = headers, rows 13+ = options)
  var classSelections = {};
  var headers = [];
  for (var c = 0; c < 4; c++) {
    if (data[11] && data[11][c]) headers.push(data[11][c]);
  }

  for (var h = 0; h < headers.length; h++) {
    var options = [];
    for (var r = 12; r < 20; r++) {
      if (data[r] && data[r][h] && data[r][h] !== '') {
        options.push(data[r][h]);
      }
    }
    classSelections[headers[h]] = options;
  }

  // Map Asset Class names to Class Selection header names
  var assetToHeader = {
    'Cash & Cash Equivalent': 'Cash & Cash Equivalents',
    'Commodities': 'Commodities',
    'Fixed Income': 'Fixed Income',
    'Hedge Fund': 'Hedge Funds'
  };

  var mappedClassSelections = {};
  for (var ac = 0; ac < assetClasses.length; ac++) {
    var headerName = assetToHeader[assetClasses[ac]] || assetClasses[ac];
    mappedClassSelections[assetClasses[ac]] = classSelections[headerName] || [];
  }

  // Duration options (Column A, rows 24-26)
  var durations = [];
  for (var i = 23; i <= 25; i++) {
    if (data[i] && data[i][0]) durations.push(data[i][0]);
  }

  // Holding Periods by Duration (Columns B, C, D starting row 24)
  var holdingPeriods = {};
  var durationHeaders = [];
  if (data[22]) {
    for (var c = 1; c <= 3; c++) {
      if (data[22][c]) durationHeaders.push({ col: c, name: data[22][c] });
    }
  }

  var durationToHeader = {
    'Short-Term': 'Short-Term Holding Period',
    'Intermediate': 'Intermediate Holding Period',
    'Long-Term': 'Long-Term Holding Period'
  };

  for (var d = 0; d < durations.length; d++) {
    var dur = durations[d];
    var headerName = durationToHeader[dur];
    var colIdx = -1;
    for (var dh = 0; dh < durationHeaders.length; dh++) {
      if (durationHeaders[dh].name === headerName) {
        colIdx = durationHeaders[dh].col;
        break;
      }
    }
    if (colIdx >= 0) {
      var periods = [];
      for (var r = 23; r < 30; r++) {
        if (data[r] && data[r][colIdx] && data[r][colIdx] !== '') {
          periods.push(data[r][colIdx]);
        }
      }
      holdingPeriods[dur] = periods;
    }
  }

  // Investor Types (Column A, rows 3-6)
  var investorTypes = [];
  for (var i = 2; i <= 5; i++) {
    if (data[i] && data[i][0]) investorTypes.push(data[i][0]);
  }

  return JSON.stringify({
    assetClasses: assetClasses,
    classSelections: mappedClassSelections,
    durations: durations,
    holdingPeriods: holdingPeriods,
    investorTypes: investorTypes
  });
}

function submitContactForm(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Contact DB');

  // Find the header row (row 7)
  var headerRow = 7;
  var lastRow = sheet.getLastRow();
  var newRow = lastRow + 1;
  if (newRow <= headerRow) newRow = headerRow + 1;

  var row = [
    formData.accountOpeningDate,
    formData.investorType,
    formData.companyName,
    formData.accountHolderName,
    formData.phoneNumber,
    formData.emailAddress,
    formData.mailingAddress,
    formData.addressLine2,
    formData.cityTown,
    formData.provinceState,
    formData.postcode,
    formData.country,
    formData.secondContactName,
    formData.secondContactPhone,
    formData.secondContactEmail,
    formData.beneficiaryName,
    formData.beneficiaryPhone,
    formData.beneficiaryEmail
  ];

  sheet.getRange(newRow, 1, 1, row.length).setValues([row]);
  return 'Contact information saved successfully.';
}

function submitBankingForm(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Banking DB');

  var headerRow = 3;
  var lastRow = sheet.getLastRow();
  var newRow = lastRow + 1;
  if (newRow <= headerRow) newRow = headerRow + 1;

  var row = [
    formData.accountName,
    formData.accountNumber,
    formData.bankName,
    formData.bankAccountName,
    formData.bankAccountNumber,
    formData.routingNumber,
    formData.iban,
    formData.swiftCode,
    formData.bankAddress,
    formData.bankAddressLine2,
    formData.bankCity,
    formData.bankProvinceState,
    formData.bankPostcode,
    formData.bankCountry,
    formData.bankPhone,
    formData.acctManagerName,
    formData.acctManagerEmail
  ];

  sheet.getRange(newRow, 1, 1, row.length).setValues([row]);
  return 'Banking information saved successfully.';
}

function submitInvestmentForm(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('LP Selection');

  var headerRow = 3;
  var lastRow = sheet.getLastRow();
  var newRow = lastRow + 1;
  if (newRow <= headerRow) newRow = headerRow + 1;

  var row = [
    formData.accountNumber,
    formData.accountName,
    formData.assetClass,
    formData.classSelection,
    formData.duration,
    formData.holdingPeriod,
    formData.committedCapital,
    formData.depositInterval,
    formData.initialInvestment,
    formData.capitalCallDate,
    formData.fundingDate,
    formData.maturityDate
  ];

  sheet.getRange(newRow, 1, 1, row.length).setValues([row]);
  return 'Investment selection saved successfully.';
}
