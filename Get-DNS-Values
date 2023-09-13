/**
 * Google Apps Script to fetch DNS records for domain names listed in a Google Sheet.
 * 
 * Last Updated: 2023-09-11
 * Script Creator: Jonas Lund
 * 
 * How it works:
 * - Requires domain names to be listed in Column A of the Google Sheet named "DNSVALUE".
 * - If you add just one domain in Column A, the script will automatically fetch all necessary DNS info for that domain.
 * - Uses Google's Public DNS-over-HTTPS API to fetch DNS records.
 * - No additional APIs need to be enabled.
 * - Recommended to set up a Google Sheets trigger to run 'onEdit' when the sheet is edited.
 * - When running 'pullDnsChanges', it will take time due to rate-limiting considerations.
 *   The script fetches and writes 40 rows at a time, so if you have many domains, you'll need to wait.
 */

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("DNSVALUE");
  
  if (!sheet) {
    sheet = ss.insertSheet("DNSVALUE");
  }
  
  var lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) {
    lastColumn = 1;
  }
  
  var range = sheet.getRange(1, 1, 1, lastColumn);
  range.setFontWeight("bold");
  sheet.setFrozenRows(1);

  ui.createMenu('DNS-Tool')
      .addItem('Clear Sheet', 'clearSheet')
      .addItem('Pull DNS Changes', 'pullDnsChanges')
      .addToUi();
}

function pullDnsChanges() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('This operation will take a long time due to rate-limiting considerations.');
  clearSheet();
  fetchDnsRecords();
}

function clearSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DNSVALUE");
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  if (lastRow > 1) {
    sheet.getRange(2, 2, lastRow - 1, lastCol - 1).clearContent();
  }
}

function fetchDnsRecords() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DNSVALUE");
  var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
  var data = dataRange.getValues();
  
  for (var i = 0; i < data.length; i += 40) {
    for (var j = i; j < i + 40 && j < data.length; j++) {
      var domain = data[j][0];
      if (domain) {
        try {
          fetchAndWriteRecords(domain, "MX", j + 2, 2);
          fetchAndWriteRecords(domain, "TXT", j + 2, 3);
          fetchAndWriteSpfRecords(domain, j + 2, 4);
          fetchAndWriteRecords(domain, "A", j + 2, 5);
          fetchAndWriteRecords(domain, "CNAME", j + 2, 6);
          fetchAndWriteRecords(domain, "NS", j + 2, 7);
          fetchAndWriteRecords(domain, "AAAA", j + 2, 10);
          // DKIM
          var dkimSelectors = ["google", "google2", "selector1", "selector2"];
          var dkimRecords = [];
          for (var k = 0; k < dkimSelectors.length; k++) {
            var dkimDomain = dkimSelectors[k] + "._domainkey." + domain;
            var dkimRecord = fetchRecord(dkimDomain, "TXT");
            if (dkimRecord !== "No TXT records found") {
              dkimRecords.push(dkimSelectors[k] + ": " + dkimRecord);
            }
          }
          sheet.getRange(j + 2, 8).setValue(dkimRecords.join(", "));
          // DMARC
          var dmarcDomain = "_dmarc." + domain;
          fetchAndWriteRecords(dmarcDomain, "TXT", j + 2, 9);
        } catch (error) {
          Logger.log('Error fetching records for domain ' + domain + ': ' + error.toString());
        }
      }
    }
    SpreadsheetApp.flush(); // Apply all pending changes
  }
}

function fetchAndWriteRecords(domain, type, row, col) {
  var apiUrl = "https://dns.google.com/resolve?name=" + domain + "&type=" + type;
  var response = UrlFetchApp.fetch(apiUrl);
  var jsonResponse = JSON.parse(response.getContentText());
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DNSVALUE");
  
  if (jsonResponse.Answer) {
    var records = jsonResponse.Answer.map(function(record) {
      return record.data;
    }).join(", ");
    sheet.getRange(row, col).setValue(records);
  } else {
    sheet.getRange(row, col).setValue("No " + type + " records found");
  }
}

function fetchAndWriteSpfRecords(domain, row, col) {
  var apiUrl = "https://dns.google.com/resolve?name=" + domain + "&type=TXT";
  var response = UrlFetchApp.fetch(apiUrl);
  var jsonResponse = JSON.parse(response.getContentText());
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DNSVALUE");
  
  if (jsonResponse.Answer) {
    var spfRecords = jsonResponse.Answer.filter(function(record) {
      return record.data.startsWith("\"v=spf1") || record.data.startsWith("v=spf1");
    }).map(function(record) {
      return record.data;
    }).join(", ");
    
    if (spfRecords) {
      sheet.getRange(row, col).setValue(spfRecords);
    } else {
      sheet.getRange(row, col).setValue("No SPF records found");
    }
  } else {
    sheet.getRange(row, col).setValue("No SPF records found");
  }
}

function fetchRecord(domain, type) {
  var apiUrl = "https://dns.google.com/resolve?name=" + domain + "&type=" + type;
  var response = UrlFetchApp.fetch(apiUrl);
  var jsonResponse = JSON.parse(response.getContentText());
  
  if (jsonResponse.Answer) {
    return jsonResponse.Answer.map(function(record) {
      return record.data;
    }).join(", ");
  } else {
    return "No " + type + " records found";
  }
}

function onEdit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DNSVALUE");
  var range = e.range;
  
  if (range.getColumn() === 1 && range.getRow() > 1) {
    var domain = range.getValue();
    if (domain) {
      var row = range.getRow();
      fetchAndWriteRecords(domain, "MX", row, 2);
      fetchAndWriteRecords(domain, "TXT", row, 3);
      fetchAndWriteSpfRecords(domain, row, 4);
      fetchAndWriteRecords(domain, "A", row, 5);
      fetchAndWriteRecords(domain, "CNAME", row, 6);
      fetchAndWriteRecords(domain, "NS", row, 7);
      fetchAndWriteRecords(domain, "AAAA", row, 10);
      // DKIM
      var dkimSelectors = ["google", "google2", "selector1", "selector2"];
      var dkimRecords = [];
      for (var j = 0; j < dkimSelectors.length; j++) {
        var dkimDomain = dkimSelectors[j] + "._domainkey." + domain;
        var dkimRecord = fetchRecord(dkimDomain, "TXT");
        if (dkimRecord !== "No TXT records found") {
          dkimRecords.push(dkimSelectors[j] + ": " + dkimRecord);
        }
      }
      sheet.getRange(row, 8).setValue(dkimRecords.join(", "));
      // DMARC
      var dmarcDomain = "_dmarc." + domain;
      fetchAndWriteRecords(dmarcDomain, "TXT", row, 9);
    }
  }
}
