/* 
This file contains Javascript code designed for use with Google Apps Script. 

Instructions:
1. Create a new Google Apps Script project.
2. Copy and paste this code into the Apps Script editor.
3. Populate the required Apps Script properties under the Apps Script settings menu. (See lines 38-41 for the required constants.)
4. Run the main() function from the editor and grant the required permissions.

*/

/* Global variables */

// Alma Analytics API URL
const analyticsURL =
  "https://api-na.hosted.exlibrisgroup.com/almaws/v1/analytics/reports";

// Namespaces for the Analytics API XML
const rowNameSpace = XmlService.getNamespace(
    "urn:schemas-microsoft-com:xml-analysis:rowset",
  ),
  schemaNameSpace = XmlService.getNamespace(
    "xsd",
    "http://www.w3.org/2001/XMLSchema",
  );

const transactionKeyColumn = "Transaction Id";

const transactionDateColumn = "Actual Transaction Date";

const initialColumsForView = ["PO Line Reference", "PO Line Title", "Vendor Name",  
"Actual Transaction Date", "Transaction Amount", "Transaction Item Sub Type", 
"PO Line Total Price"];

var config = {};
// Load configuration values from the Apps Script properties service
try {
  const scriptProperties = PropertiesService.getScriptProperties();
  config.apiKey = scriptProperties.getProperty("ALMA_API_KEY");
  config.reportPath = scriptProperties.getProperty("REPORT_PATH");
  config.spreadsheetTab = scriptProperties.getProperty("DATA_TAB");
  config.spreadsheetId = scriptProperties.getProperty("SPREADSHEET_ID");
} catch (err) {
  Logger.log("Unable to load properties; failed with error %s", err.message);
  throw err;
}

const filters = {
  // Find all transactions from the previous month
  previousMonth: {
    groupKey: "PO Line Reference",
    filterKey: "Actual Transaction Date",
    condition: (function () {
      // returns a conditional function with a closure
      const monthPrior = new Date();
      monthPrior.setMonth(monthPrior.getMonth() - 1); // date for previous month
      monthPrior.setDate(1); // set to the first date of the month
      return (value) => Date.parse(value) >= monthPrior;
    })(),
    // for sorting the groups
    // sorts by max transaction date descending
    sortFunc: (group1, group2) => {
      let dates1 = group1.map((row) => new Date(row['Actual Transaction Date']));
      let dates2 = group2.map((row) => new Date(row['Actual Transaction Date']));
      let maxDate1 = Math.max(...dates1);
      let maxDate2 = Math.max(...dates2);
      return maxDate2 - maxDate1;
    },
    tabName: "preceding-month",
  },
  // call this field's function with the numDays parameter in order to get the filter object
  previousNDays: function (numDays) {
    return {
      groupKey: "PO Line Reference",
      filterKey: "Actual Transaction Date",
      condition: (function () {
        const daysPrior = new Date();
        daysPrior.setDate(daysPrior.getDate() - numDays); // date for previous numDays
        return (value) => Date.parse(value) >= daysPrior;
      })(),
      tabName: `last-${numDays}-days`,
    };
  },
  noDisencumbrance: {
    filterKey: "Transaction Item Sub Type",
    condition: (value) => value == "ENCUMBRANCE" || value == "EXPENDITURE",
  },
};

function createConditionalFormatting(sheet, range, groupKeyColIndex) {
  const groupKeyCol = String.fromCharCode(groupKeyColIndex + 65);
  const formula = '=ISODD(MATCH($' + groupKeyCol + '2, UNIQUE($' + groupKeyCol + '$2:$' + groupKeyCol + '), 0))';
  try {
    const rule = SpreadsheetApp.newConditionalFormatRule()
                  .whenFormulaSatisfied(formula)
                 .setBackground('#B7E1CD')
                 .setRanges([range])
                 .build();
    const rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
  } catch(e) {
    Logger.error("Error creating conditional formatting rule.", e)
  }
}

function encodePath(path) {
  // encodes an Alma Analytics path, replacing the ASCII characters for forward slash and whitespace with the appropriate code
  return path.replace(/\//g, "%2F").replace(/\s/g, "%20");
}

class Table {
  constructor(rows) {
    // when loading data from the spreadsheet, the first row should contain the column headings
    this.columns = rows[0];
    // create an object for each row, mapping column headings to values
    this.data = rows.slice(1).map((row) =>
      row.reduce((rowObj, cellValue, i) => {
        rowObj[this.columns[i]] = cellValue;
        return rowObj;
      }, {}),
    );
    // new rows to be added
    this.additions = [];
  }
  mapTransactions() {
    // for each existing row, associate it with its unique identifier
    this.rowMap = this.data.reduce((rowMap, row) => {
      rowMap.set(row[transactionKeyColumn], row);
      return rowMap;
    }, new Map());
  }

  update(row) {
    // if the provided row's unique identifier is not in the exising dataset, add it and include a timestamp
    if (!this.rowMap.has(row[transactionKeyColumn])) {
      let transactionDate = new Date();
      // subtract one day because of the Analytics lag
      transactionDate.setDate(transactionDate.getDate() - 1);
      row[transactionDateColumn] = transactionDate.toJSON();
      this.additions.push(row);
    }
  }

  toRange() {
    // If no additions, nothing to add
    if (this.additions.length == 0) return;
    // Add the column headers to the new range, if the existing dataset is empty
    let range = this.data.length == 0 ? [this.columns] : [];
    // Add the new rows to the range
    for (row of this.additions) {
      // construct the row by extracting the values for each column header -- preserves original column order
      range.push(this.columns.map((columnHeading) => row[columnHeading]));
    }
    return range;
  }

  numRows() {
    // include a row for the column headers if not appending
    if (this.data.length == 0) return this.additions.length + 1;
    return this.additions.length;
  }

  length() {
    return this.data.length + this.additions.length + 1;
  }

  rangeDimensions() {
    // Starting with a blank resport, we'll need to add the columns
    if (!this.columns.includes(transactionDateColumn)) {
      this.columns = this.additions.reduce((colSet, row) => {
        for (let key of Object.keys(row)) {
          if (!colSet.includes(key)) {
            colSet.push(key);
          }
        }
        return colSet;
      }, []);
    }
    // dimensions of the updates to the dataset
    return {
      firstRow: this.data.length == 0 ? 1 : this.data.length + 2, // if data already present, we need to account for the column headings
      firstCol: 1,
      numRows: this.numRows(),
      numCols: this.columns.length,
    };
  }
}

class TableFromXml {
  constructor(rowNameSpace) {
    this.data = [];
    this.rowNameSpace = rowNameSpace;
  }

  *[Symbol.iterator]() {
    // convenience iterator
    for (let row of this.data) {
      yield row;
    }
  }

  extend(rows, xmlColumnMap) {
    /* Parses XML rows into JS objects where the keys correspond to the column headings provided in the xmlColumnMap parameter */
    this.data = this.data.concat(
      rows.reduce((tableArray, rowElement) => {
        // Maps the list of column names to their generic <Column> elements in the XML, parsing the data in each
        let row = xmlColumnMap.reduce((rowObj, column) => {
          let cellValue = rowElement.getChildText(
            column.name,
            this.rowNameSpace,
          );
          if (!cellValue) cellValue = rowElement.getChildText(column.name); // namespaces are present only in the first page of results
          rowObj[column.columnHeading] = cellValue;
          return rowObj;
        }, {});
        tableArray.push(row);
        return tableArray;
      }, []),
    );
  }
}

class Filter {
  constructor({ filterKey, condition }) {
    this.filterKey = filterKey;
    this.condition = condition;
  }
  applyFilter(data) {
    return data.filter((row) => this.condition(row[this.filterKey]));
  }
}

class FilterGroups {
  constructor({ groupKey, filterKey, condition, sortFunc }) {
    this.groupKey = groupKey;
    this.filterKey = filterKey;
    this.condition = condition;
    if (sortFunc != undefined) {
      this.sortFunc = sortFunc;
    } else {
      // If not defined, do nothing
      this.sortFunc = (a, b) => 0;
    }
  }

  groupBy(data) {
    return Map.groupBy(data, (row) => row[this.groupKey]);
  }

  filter(groups) {
    return groups.values().filter((group) => {
      let keep = false;
      for (let row of group) {
        if (this.condition(row[this.filterKey])) {
          keep = true;
        }
      }
      return keep;
    });
  }

  applyFilter(data) {
    return Array.from(this.filter(this.groupBy(data))).sort(this.sortFunc);
  }
}

class FilteredTable {
  constructor(data, filters, tabName, hiddenColumns) {
    this.data = data;
    this.filters = filters;
    this.hiddenColumns = hiddenColumns || [];
  }

  toRange() {
    // Extract column headings before filtering
    let columns = this.data.reduce((colSet, row) => {
      for (let key of Object.keys(row)) {
        if (!colSet.includes(key) && !this.hiddenColumns.includes(key)) {
          colSet.push(key);
        }
      }
      return colSet;
    }, []);
    // Put any designated columns first
    columns = columns.filter((c) => initialColumsForView.includes(c)).concat(columns.filter((c) => !initialColumsForView.includes(c)));
    // apply filters
    // assumes that any filters that do not group precede those that do group
    for (let filter of this.filters) {
      this.data = filter.applyFilter(this.data);
    }

    return [columns].concat(
      this.data.reduce(
        (rows, group) =>
          rows.concat(group.map((row) => columns.map((c) => row[c]))),
        [],
      ),
    );
  }
}

function test() {
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
  } catch (e) {
    Logger.log("Unable to open spreadsheet");
    throw e;
  }
  // get the sheet corresponding to this report
  const sheet = spreadsheet.getSheetByName(config.spreadsheetTab);
  const extantData = sheet.getDataRange();
  const table = new Table(extantData.getValues());
  let monthFilter = new FilterGroups(filters.previousMonth);
  let typeFilter = new Filter(filters.noDisencumbrance);
  let ft = new FilteredTable(table.data, [typeFilter, monthFilter]);
  let dayFilter = new FilterGroups(filters.previousNDays(2));
  ft = new FilteredTable(table.data, [dayFilter]);
}

function main() {
  try {
    var table = callAnalyticsAPI({
      apiKey: config.apiKey,
      reportPath: config.reportPath,
    });
  } catch (e) {
    if (e == "Query failed!") {
      Logger.log("Query failed on " + config.reportPath);
    }
    throw e;
  }
  // Now pass  the data to the spreadsheetApp
  sendToSpreadsheet(table);
}

function callAnalyticsAPI(report) {
  /* Calls the Analtyics API to fetch a given report. Handles pagination when more than 1000 results are returned. Report should be a JS object with an API key and Analytics path as properties. */
  var headers = { Authorization: "apikey " + report.apiKey },
    params = { headers: headers, muteHttpExceptions: true },
    query = "?limit=1000&path=" + encodePath(report.reportPath),
    token = "", // initially, no token parameter needed
    table = new TableFromXml(rowNameSpace),
    xmlColumnMap = [],
    isFinished = false; // flag that will be set to true when the last page of results is reached
  let response, parsedResponse;
  while (!isFinished) {
    response = UrlFetchApp.fetch(analyticsURL + query + token, params);
    // Check for valid response
    if (response.getResponseCode() != 200) {
      // throw exception
      Logger.log(analyticsURL + query);
      Logger.log(response.getContentText());
      throw "Query failed!";
    }

    parsedResponse = parseXMLResponse(response.getContentText(), xmlColumnMap);

    if (parsedResponse.isFinished == "true") isFinished = true;

    if (parsedResponse.resumptionToken)
      token = "&token=" + parsedResponse.resumptionToken; // if we have a token, save it to use in subsequent requests

    table.extend(parsedResponse.rows, parsedResponse.columnMap);
    xmlColumnMap = parsedResponse.columnMap;
  }
  return table;
}

function parseXMLResponse(data, columnMap) {
  /* Iterate over the XML results, building up a 2-D array where each inner array corresponds to a  single row. columnMap should be non-null when we are retrieving results beyond the first page. Otherwise, it will be created
from the XML data.*/

  let document = XmlService.parse(data),
    result = document.getRootElement().getChild("QueryResult"),
    isFinished = result.getChild("IsFinished").getText(), // Is the report done, or are there more pages to return?
    resumptionToken = result.getChild("ResumptionToken"),
    // rowset contains all the rows in the result (for this page of results)
    rows = result.getChild("ResultXml").getChild("rowset", rowNameSpace);
  if (columnMap.length == 0) {
    // If this is the first page of results, we need to get the column names
    columnMap = getColumnMap(rows).filter(function (column) {
      return column.columnHeading != "0"; // Ignore columns with '0' as the heading -- these are spurious columns inserted by the Analytics API on some reports
    });
    // each row is its own element
    try {
      rows = rows.getChildren("Row", rowNameSpace);
    } catch (e) {
      Logger.log("Problem fetching rows from the first page");
      throw e;
    }
  } else {
    try {
      rows = result
        .getChild("ResultXml")
        .getChild("rowset", rowNameSpace)
        .getChildren("Row", rowNameSpace); // If it's not the first page of results, the rows won't have the namespace attribute
    } catch (e) {
      Logger.log("Problem fetching rows from subsequent pages");
      throw e;
    }
  }
  return {
    isFinished: isFinished,
    resumptionToken: resumptionToken
      ? result.getChild("ResumptionToken").getText()
      : null, // Token for fetching more pages -- present only on the first page
    rows: rows,
    columnMap: columnMap,
  };
}

function getColumnMap(data) {
  // Because Analytics returns the columns in what appears to be arbitrary order, it is necessary to extract a mapping from the XML itself
  // This data is in enclosed in the following path //rowset/xsd:schema/xsd:complexTypes/xsd:sequence/
  // Using namespaces with XmlService is not straightforward, so the following is an inelegant method that avoids that problem

  // assumes that the <xsd:schema> element will always be the first child of <rowset>
  try {
    var sequence = data.getChildren()[0].getChildren()[0].getChildren()[0];
  } catch (e) {
    Logger.log("Problem fetching <sequence> element for column names");
    throw e;
  }
  // Returns a list of objects, each of which maps an element name to a columnHeading.
  try {
    let result = sequence.getChildren().map((element) => {
      let attributes = element.getAttributes();
      return attributes.reduce((attrObj, attr) => {
        let attrName = attr.getName();
        // Looking for two particular headings
        if (attrName == "name" || attrName == "columnHeading") {
          attrObj[attrName] = attr.getValue();
        }
        return attrObj;
      }, {});
    });
    return result;
  } catch (e) {
    Logger.log(
      "Error fetching children of <sequence> element for column names",
    );
    throw e;
  }
}

function sendToSpreadsheet(incomingTable) {
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
  } catch (e) {
    Logger.log("Unable to open spreadsheet");
    throw e;
  }
  // get the sheet corresponding to this report
  const sheet = spreadsheet.getSheetByName(config.spreadsheetTab);
  const extantData = sheet.getDataRange();
  const extantTable = new Table(extantData.getValues());
  extantTable.mapTransactions();
  for (row of incomingTable) {
    extantTable.update(row);
  } // Get the range of the required dimensions for the rows to be added
  const { firstRow, firstCol, numRows, numCols } =
    extantTable.rangeDimensions();
  let rangeValues = extantTable.toRange();
  if (rangeValues) {
    // Test for data that exceeds the current number of available rows
    let sheetLength = sheet.getMaxRows();
    let dataSize = extantTable.length();
    if (sheetLength < dataSize) {
      sheet.insertRowsAfter(sheetLength, dataSize - sheetLength);
    }
    let range = sheet.getRange(firstRow, firstCol, numRows, numCols);
    try {
      range.setValues(rangeValues);
      
    } catch (e) {
      Logger.log("Unable to save data to spreadsheet");
      throw e;
    }
  }
  createFilteredViews(extantTable, spreadsheet);
}

function createFilteredViews(table, spreadsheet) {
  let monthFilter = new FilterGroups(filters.previousMonth);
  let tabName = filters.previousMonth.tabName;
  let typeFilter = new Filter(filters.noDisencumbrance);
  let ft = new FilteredTable(
    table.data.concat(table.additions),
    [typeFilter, monthFilter],
    tabName,
  );
  let data = ft.toRange();
  try {
    let sheet = spreadsheet.getSheetByName(tabName);
    sheet.clear();
    let range = sheet.getRange(1, 1, data.length, data[0].length);
    range.setValues(data);
    // exclude column headers from formatting
    let formattingRange = sheet.getRange(2, 1, data.length, data[0].length);
    createConditionalFormatting(sheet, formattingRange, data[0].indexOf(monthFilter.groupKey))
  } catch (e) {
    Logger.log("Error saving filtered data");
    throw e;
  }
}
