const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const { getValueParsed } = require('./strategies');

const filePath = process.env.FILE_PATH || `${path.resolve('test.xlsx')}`;

const getWorkbook = () => {
  return XLSX.readFile(filePath, {});
}

/**
 * Count how many cells was filled in sequence on the first row
 * 
 * @param {XLSX.Worksheet} worksheet 
 */
const getColumnNames = (worksheet) => {
  const columnNames = [];
  
  const row = 0;

  let col = 0;
  while (true) {
    const cellRef = getCellRef(getCellAddress(row, col));
    if (worksheet[cellRef] && worksheet[cellRef].v !== undefined) {
      columnNames.push(worksheet[cellRef].v);
    } else {
      break;
    }
    col++;
  }

  return columnNames;
};

const getDataRowsCount = (worksheet) => {
  let row = 0;
  while (true) {
    const cellRef = getCellRef(getCellAddress(row, 0));
    if (worksheet[cellRef] && worksheet[cellRef].v !== undefined) row++;
    else break;
  }

  return row - 1; // desconsidering the title row
}

const getQuery = (woorksheet, tableName, columnNames, row) => {
  const shortName = tableName.slice(0, 1);
  let query = `UPDATE public.${tableName} ${shortName} SET `;
  const queryWithSet = columnNames.reduce((accQuery, columnName, index) => {
    const cellRef = getCellRef(getCellAddress(row, index));
    const cellValue = getValueParsed(woorksheet[cellRef]);

    if (index !== 0 && index > 1) {
      accQuery += `, ${shortName}."${columnName}" = ${cellValue}`;
    } else if (index !== 0) {
      accQuery += `${shortName}."${columnName}" = ${cellValue}`;
    }

    return accQuery;
  }, query);

  const cellRef = getCellRef(getCellAddress(row, 0));
  const cellValue = getValueParsed(woorksheet[cellRef]);
  return `${queryWithSet} WHERE "${columnNames[0]} = ${cellValue};`;
}

/**
 * 
 * @param {string} sheetName 
 */
const generateSqlQueries = (sheetName, tableName) => {
  const workbook = getWorkbook();
  const worksheet = workbook['Sheets'][sheetName];
  const columnNames = getColumnNames(worksheet);
  const dataRowsCount = getDataRowsCount(worksheet);
  
  const queries = [];
  for(let row = 1; row <= dataRowsCount; row++) {
    queries.push(getQuery(worksheet, tableName, columnNames, row));
  }

  console.log(queries[0]);
  return queries;
};

/**
 * 
 * @param {number} row 
 * @param {number} column 
 * @returns XLSX.CellAddress
 */
const getCellAddress = (row, column) => {
  return { r: row, c: column };
}

/**
 * 
 * @param {XLSX.CellAddress} cellToCheck 
 */
const getCellRef = (cellToCheck) => {
  return XLSX.utils.encode_cell(cellToCheck);
}

generateSqlQueries('new_report_group', 'report_group');

module.exports = {
  getWorkbook,
}
