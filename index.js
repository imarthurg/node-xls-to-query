const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const { getValueParsed } = require('./strategies');
const { getCellRef, getCellAddress } = require('./lib/cell');

const filePath = process.env.FILE_PATH || `${path.resolve('test.xlsx')}`;

const getWorkbook = () => XLSX.readFile(filePath, {});

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
    col += 1;
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
};

const getQuery = (woorksheet, tableName, columnNames, row) => {
  const tableReference = `public.${tableName}`;
  const query = `UPDATE ${tableReference} SET `;
  const queryWithSet = columnNames.reduce((accQuery, columnName, index) => {
    let queryIncrement = accQuery;
    const cellRef = getCellRef(getCellAddress(row, index));
    const cellValue = getValueParsed(woorksheet[cellRef]);

    if (index !== 0 && index > 1) {
      queryIncrement += `, "${columnName}" = ${cellValue}`;
    } else if (index !== 0) {
      queryIncrement += `"${columnName}" = ${cellValue}`;
    }

    return queryIncrement;
  }, query);

  const cellRef = getCellRef(getCellAddress(row, 0));
  const cellValue = getValueParsed(woorksheet[cellRef]);
  return `${queryWithSet} WHERE "${columnNames[0]}" = ${cellValue};`;
};

/**
 *
 * @param {string} sheetName
 */
const generateSqlQueries = (sheetName, tableName) => {
  const workbook = getWorkbook();
  const worksheet = workbook.Sheets[sheetName];
  const columnNames = getColumnNames(worksheet);
  const dataRowsCount = getDataRowsCount(worksheet);

  const queries = [];
  for (let row = 1; row <= dataRowsCount; row += 1) {
    queries.push(getQuery(worksheet, tableName, columnNames, row));
  }

  fs.writeFileSync(`${filePath}.sql`, '');
  queries.forEach((q) => {
    fs.appendFileSync(`${filePath}.sql`, `${q} \n`);
  });

  console.log('Done!');
};

generateSqlQueries('new_input', 'input');

module.exports = {
  getWorkbook,
};
