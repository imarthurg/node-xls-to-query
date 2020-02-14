const XLSX = require('xlsx');

/**
 *
 * @param {number} row
 * @param {number} column
 * @returns XLSX.CellAddress
 */
const getCellAddress = (row, column) => ({ r: row, c: column });

/**
 *
 * @param {XLSX.CellAddress} cellToCheck
 */
const getCellRef = (cellToCheck) => XLSX.utils.encode_cell(cellToCheck);

module.exports = {
  getCellAddress,
  getCellRef,
};
