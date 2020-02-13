const getValueParsed = ({ v: value, t: type }) => {
  switch (type) {
    case 'b':
      return JSON.parse(value);
    case 'n':
      return value;
    case 's':
      return value === 'NULL' ? value : `'${value}'`;
    case 'd':
      return new Date(value);
    default:
      return value;
  } 
}

module.exports = {
  getValueParsed
}
