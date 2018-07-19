const currencies = require('./currencies.json');

export const getCurrency = function(lcid) {
  if (lcid in currencies) {
    return currencies[lcid];
  }
  return null;
};
