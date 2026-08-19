// Shared between the browser app (src/App.jsx) and the bake-upload script
// (scripts/bake-upload.js), so header matching stays identical in both places.
const FIELD_ALIASES = [
  { key: 'VIN',                  aliases: ['vin', 'vehicle identification number'] },
  { key: 'STOCK',                aliases: ['stock #', 'stk', 'stock number', 'stock'] },
  { key: 'YEAR',                 aliases: ['yr', 'model year', 'year'] },
  { key: 'MAKE',                 aliases: ['manufacturer', 'brand', 'make', 'vehicle'] },
  { key: 'MODEL',                aliases: ['vehicle model', 'model', 'vehiclemodel'] },
  { key: 'TRIM',                 aliases: ['series', 'trim level', 'edition', 'trim'] },
  { key: 'COLOR',                aliases: ['ext color', 'exterior color', 'colour', 'color', 'col.', 'col'] },
  { key: 'ODOMETER',             aliases: ['odo', 'mileage', 'miles', 'odometer'] },
  { key: 'AGE',                  aliases: ['days', 'days in stock', 'lot age', 'age'] },
  { key: 'PRICE',                aliases: ['asking price', 'list price', 'sale price', 'price', 'listing price', 'listed price'] },
  { key: 'WHOLESALE / TRADE-IN', aliases: ['trade', 'trade-in value', 'wholesale', 'wholesale value', 'j.d. power trade in clean', 'j.d. power trade clean'] },
  { key: 'RETAIL / MSRP',       aliases: ['retail value', 'msrp', 'market value', 'j.d. power retail clean', 'j.d. power retail in clean'] },
];

function normalizeHeader(str) {
  return String(str).replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim().toUpperCase();
}

function matchAlias(rawHeader) {
  const n = normalizeHeader(rawHeader);
  for (const field of FIELD_ALIASES) {
    if (field.key === n) return field.key;
    for (const alias of field.aliases) {
      if (normalizeHeader(alias) === n) return field.key;
    }
  }
  return null;
}

function buildInitialMap(headers) {
  const initialMap = {};
  const usedHeaders = new Set();
  for (const field of FIELD_ALIASES) {
    let matched = null;
    for (const header of headers) {
      if (usedHeaders.has(header)) continue;
      if (matchAlias(header) === field.key) { matched = header; break; }
    }
    if (matched) { initialMap[field.key] = matched; usedHeaders.add(matched); }
    else          { initialMap[field.key] = ''; }
  }
  return initialMap;
}

module.exports = { FIELD_ALIASES, normalizeHeader, matchAlias, buildInitialMap };
