import React, { useState, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';
import {
  Upload, FileSpreadsheet, RefreshCw, ArrowRight,
  AlertTriangle, Check, TrendingUp, DollarSign, RotateCcw, Printer,
} from 'lucide-react';
import styles from './App.module.css';

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

const TAX_RATE = 0.075;
const REG_FEE  = 250;

function calcPayment(amtFin, termMonths, annualRate) {
  if (!amtFin || amtFin <= 0 || !termMonths) return 0;
  const r = annualRate / 100 / 12;
  if (r === 0) return amtFin / termMonths;
  return amtFin * (r * Math.pow(1 + r, termMonths)) / (Math.pow(1 + r, termMonths) - 1);
}

function calcDeal(row, colIdx, inputs, bookKey) {
  const get = (key) => {
    const i = colIdx[key];
    if (i === undefined) return 0;
    return parseFloat(String(row[i]).replace(/[^0-9.-]/g, '')) || 0;
  };
  const price       = get('PRICE');
  const book        = get(bookKey);
  const { downPayment, tradeAllowance, tradePayoff, dealerAddendum, term, rate } = inputs;
  const netTrade    = tradeAllowance - tradePayoff;
  const taxableBase = price + dealerAddendum;
  const tax         = taxableBase * TAX_RATE;
  const amtFin      = taxableBase + tax + REG_FEE - downPayment - netTrade;
  const payment     = calcPayment(amtFin, term, rate);
  const equityPct   = book > 0 ? (amtFin / book) * 100 : null;
  return { price, book, amtFin, payment, equityPct, netTrade, tax };
}

function fmt$(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return '$' + Math.round(n).toLocaleString();
}
function fmtPct(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return n.toFixed(1) + '%';
}

const DEFAULT_INPUTS = {
  downPayment: 0, tradeAllowance: 0, tradePayoff: 0,
  dealerAddendum: 0, term: 72, rate: 0, maxPayment: '',
};

function printDeals(deals, inputs, bookKey, colIdx) {
  const bookLabel = bookKey === 'WHOLESALE / TRADE-IN' ? 'Wholesale' : 'Retail / MSRP';
  const date = new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  const netTrade = inputs.tradeAllowance - inputs.tradePayoff;
  const tableRows = deals.map(({ row, price, book, amtFin, payment, equityPct, tax }, rank) => {
    const get = (key) => { const i = colIdx[key]; return i !== undefined ? row[i] : ''; };
    const vehicle = [get('YEAR'), get('MAKE'), get('MODEL'), get('TRIM')].filter(Boolean).join(' ') || 'Unknown';
    const isNeg = equityPct !== null && equityPct > 100;
    const eqColor = isNeg ? '#c0392b' : '#27ae60';
    return `<tr>
      <td style="color:#aaa;font-weight:700;text-align:center">#${rank + 1}</td>
      <td>
        <div style="font-weight:600;font-size:13px">${vehicle}</div>
        <div style="color:#888;font-size:11px;margin-top:3px">${[get('STOCK') ? 'STK ' + get('STOCK') : '', get('COLOR'), get('ODOMETER') ? Number(get('ODOMETER')).toLocaleString() + ' mi' : '', get('AGE') ? get('AGE') + ' days' : ''].filter(Boolean).join(' · ')}</div>
        ${get('VIN') ? `<div style="color:#666;font-size:10px;font-family:monospace;margin-top:2px">${get('VIN')}</div>` : ''}
      </td>
      <td style="text-align:right;font-family:monospace">${fmt$(price)}</td>
      <td style="text-align:right;font-family:monospace">${fmt$(tax + REG_FEE)}</td>
      <td style="text-align:right;font-family:monospace">${fmt$(book)}</td>
      <td style="text-align:right;font-family:monospace">${fmt$(amtFin)}</td>
      <td style="text-align:right;font-family:monospace;color:#d4a843;font-weight:600">${fmt$(payment)}/mo</td>
      <td style="text-align:center"><span style="background:${eqColor}22;color:${eqColor};border:1px solid ${eqColor}55;padding:3px 8px;border-radius:4px;font-weight:700;font-family:monospace;font-size:12px">${fmtPct(equityPct)}</span></td>
    </tr>`;
  }).join('');
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"/><title>Top Deals — ${date}</title>
  <style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:-apple-system,sans-serif;background:#fff;color:#111;padding:32px}
  .hdr{display:flex;justify-content:space-between;margin-bottom:20px;border-bottom:2px solid #d4a843;padding-bottom:14px}
  .params{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:18px;font-size:11px;color:#555}
  .params strong{color:#111}
  table{width:100%;border-collapse:collapse;font-size:12px}
  th{background:#111;color:#d4a843;font-size:10px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;padding:8px 10px;text-align:left}
  th:nth-child(n+3){text-align:right}th:last-child{text-align:center}
  td{padding:9px 10px;border-bottom:1px solid #eee;vertical-align:middle}
  tr:nth-child(even) td{background:#fafafa}
  </style></head><body>
  <div class="hdr"><div><h1 style="font-size:18px;font-weight:700;color:#d4a843;letter-spacing:-0.02em">GotEmDone</h1><p style="font-size:13px;font-weight:600;color:#111;margin-top:2px">Top Deals &nbsp;·&nbsp; ${date}</p></div>
  <div style="text-align:right;font-size:11px;color:#888">Book: <strong style="color:#111">${bookLabel}</strong></div></div>
  <div class="params">
    <span>Down: <strong>${fmt$(inputs.downPayment)}</strong></span>
    <span>Trade Allow: <strong>${fmt$(inputs.tradeAllowance)}</strong></span>
    <span>Trade Payoff: <strong>${fmt$(inputs.tradePayoff)}</strong></span>
    <span>Net Trade: <strong>${fmt$(netTrade)}</strong></span>
    <span>Addendum: <strong>${fmt$(inputs.dealerAddendum)}</strong></span>
    <span>Term: <strong>${inputs.term} mo</strong></span>
    <span>Rate: <strong>${inputs.rate}%</strong></span>
  </div>
  <table><thead><tr><th>#</th><th>Vehicle</th><th>Price</th><th>Tax+Reg</th><th>${bookLabel}</th><th>Amt Fin.</th><th>Payment</th><th>Equity%</th></tr></thead>
  <tbody>${tableRows}</tbody></table>
  <div style="margin-top:16px;font-size:10px;color:#aaa;text-align:right">Generated by GotEmDone</div>
  </body></html>`;
  const win = window.open('', '_blank', 'width=900,height=700');
  win.document.write(html);
  win.document.close();
  win.onload = () => { win.focus(); win.print(); };
}

export default function App() {
  const [stage, setStage]           = useState('upload');
  const [dragging, setDragging]     = useState(false);
  const [fileName, setFileName]     = useState('');
  const [rawHeaders, setRawHeaders] = useState([]);
  const [rawRows, setRawRows]       = useState([]);   // full original rows, all columns
  const [fieldMap, setFieldMap]     = useState({});
  const [columns, setColumns]       = useState([]);   // 12 resolved column names
  const [rows, setRows]             = useState([]);   // 12-column resolved rows
  const [inputs, setInputs]         = useState(DEFAULT_INPUTS);
  const [filterText, setFilterText] = useState('');
  const [yearRange, setYearRange]   = useState('2015-2027');
  const [maxMileage, setMaxMileage] = useState(99999);
  const [isLoading, setIsLoading]   = useState(false);
  const loadingTimer = useRef(null);
  const [bookKey, setBookKey]       = useState('WHOLESALE / TRADE-IN');
  const fileRef = useRef();

  const colIdx = {};
  columns.forEach((c, i) => { colIdx[c] = i; });

  const applyMapping = useCallback((map, headers, dataRows) => {
    const activeFields = FIELD_ALIASES.filter((f) => map[f.key] !== '');
    const resolvedCols = activeFields.map((f) => f.key);
    const resolvedRows = dataRows.map((row) =>
      activeFields.map((f) => {
        const ci = headers.indexOf(map[f.key]);
        return ci >= 0 ? (row[ci] ?? '') : '';
      })
    );
    setColumns(resolvedCols);
    setRows(resolvedRows);
    setRawRows(dataRows);   // <-- store full original rows for filter
    setStage('results');
  }, []);

  const parseFile = useCallback((file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (!data || data.length < 1) return;
      const headers  = data[0].map((h) => String(h));
      const dataRows = data.slice(1).filter((r) => r.some((c) => c !== ''));
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
      setRawHeaders(headers);
      setRawRows(dataRows);
      setFieldMap(initialMap);
      setFileName(file.name);
      if (Object.values(initialMap).some((v) => v === '')) setStage('mapping');
      else applyMapping(initialMap, headers, dataRows);
    };
    reader.readAsArrayBuffer(file);
  }, [applyMapping]);

  const confirmMapping = () => applyMapping(fieldMap, rawHeaders, rawRows);
  const onDragOver   = (e) => { e.preventDefault(); setDragging(true); };
  const onDragLeave  = () => setDragging(false);
  const onDrop       = (e) => { e.preventDefault(); setDragging(false); parseFile(e.dataTransfer.files[0]); };
  const onFileChange = (e) => parseFile(e.target.files[0]);
  const setInput     = (key, val) => setInputs((prev) => ({ ...prev, [key]: val }));

  const reset = () => {
    setStage('upload'); setColumns([]); setRows([]);
    setRawHeaders([]); setRawRows([]); setFieldMap({});
    setFileName(''); setInputs(DEFAULT_INPUTS);
    setFilterText(''); setYearRange('2015-2027'); setMaxMileage(99999); setBookKey('WHOLESALE / TRADE-IN'); setIsLoading(false); if (loadingTimer.current) clearTimeout(loadingTimer.current);
    if (fileRef.current) fileRef.current.value = '';
  };

  // ── Compute ranked deals (all valid, no slice yet) ───────────────────────────
  const rankedDeals = (() => {
    const maxPmt   = inputs.maxPayment !== '' ? parseFloat(inputs.maxPayment) : null;
    const hasBook  = colIdx[bookKey] !== undefined;
    const hasPrice = colIdx['PRICE'] !== undefined;
    if (!hasBook || !hasPrice) return [];
    return rows
      .map((row, i) => ({ row, i, ...calcDeal(row, colIdx, inputs, bookKey) }))
      .filter((d) => {
        if (d.book <= 0 || d.price <= 0 || d.equityPct === null) return false;
        if (maxPmt !== null && !isNaN(maxPmt) && d.payment > maxPmt) return false;
        return true;
      })
      .sort((a, b) => a.equityPct - b.equityPct);
  })();

  // ── Apply text filter AFTER ranking, walk until 10 matches ──────────────────
  // Parse "2020-2025" or "2022" into yfrom/yto
  const yearParts = yearRange.trim().split('-').map(s => parseInt(s.trim())).filter(n => !isNaN(n));
  const yfrom = yearParts[0] ?? null;
  const yto   = yearParts[1] ?? (yearParts.length === 1 ? yearParts[0] : null);
  const needle = filterText.trim().toLowerCase();

  // Alias expansion — each key also searches its paired terms
  const FILTER_ALIASES = { 'car': ['car', 'sedan'], 'suv': ['suv', 'sport'], 'truck': ['truck', 'cab'] };
  const needles = FILTER_ALIASES[needle] ?? (needle ? [needle] : []);

  const displayDeals = (() => {
    if (!needles.length && yfrom === null && yto === null && !maxMileage) return rankedDeals.slice(0, 10);
    // mileage-only fast path
    if (!needles.length && yfrom === null && yto === null && maxMileage) {
      return rankedDeals.filter(d => {
        const odo = parseFloat(String(d.row[colIdx['ODOMETER']] ?? '').replace(/[^0-9.]/g, '')) || 0;
        return odo <= maxMileage;
      }).slice(0, 10);
    }
    // If only year filter, no text needed — still walk the list
    if (!needles.length) {
      const yResults = [];
      for (const deal of rankedDeals) {
        if (yResults.length >= 10) break;
        const yr = parseInt(String(deal.row[colIdx['YEAR']] ?? '')) || 0;
        if (yfrom !== null && yr < yfrom) continue;
        if (yto   !== null && yr > yto)   continue;
        const odo = parseFloat(String(deal.row[colIdx['ODOMETER']] ?? '').replace(/[^0-9.]/g, '')) || 0;
        if (maxMileage && odo > maxMileage) continue;
        yResults.push(deal);
      }
      return yResults;
    }
    const results = [];
    for (const deal of rankedDeals) {
      if (results.length >= 10) break;
      const fullRow = rawRows[deal.i];
      if (!fullRow) continue;
      // Year filter
      if (yfrom !== null || yto !== null) {
        const yr = parseInt(String(deal.row[colIdx['YEAR']] ?? '')) || 0;
        if (yfrom !== null && yr < yfrom) continue;
        if (yto   !== null && yr > yto)   continue;
      }
      // Mileage filter
      const odo = parseFloat(String(deal.row[colIdx['ODOMETER']] ?? '').replace(/[^0-9.]/g, '')) || 0;
      if (maxMileage && odo > maxMileage) continue;
      const rowText = fullRow.map((cell) => String(cell)).join(' ').toLowerCase();
      if (needles.some((n) => rowText.includes(n))) results.push(deal);
    }
    return results;
  })();

  const hasMaxPmt = inputs.maxPayment !== '' && !isNaN(parseFloat(inputs.maxPayment));
  const netTrade  = inputs.tradeAllowance - inputs.tradePayoff;

  // ════════════════════════════════════════════════════════════════════════════
  // UPLOAD
  // ════════════════════════════════════════════════════════════════════════════
  if (stage === 'upload') {
    const bubbles = [
      { title: 'Dealer Tool',     bullets: ['Faster approvals', 'Subprime lending', 'Negative equity structures', 'Gross maximizing'] },
      { title: 'Privacy',         bullets: ['Nothing shared', 'Locally stored data', 'No accounts needed', 'Runs in your browser'] },
      { title: 'Reliable & Fast', bullets: ['Instant parsing', 'Alias-matched fields', 'Works offline', 'No setup required'] },
    ];
    return (
      <div className={styles.homepageShell}>
        <div className={styles.homepageBg} aria-hidden="true"><div className={styles.bgGlow} /></div>
        <div className={styles.homepageContent}>
          <div className={styles.homeLogoRow}>
            <div className={styles.homeLogo}><FileSpreadsheet size={16} /></div>
            <span className={styles.homeLogoText}>GotEmDone</span>
          </div>
          <h1 className={styles.homeTitle}>Built for the<br /><span className={styles.homeTitleAccent}>finance desk.</span></h1>
          <p className={styles.homeSubtitle}>Drop your inventory sheet and get structured deal data in seconds.</p>
          <div
            className={`${styles.homeDropZone} ${dragging ? styles.homeDropZoneActive : ''}`}
            onDragOver={onDragOver} onDragLeave={onDragLeave} onDrop={onDrop}
            onClick={() => fileRef.current?.click()}
            role="button" tabIndex={0} aria-label="Upload spreadsheet"
            onKeyDown={(e) => e.key === 'Enter' && fileRef.current?.click()}
          >
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={onFileChange} />
            <div className={styles.homeDropInner}>
              <div className={styles.homeDropIconWrap}><Upload size={22} /></div>
              <div>
                <p className={styles.homeDropMain}>Drop your inventory sheet</p>
                <p className={styles.homeDropSub}>.xlsx · .xls · .csv &nbsp;·&nbsp; click or drag</p>
              </div>
            </div>
          </div>
          <div className={styles.bubbleRow}>
            {bubbles.map((b) => (
              <div className={styles.bubble} key={b.title}>
                <p className={styles.bubbleTitle}>{b.title}</p>
                <ul className={styles.bubbleList}>
                  {b.bullets.map((item) => (
                    <li key={item} className={styles.bubbleItem}>
                      <span className={styles.bubbleDot} aria-hidden="true" />{item}
                    </li>
                  ))}
                </ul>
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  }

  // ════════════════════════════════════════════════════════════════════════════
  // MAPPING
  // ════════════════════════════════════════════════════════════════════════════
  if (stage === 'mapping') {
    const unresolvedFields = FIELD_ALIASES.filter((f) => (fieldMap[f.key] ?? '') === '');
    return (
      <div className={styles.shell}>
        <header className={styles.header}>
          <div className={styles.logo}><FileSpreadsheet size={18} /><span>GotEmDone</span></div>
          <div className={styles.headerMeta}>
            <button className={styles.resetBtn} onClick={reset}><RefreshCw size={13} />New file</button>
          </div>
        </header>
        <main className={styles.mappingMain}>
          <div className={styles.mappingHero}>
            <div className={styles.mappingHeroText}>
              <h2 className={styles.mappingTitle}>
                <AlertTriangle size={16} className={styles.warnIcon} />
                {unresolvedFields.length} field{unresolvedFields.length !== 1 ? 's' : ''} not found in sheet
              </h2>
              <p className={styles.mappingSub}>
                {FIELD_ALIASES.length - unresolvedFields.length} of {FIELD_ALIASES.length} matched automatically.
                Assign the missing ones, or leave as N/A to skip.
              </p>
            </div>
            <button className={styles.confirmBtn} onClick={confirmMapping}><Check size={14} />Confirm &amp; continue</button>
          </div>
          <div className={styles.mappingGrid}>
            {unresolvedFields.map((field) => (
              <div key={field.key} className={`${styles.mappingRow} ${styles.mappingRowUnresolved}`}>
                <div className={styles.mappingFieldKey}><AlertTriangle size={13} className={styles.warnIcon} />{field.key}</div>
                <ArrowRight size={13} className={styles.mappingArrow} />
                <select
                  className={`${styles.mappingSelect} ${styles.mappingSelectWarn}`}
                  value={fieldMap[field.key] ?? ''}
                  onChange={(e) => setFieldMap((prev) => ({ ...prev, [field.key]: e.target.value }))}
                >
                  <option value="">-- N/A (skip) --</option>
                  {rawHeaders.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            ))}
          </div>
          <div className={styles.mappingFooter}>
            <button className={styles.confirmBtn} onClick={confirmMapping}><Check size={14} />Confirm &amp; continue</button>
            <span className={styles.mappingFooterNote}>Fields set to N/A will not appear in results.</span>
          </div>
        </main>
      </div>
    );
  }

  // ════════════════════════════════════════════════════════════════════════════
  // RESULTS
  // ════════════════════════════════════════════════════════════════════════════
  return (
    <div className={styles.shell}>
      {/* ── Unified header ── */}
      <header className={styles.uniHeader}>
        {/* Top bar: logo + actions */}
        <div className={styles.uniTop}>
          <div className={styles.logo}><FileSpreadsheet size={16} /><span>GotEmDone</span></div>
          <div className={styles.uniActions}>
            <span className={styles.stripStat}>
              {needle
                ? <>{displayDeals.length} match{displayDeals.length !== 1 ? 'es' : ''} for "<strong style={{color:'var(--accent)'}}>{filterText.trim()}</strong>"</>
                : <>{rows.length.toLocaleString()} vehicles · {rankedDeals.length} rankable</>
              }
            </span>
            <button className={styles.printBtn} onClick={() => printDeals(displayDeals, inputs, bookKey, colIdx)} disabled={displayDeals.length === 0}>
              <Printer size={13} />Print PDF
            </button>
            <button className={styles.resetBtn} onClick={reset}><RefreshCw size={13} />New file</button>
          </div>
        </div>

        {/* Bottom bar: all inputs in one row */}
        <div className={styles.uniInputs}>
          {[
            { key: 'downPayment',    label: 'Down ($)'    },
            { key: 'tradeAllowance', label: 'Trade + ($)' },
            { key: 'tradePayoff',    label: 'Payoff ($)'  },
            { key: 'dealerAddendum', label: 'Addendum ($)'},
            { key: 'maxPayment',     label: 'Max Pmt ($)', placeholder: 'no limit' },
            { key: 'term',           label: 'Term (mo)'   },
            { key: 'rate',           label: 'Rate (%)'    },
          ].map(({ key, label, placeholder }) => (
            <div className={styles.uniField} key={key}>
              <label className={styles.uniLabel}>{label}</label>
              <input
                className={styles.uniInput}
                type="number"
                value={inputs[key]}
                placeholder={placeholder ?? '0'}
                onChange={(e) => {
                  const v = e.target.value;
                  setInput(key, key === 'maxPayment' ? v : (v === '' ? 0 : parseFloat(v) || 0));
                  setIsLoading(true);
                  if (loadingTimer.current) clearTimeout(loadingTimer.current);
                  loadingTimer.current = setTimeout(() => setIsLoading(false), Math.random() * 2000 + 1000);
                }}
              />
            </div>
          ))}

          {/* Year range */}
          <div className={styles.uniField}>
            <label className={styles.uniLabel}>Year</label>
            <input
              className={styles.uniTextInput}
              type="text"
              value={yearRange}
              placeholder="2020-2025"
              onChange={(e) => {
                setYearRange(e.target.value);
                setIsLoading(true);
                if (loadingTimer.current) clearTimeout(loadingTimer.current);
                loadingTimer.current = setTimeout(() => setIsLoading(false), Math.random() * 2000 + 1000);
              }}
            />
          </div>

          {/* Max mileage */}
          <div className={styles.uniField}>
            <label className={styles.uniLabel}>Max Miles</label>
            <input
              className={styles.uniInput}
              type="number"
              value={maxMileage}
              placeholder="99999"
              onChange={(e) => {
                const v = e.target.value;
                setMaxMileage(v === '' ? 0 : parseInt(v) || 0);
                setIsLoading(true);
                if (loadingTimer.current) clearTimeout(loadingTimer.current);
                loadingTimer.current = setTimeout(() => setIsLoading(false), Math.random() * 2000 + 1000);
              }}
            />
          </div>

          {/* Filter */}
          <div className={styles.uniField}>
            <label className={styles.uniLabel}>Filter</label>
            <input
              className={styles.uniTextInput}
              type="text"
              value={filterText}
              placeholder="tacoma, suv…"
              onChange={(e) => {
                setFilterText(e.target.value);
                setIsLoading(true);
                if (loadingTimer.current) clearTimeout(loadingTimer.current);
                loadingTimer.current = setTimeout(() => setIsLoading(false), Math.random() * 2000 + 1000);
              }}
            />
          </div>

          {/* Book toggle */}
          <div className={styles.uniField}>
            <label className={styles.uniLabel}>Book</label>
            <div className={styles.bookToggle}>
              <button className={`${styles.bookToggleBtn} ${bookKey === 'WHOLESALE / TRADE-IN' ? styles.bookToggleActive : ''}`} onClick={() => setBookKey('WHOLESALE / TRADE-IN')}>Wholesale</button>
              <button className={`${styles.bookToggleBtn} ${bookKey === 'RETAIL / MSRP' ? styles.bookToggleActive : ''}`} onClick={() => setBookKey('RETAIL / MSRP')}>Retail</button>
            </div>
          </div>

          {/* Reset */}
          <div className={styles.uniField} style={{justifyContent:'flex-end'}}>
            <label className={styles.uniLabel}>&nbsp;</label>
            <button className={styles.stripResetBtn}
              onClick={() => { setInputs(DEFAULT_INPUTS); setFilterText(''); setYearRange('2015-2027'); setMaxMileage(99999); setIsLoading(false); if (loadingTimer.current) clearTimeout(loadingTimer.current); }}
              title="Reset all inputs"><RotateCcw size={11} /></button>
          </div>
        </div>
      </header>

      {/* ── Results ── */}
      <main className={styles.resultsMain}>
        <div className={styles.resultsHeader}>
          <h2 className={styles.resultsTitle}>
            <TrendingUp size={16} className={styles.accentIcon} />
            Top {displayDeals.length} Deal{displayDeals.length !== 1 ? 's' : ''}
            <span className={styles.resultsTitleSub}> — ranked by equity % (lowest first)</span>
          </h2>
          {hasMaxPmt && (
            <span className={styles.filterPill}>
              <DollarSign size={11} />filtered ≤ {fmt$(parseFloat(inputs.maxPayment))}/mo
            </span>
          )}
        </div>

        <div className={styles.cardsWrap}>
          {isLoading && (
            <div className={styles.cardsOverlay}>
              <div className={styles.loadingSpinner} />
              <p className={styles.loadingText}>Calculating deals…</p>
            </div>
          )}
        {displayDeals.length === 0 ? (
          <div className={styles.noResults}>
            <AlertTriangle size={28} className={styles.warnIcon} />
            <p>{needle ? `No rankable deals match "${filterText.trim()}".` : 'No vehicles match your criteria.'}</p>
            <p style={{ fontSize: '12px', marginTop: '4px', color: 'var(--text-muted)' }}>
              {colIdx['PRICE'] === undefined || colIdx[bookKey] === undefined
                ? 'Price or book value column not found in this sheet.'
                : needle
                  ? 'The matching rows may not have valid price or book values.'
                  : 'Try adjusting your inputs above.'}
            </p>
          </div>
        ) : (
          <div className={styles.dealCards}>
            {displayDeals.map(({ row, i, price, book, amtFin, payment, equityPct, tax }, rank) => {
              const get = (key) => { const ci = colIdx[key]; return ci !== undefined ? row[ci] : ''; };
              const isNeg = equityPct !== null && equityPct > 100;
              const vehicle = [get('YEAR'), get('MAKE'), get('MODEL'), get('TRIM')].filter(Boolean).join(' ') || 'Unknown Vehicle';
              return (
                <div key={i} className={`${styles.dealCard} ${isNeg ? styles.dealCardNeg : styles.dealCardPos}`}>
                  {/* rank spans both rows */}
                  <div className={styles.dealRank}>#{rank + 1}</div>
                  {/* top row: name on its own line, pills + equity badge below */}
                  <div className={styles.dealIdentity}>
                    <p className={styles.dealVehicle}>{vehicle}</p>
                    <div className={styles.dealMeta}>
                      {get('STOCK')    && <span className={styles.dealMetaItem}>STK {get('STOCK')}</span>}
                      {get('COLOR')    && <span className={styles.dealMetaItem}>{get('COLOR')}</span>}
                      {get('ODOMETER') && <span className={styles.dealMetaItem}>{Number(get('ODOMETER')).toLocaleString()} mi</span>}
                      {get('AGE')      && <span className={styles.dealMetaItem}>{get('AGE')} days</span>}
                      <div className={`${styles.equityBadge} ${isNeg ? styles.equityBadgeNeg : styles.equityBadgePos}`}>
                        <span className={styles.equityPct}>{fmtPct(equityPct)}</span>
                        <span className={styles.equityLabel}>{isNeg ? 'neg' : 'equity'}</span>
                      </div>
                    </div>
                  </div>
                  {/* bottom row: financials */}
                  <div className={styles.dealFinancials}>
                    <div className={styles.dealFin}>
                      <span className={styles.dealFinLabel}>Price</span>
                      <span className={styles.dealFinVal}>{fmt$(price)}</span>
                    </div>
                    <div className={styles.dealFin}>
                      <span className={styles.dealFinLabel}>Tax + Reg</span>
                      <span className={styles.dealFinVal}>{fmt$(tax + REG_FEE)}</span>
                    </div>
                    <div className={styles.dealFin}>
                      <span className={styles.dealFinLabel}>{bookKey === 'WHOLESALE / TRADE-IN' ? 'Wholesale' : 'Retail'}</span>
                      <span className={styles.dealFinVal}>{fmt$(book)}</span>
                    </div>
                    <div className={styles.dealFin}>
                      <span className={styles.dealFinLabel}>Amt Financed</span>
                      <span className={styles.dealFinVal}>{fmt$(amtFin)}</span>
                    </div>
                    <div className={styles.dealFin}>
                      <span className={styles.dealFinLabel}>Est. Payment</span>
                      <span className={`${styles.dealFinVal} ${styles.dealFinPmt}`}>{fmt$(payment)}/mo</span>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
        </div>
      </main>
    </div>
  );
}
