const fs     = require('fs');
const path   = require('path');
const XLSX   = require('xlsx');
const bcrypt = require('bcryptjs');

const DB_PATH    = path.join(__dirname, 'db.json');
const EXCEL_PATH = path.join(__dirname, '../files/Sample Style details for Proposal.xlsx');

const LOOKUP_KEYS = ['Style #', 'Style', 'STYLE', 'Item', 'Item #', 'Bar Code', 'BARCODE'];

const DEFAULT_PERMISSIONS = {
  submit_proposals:   true,
  view_own_proposals: true,
  view_all_proposals: false,
  bulk_import:        false,
  view_catalog:       true,
  view_pricing:       false,
};

// ── Persistence ───────────────────────────────────────────────────────────────

function load() {
  if (fs.existsSync(DB_PATH)) {
    try { return JSON.parse(fs.readFileSync(DB_PATH, 'utf8')); } catch { /* fall through */ }
  }
  return { credentials: null, settings: {}, stock: [], columns: {}, proposals: [], sales_reps: [], activity_log: [], history: [] };
}

function save(db) {
  fs.writeFileSync(DB_PATH, JSON.stringify(db, null, 2), 'utf8');
}

// ── Init ──────────────────────────────────────────────────────────────────────

async function init() {
  const db = load();

  if (!db.credentials) {
    db.credentials = { username: 'admin', password_hash: await bcrypt.hash('jewelry123', 10) };
    console.log('✓ Default admin credentials set  (admin / jewelry123)');
  }

  if ((!db.stock || !db.stock.length) && fs.existsSync(EXCEL_PATH)) {
    const wb   = XLSX.readFile(EXCEL_PATH);
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    db.stock   = rows.map(r => { if (r['Bar Code'] !== undefined) r['Bar Code'] = String(r['Bar Code']); return r; });
    if (rows.length) Object.keys(rows[0]).forEach(c => { if (!(c in (db.columns || {}))) db.columns[c] = true; });
    console.log(`✓ Seeded ${db.stock.length} stock items from Excel`);
  }

  db.proposals    = db.proposals    || [];
  db.settings     = db.settings     || {};
  db.sales_reps   = db.sales_reps   || [];
  db.activity_log = db.activity_log || [];
  db.history      = db.history      || [];

  // One-time: seed history from existing proposals
  if (!db.history_seeded) {
    const seeded = [];
    for (const p of db.proposals) {
      seeded.push({ id: `hist_seed_${p.id}_created`, type: 'proposal_created', timestamp: p.created_at,
        actor_type: p.rep_id ? 'rep' : 'public', actor_id: p.rep_id || null, actor_name: p.salesperson_name,
        proposal_id: p.id, client_company: p.client_company, rep_id: p.rep_id || null, rep_name: p.salesperson_name,
        details: `Proposal created for ${p.client_company} (${(p.items||[]).length} items)`, proposal_status: p.status });
      if (p.status === 'approved' || p.status === 'denied') {
        seeded.push({ id: `hist_seed_${p.id}_${p.status}`, type: `proposal_${p.status}`, timestamp: p.created_at,
          actor_type: 'admin', actor_id: null, actor_name: 'Admin',
          proposal_id: p.id, client_company: p.client_company, rep_id: p.rep_id || null, rep_name: p.salesperson_name,
          details: `Proposal ${p.status} for ${p.client_company}`, proposal_status: p.status });
      }
    }
    seeded.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    db.history = [...seeded, ...db.history];
    db.history_seeded = true;
  }

  save(db);
  console.log('✓ Store ready →', DB_PATH);
}

// ── Credentials ───────────────────────────────────────────────────────────────

function getCredentials()                  { return load().credentials; }
function updateCredentials(username, hash) { const db = load(); db.credentials = { username, password_hash: hash }; save(db); }

// ── Settings ──────────────────────────────────────────────────────────────────

function getAllSettings()       { return load().settings || {}; }
function getSetting(key)        { return (load().settings || {})[key] ?? null; }
function setSetting(key, value) { const db = load(); db.settings = db.settings || {}; db.settings[key] = value; save(db); }
function deleteSetting(key)     { const db = load(); if (db.settings) delete db.settings[key]; save(db); }

// ── Stock ─────────────────────────────────────────────────────────────────────

function getStock() { return load().stock || []; }

function setStock(rows) {
  const db = load();
  db.stock = rows.map(r => { if (r['Bar Code'] !== undefined) r['Bar Code'] = String(r['Bar Code']); return r; });
  const cols = rows.length ? Object.keys(rows[0]) : [];
  cols.forEach(c => { if (!(c in (db.columns || {}))) db.columns[c] = true; });
  save(db);
  return cols;
}

function lookupItem(num) {
  const needle = String(num).toLowerCase().trim();
  return load().stock.find(s => LOOKUP_KEYS.some(k => String(s[k] ?? '').toLowerCase().trim() === needle)) ?? null;
}

// ── Column settings ───────────────────────────────────────────────────────────

function getColumns()         { return Object.entries(load().columns || {}).map(([column_name, is_active]) => ({ column_name, is_active: Boolean(is_active) })); }
function saveColumns(columns) { const db = load(); db.columns = {}; columns.forEach(c => { db.columns[c.column_name] = c.is_active; }); save(db); }

// ── Proposals ─────────────────────────────────────────────────────────────────

function getProposals(statusFilter, repId) {
  let list = (load().proposals || []).slice().reverse();
  if (statusFilter && statusFilter !== 'all') list = list.filter(p => p.status === statusFilter);
  if (repId) list = list.filter(p => p.rep_id === repId);
  return list;
}

function getProposal(id) { return (load().proposals || []).find(p => p.id === id) ?? null; }

function createProposal({ salesperson_name, salesperson_email, client_company, client_email, client_phone, notes, items, rep_id }) {
  if (!salesperson_name || !salesperson_email || !client_company || !items?.length)
    throw Object.assign(new Error('salesperson_name, salesperson_email, client_company, and items are required'), { status: 400 });
  const db = load();
  const proposal = {
    id: 'JP' + Date.now(), rep_id: rep_id || null,
    salesperson_name, salesperson_email, client_company,
    client_email: client_email || null, client_phone: client_phone || null, notes: notes || null,
    items: [...items], status: 'pending', created_at: new Date().toISOString(),
  };
  db.proposals.push(proposal);
  save(db);
  return proposal;
}

function updateProposalStatus(id, status) {
  if (!['pending', 'approved', 'denied'].includes(status)) throw Object.assign(new Error('Invalid status'), { status: 400 });
  const db  = load();
  const idx = (db.proposals || []).findIndex(p => p.id === id);
  if (idx === -1) throw Object.assign(new Error('Not found'), { status: 404 });
  db.proposals[idx].status = status;
  save(db);
}

function getStats() {
  const p = load().proposals || [];
  return { total: String(p.length), pending: String(p.filter(x => x.status==='pending').length), approved: String(p.filter(x => x.status==='approved').length), denied: String(p.filter(x => x.status==='denied').length) };
}

// ── Sales reps ────────────────────────────────────────────────────────────────

function getSalesReps()            { return load().sales_reps || []; }
function getSalesRep(id)           { return getSalesReps().find(r => r.id === id) ?? null; }
function getSalesRepByEmail(email) { return getSalesReps().find(r => r.email.toLowerCase() === email.toLowerCase()) ?? null; }

function createSalesRep({ name, email, password_hash, permissions }) {
  const db  = load();
  if (!db.sales_reps) db.sales_reps = [];
  const rep = { id: 'sr_' + Date.now(), name, email, password_hash, permissions: { ...DEFAULT_PERMISSIONS, ...permissions }, active: true, created_at: new Date().toISOString() };
  db.sales_reps.push(rep);
  save(db);
  return rep;
}

function updateSalesRep(id, updates) {
  const db  = load();
  const idx = (db.sales_reps || []).findIndex(r => r.id === id);
  if (idx === -1) throw Object.assign(new Error('Sales rep not found'), { status: 404 });
  const { password_hash, ...safe } = updates;
  db.sales_reps[idx] = { ...db.sales_reps[idx], ...safe };
  save(db);
  return db.sales_reps[idx];
}

function updateRepPassword(id, password_hash) {
  const db  = load();
  const idx = (db.sales_reps || []).findIndex(r => r.id === id);
  if (idx === -1) throw Object.assign(new Error('Sales rep not found'), { status: 404 });
  db.sales_reps[idx].password_hash = password_hash;
  save(db);
}

// ── Activity log ──────────────────────────────────────────────────────────────

function logActivity({ rep_id, rep_name, action, detail }) {
  const db = load();
  if (!db.activity_log) db.activity_log = [];
  db.activity_log.unshift({ id: 'log_' + Date.now(), rep_id, rep_name, action, detail, timestamp: new Date().toISOString() });
  if (db.activity_log.length > 1000) db.activity_log = db.activity_log.slice(0, 1000);
  save(db);
}

function getActivityLog(limit = 100) { return (load().activity_log || []).slice(0, limit); }

// ── History ───────────────────────────────────────────────────────────────────

function logHistory(entry) {
  const db = load();
  if (!db.history) db.history = [];
  db.history.unshift({ id: 'hist_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6), timestamp: new Date().toISOString(), ...entry });
  if (db.history.length > 10000) db.history = db.history.slice(0, 10000);
  save(db);
}

function getHistory({ page = 1, limit = 50, search = '', startDate, endDate, repId, repOnly, types, sort = 'desc' } = {}) {
  let list = load().history || [];

  if (repOnly)                          list = list.filter(h => h.rep_id === repOnly || h.actor_id === repOnly);
  if (repId && repId !== 'all')         list = list.filter(h => h.rep_id === repId   || h.actor_id === repId);
  if (search && search.trim()) {
    const q = search.toLowerCase().trim();
    list = list.filter(h => [(h.client_company||''), (h.rep_name||''), (h.actor_name||''), (h.proposal_id||''), (h.details||'')].some(v => v.toLowerCase().includes(q)));
  }
  if (startDate) { const from = new Date(startDate).setHours(0,0,0,0); list = list.filter(h => new Date(h.timestamp).getTime() >= from); }
  if (endDate)   { const to   = new Date(endDate).setHours(23,59,59,999); list = list.filter(h => new Date(h.timestamp).getTime() <= to); }
  if (types && types.length)            list = list.filter(h => types.includes(h.type));
  if (sort === 'asc')                   list = list.slice().sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));

  const total  = list.length;
  const pages  = Math.max(1, Math.ceil(total / limit));
  const offset = (Math.max(1, page) - 1) * limit;
  return { data: list.slice(offset, offset + limit), total, page: Math.max(1, page), pages, limit };
}

module.exports = {
  init, DEFAULT_PERMISSIONS,
  getCredentials, updateCredentials,
  getAllSettings, getSetting, setSetting, deleteSetting,
  getStock, setStock, lookupItem,
  getColumns, saveColumns,
  getProposals, getProposal, createProposal, updateProposalStatus, getStats,
  getSalesReps, getSalesRep, getSalesRepByEmail, createSalesRep, updateSalesRep, updateRepPassword,
  logActivity, getActivityLog,
  logHistory, getHistory,
};
