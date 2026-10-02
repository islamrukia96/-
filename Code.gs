// Google Sheet এর Extensions > Apps Script এ এটা পেস্ট করুন।
// পাসওয়ার্ড: Project Settings > Script properties > ADMIN_PASS
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const sh = n => { const s = SpreadsheetApp.getActiveSpreadsheet(); return s.getSheetByName(n) || s.insertSheet(n); };

function doPost(e) {
  const r = JSON.parse(e.postData.contents);
  const P = PropertiesService.getScriptProperties();
  const users = JSON.parse(P.getProperty('USERS') || '{}');
  const admin = r.user === 'admin' && !!r.pass && r.pass === P.getProperty('ADMIN_PASS');
  const staff = !!r.user && users.hasOwnProperty(r.user) && users[r.user] === r.pass;
  if (!admin && !staff) return out({ error: 'wrong' });
  if (r.action === 'save') { write(r.data); return out({ ok: true }); }
  if (r.action === 'users' || r.action === 'adduser' || r.action === 'deluser') {
    if (!admin) return out({ error: 'শুধু অ্যাডমিন' });
    const d = r.data || {};
    if (r.action === 'adduser') {
      const n = String(d.name || '').trim();
      if (!n || n === 'admin' || !d.pass) return out({ error: 'নাম/পাসওয়ার্ড ঠিক নয়' });
      if (!users.hasOwnProperty(n) && Object.keys(users).length >= 10) return out({ error: 'সর্বোচ্চ ১০ জন ইউজার' });
      users[n] = String(d.pass);
    }
    if (r.action === 'deluser') delete users[d.name];
    P.setProperty('USERS', JSON.stringify(users));
    return out({ users: Object.keys(users) });
  }
  return out({ data: read(), role: admin ? 'admin' : 'staff' });
}

function put(n, head, rows) {
  const s = sh(n); s.clear();
  const all = [head].concat(rows);
  const g = s.getRange(1, 1, all.length, head.length);
  g.setNumberFormat('@'); g.setValues(all);
}

function write(d) {
  put('Products', ['id', 'name', 'buy', 'price', 'stock'], d.products.map(p => [p.id, p.name, p.buy, p.price, p.stock]));
  put('Customers', ['id', 'name', 'phone'], d.customers.map(c => [c.id, c.name, c.phone]));
  put('Txs', ['id', 'type', 'date', 'cid', 'customer', 'items', 'total', 'paid', 'amount', 'cost', 'by'],
    d.txs.map(t => [t.id, t.type, t.date, t.cid || '', t.cname || '', JSON.stringify(t.items || []), t.total || 0, t.paid || 0, t.amount || 0, t.cost || 0, t.by || '']));
}

function read() {
  const g = n => sh(n).getDataRange().getValues().slice(1).filter(r => r[0] !== '');
  return {
    products: g('Products').map(r => ({ id: String(r[0]), name: String(r[1]), buy: +r[2], price: +r[3], stock: +r[4] })),
    customers: g('Customers').map(r => ({ id: String(r[0]), name: String(r[1]), phone: String(r[2]) })),
    txs: g('Txs').map(r => ({ id: String(r[0]), type: String(r[1]), date: String(r[2]), cid: String(r[3]), cname: String(r[4]),
      items: JSON.parse(r[5] || '[]'), total: +r[6], paid: +r[7], amount: +r[8], cost: +r[9], by: String(r[10] || '') }))
  };
}
