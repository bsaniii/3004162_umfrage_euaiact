const express = require('express');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_FILE = path.join(__dirname, 'data', 'responses.json');

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

if (!fs.existsSync(path.join(__dirname, 'data'))) {
  fs.mkdirSync(path.join(__dirname, 'data'));
}
if (!fs.existsSync(DATA_FILE)) {
  fs.writeFileSync(DATA_FILE, JSON.stringify([]));
}

function load() {
  try { return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')); }
  catch { return []; }
}
function save(data) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2));
}

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/api/save', (req, res) => {
  const { sessionId, step, answers, completed } = req.body;
  if (!sessionId) return res.status(400).json({ error: 'Kein sessionId' });
  const data = load();
  const idx = data.findIndex(r => r.sessionId === sessionId);
  const entry = {
    sessionId,
    step: step || 0,
    answers: answers || {},
    completed: !!completed,
    lastUpdated: new Date().toISOString(),
    createdAt: idx === -1 ? new Date().toISOString() : data[idx].createdAt
  };
  if (idx === -1) data.push(entry);
  else data[idx] = entry;
  save(data);
  res.json({ ok: true });
});

app.get('/admin/export', async (req, res) => {
  if (req.query.token !== (process.env.ADMIN_TOKEN || 'admin123')) {
    return res.status(401).send('Nicht autorisiert');
  }
  const data = load();
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Antworten');
  const ws2 = wb.addWorksheet('Zusammenfassung');

  const felder = [
    ['email','E-Mail (Gewinnspiel)'],
    ['size','Unternehmensgröße'],
    ['role','Funktion'],
    ['sector','Branche'],
    ['ai','KI-Einsatz'],
    ['know','Kenntnisstand EU AI Act'],
    ['riskclass','Risikoklassifizierung bekannt'],
    ['compliance','Compliance-Status'],
    ['bbawar','Black-Box-Bewusstsein'],
    ['xai','XAI-Methoden'],
    ['barrier','Größte Hürde'],
    ['support','Gewünschte Unterstützung'],
    ['feas','Technische Umsetzbarkeit'],
    ['opinion','Bewertung EU AI Act'],
    ['comment','Kommentar']
  ];

  const likert = {
    lk1: ['Intransparenz erschwert EU AI Act','XAI-Methoden reichen aus','Black-Box mindert Kundenvertrauen','Bereit Leistung zu opfern'],
    lk2: ['Fehlende Standards','Unklare Anforderungen','Fehlendes Fachwissen','Hoher Dokumentationsaufwand','Fehlende Tools','Wirtschaftliche Kosten'],
    lk3: ['XAI wird Standard in 3 Jahren','EU AI Act verbessert KI-Qualität','Wettbewerbsnachteil ohne Compliance']
  };

  const headers = ['Session-ID','Erstellt','Aktualisiert','Schritt','Abgeschlossen'];
  felder.forEach(([,l]) => headers.push(l));
  Object.values(likert).forEach(items => items.forEach(i => headers.push(i)));

  ws.addRow(headers);
  ws.getRow(1).font = { bold: true };
  ws.getRow(1).fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFEEEDFE' }};

  data.forEach(r => {
    const a = r.answers || {};
    const row = [
      r.sessionId,
      r.createdAt ? new Date(r.createdAt).toLocaleString('de-DE') : '',
      r.lastUpdated ? new Date(r.lastUpdated).toLocaleString('de-DE') : '',
      r.step,
      r.completed ? 'Ja' : 'Nein'
    ];
    felder.forEach(([k]) => {
      const v = a[k];
      row.push(Array.isArray(v) ? v.join('; ') : (v || ''));
    });
    Object.keys(likert).forEach(k => {
      const obj = a[k] || {};
      likert[k].forEach(item => row.push(obj[item] || ''));
    });
    ws.addRow(row);
  });
  ws.columns.forEach(c => { c.width = 30; });

  const done = data.filter(r => r.completed);
  ws2.addRow(['Kennzahl','Wert']);
  ws2.getRow(1).font = { bold: true };
  ws2.addRow(['Gesamt Teilnahmen', data.length]);
  ws2.addRow(['Abgeschlossen', done.length]);
  ws2.addRow(['Abgebrochen', data.length - done.length]);
  ws2.addRow(['Export erstellt', new Date().toLocaleString('de-DE')]);
  ws2.columns = [{ width: 30 }, { width: 20 }];

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=umfrage_' + new Date().toISOString().slice(0,10) + '.xlsx');
  await wb.xlsx.write(res);
  res.end();
});

app.listen(PORT, () => {
  console.log('Server läuft auf Port ' + PORT);
});
