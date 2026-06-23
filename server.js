const express = require('express');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_FILE = path.join(__dirname, 'data', 'responses.json');

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

if (!fs.existsSync(path.join(__dirname, 'data'))) fs.mkdirSync(path.join(__dirname, 'data'));
if (!fs.existsSync(DATA_FILE)) fs.writeFileSync(DATA_FILE, JSON.stringify([]));

function loadData() {
  try { return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')); } catch { return []; }
}
function saveData(data) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2));
}
function auth(req, res) {
  const token = req.query.token;
  if (token !== process.env.ADMIN_TOKEN && token !== 'admin123') {
    res.status(401).send('Nicht autorisiert'); return false;
  }
  return true;
}

// Antworten speichern (nach jeder Seite)
app.post('/api/save', (req, res) => {
  const { sessionId, step, answers, completed, language } = req.body;
  if (!sessionId) return res.status(400).json({ error: 'sessionId fehlt' });
  const data = loadData();
  const idx = data.findIndex(r => r.sessionId === sessionId);
  const entry = {
    sessionId,
    language: language || 'de',
    step: step || 0,
    answers: answers || {},
    completed: completed || false,
    lastUpdated: new Date().toISOString(),
    createdAt: idx === -1 ? new Date().toISOString() : data[idx].createdAt
  };
  if (idx === -1) data.push(entry); else data[idx] = entry;
  saveData(data);
  res.json({ ok: true });
});

// Admin: Testeinträge löschen
app.delete('/admin/clear', (req, res) => {
  if (!auth(req, res)) return;
  saveData([]);
  res.json({ ok: true, message: 'Alle Einträge gelöscht.' });
});

// Admin: einzelnen Eintrag per sessionId löschen
app.delete('/admin/delete/:sessionId', (req, res) => {
  if (!auth(req, res)) return;
  const data = loadData().filter(r => r.sessionId !== req.params.sessionId);
  saveData(data);
  res.json({ ok: true });
});

// Admin: Excel-Export
app.get('/admin/export', (req, res) => {
  if (!auth(req, res)) return;
  const data = loadData();
  const workbook = new ExcelJS.Workbook();

  // ---- Sheet 1: Alle Antworten ----
  const ws = workbook.addWorksheet('Alle Antworten');

  // Feld-IDs exakt wie in index_bilingual.html
  const labelMap = {
    email:      'E-Mail (Gewinnspiel)',
    size:       'Unternehmensgröße',
    role:       'Funktion',
    sector:     'Branche',
    ai:         'KI-Einsatz',
    know:       'Kenntnisstand EU AI Act',
    riskclass:  'Risikoklasse bekannt',
    compliance: 'Compliance-Status',
    bbawar:     'Black-Box-Bewusstsein',
    xai:        'Genutzte XAI-Methoden',
    barrier:    'Größte Hürde',
    support:    'Gewünschte Unterstützung',
    feas:       'Technische Umsetzbarkeit',
    opinion:    'Bewertung EU AI Act',
    comment:    'Offener Kommentar'
  };

  // Likert-Items: Keys wie in A[qid][item] gespeichert
  const likertDef = {
    lk1: [
      'Die Intransparenz unserer KI-Systeme erschwert die Erfüllung der EU AI Act-Anforderungen erheblich.',
      'Verfügbare XAI-Methoden reichen aus, um die Transparenzanforderungen des EU AI Acts zu erfüllen.',
      'Das Black-Box-Problem beeinträchtigt das Vertrauen unserer Kunden und Stakeholder in unsere KI-Systeme.',
      'Wir priorisieren Modellleistung gegenüber Interpretierbarkeit, wenn beides nicht gleichzeitig erreichbar ist.'
    ],
    lk2: [
      'Fehlende technische Normen für Erklärbarkeit (z.\u00a0B. ISO/IEC, CEN)',
      'Unklare oder interpretationsoffene rechtliche Anforderungen im Gesetzestext',
      'Mangelndes internes Fachwissen zu KI-Regulierung und Compliance',
      'Hoher Aufwand für technische Dokumentation und Konformitätsbewertung',
      'Fehlende Software-Tools zur Unterstützung der Compliance',
      'Wirtschaftliche Kosten der Compliance im Verhältnis zum Nutzen',
      'Fehlende Ressourcen (Personal, Budget, Zeit) für die Umsetzung'
    ],
    lk3: [
      'Erklärbare KI (XAI) wird in den nächsten 3 Jahren zum Industriestandard in unserer Branche.',
      'Der EU AI Act wird langfristig die Qualität und Sicherheit von KI-Systemen in der EU verbessern.',
      'Unternehmen ohne EU AI Act-Compliance werden erhebliche Wettbewerbsnachteile haben.',
      'Die Anforderungen des EU AI Acts werden in 5 Jahren technisch vollständig erfüllbar sein.',
      'Wir sind bereit, auf Modellleistung zugunsten besserer Interpretierbarkeit zu verzichten, wenn beides nicht gleichzeitig erreichbar ist.'
    ]
  };
  const likertShort = {
    lk1: ['lk1: Intransparenz erschwert EU AI Act','lk1: XAI-Methoden reichen aus','lk1: Black-Box mindert Vertrauen','lk1: Priorisierung Leistung > Interpretierbarkeit'],
    lk2: ['lk2: Fehlende tech. Normen','lk2: Unklare rechtl. Anforderungen','lk2: Mangelndes Fachwissen','lk2: Hoher Dokumentationsaufwand','lk2: Fehlende Compliance-Tools','lk2: Wirtschaftl. Kosten','lk2: Fehlende Ressourcen'],
    lk3: ['lk3: XAI wird Standard 3J','lk3: EU AI Act verbessert KI-Qualität','lk3: Nicht-Compliance = Wettbewerbsnachteil','lk3: Anforderungen in 5J erfüllbar','lk3: [KONTROLLFRAGE] Bereit Leistung für Interpretierbarkeit opfern']
  };

  const headers = ['Session-ID','Sprache','Erstellt','Abgeschlossen','Schritt'];
  Object.values(labelMap).forEach(l => headers.push(l));
  Object.values(likertShort).forEach(arr => arr.forEach(l => headers.push(l)));
  ws.addRow(headers);
  ws.getRow(1).font = { bold: true };
  ws.getRow(1).fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFEEEDFE' } };

  data.forEach(r => {
    const a = r.answers || {};
    const row = [
      r.sessionId,
      r.language || 'de',
      r.createdAt ? new Date(r.createdAt).toLocaleString('de-DE') : '',
      r.completed ? 'Ja' : 'Nein',
      r.step
    ];
    Object.keys(labelMap).forEach(k => {
      const v = a[k];
      row.push(Array.isArray(v) ? v.join('; ') : (v || ''));
    });
    Object.keys(likertDef).forEach(lkId => {
      const obj = a[lkId] || {};
      likertDef[lkId].forEach(item => row.push(obj[item] || ''));
    });
    ws.addRow(row);
  });
  ws.columns.forEach(c => { c.width = 32; });

  // ---- Sheet 2: Zusammenfassung ----
  const ws2 = workbook.addWorksheet('Zusammenfassung');
  const completed = data.filter(r => r.completed);
  ws2.addRow(['Kennzahl','Wert']);
  ws2.getRow(1).font = { bold: true };
  ws2.addRow(['Gesamt Einträge', data.length]);
  ws2.addRow(['Davon abgeschlossen', completed.length]);
  ws2.addRow(['Abbruchquote', data.length ? Math.round((1 - completed.length/data.length)*100)+'%' : '-']);
  ws2.addRow(['Export erstellt', new Date().toLocaleString('de-DE')]);
  ws2.columns = [{ width: 36 }, { width: 20 }];

  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition','attachment; filename=umfrage_export_'+new Date().toISOString().slice(0,10)+'.xlsx');
  workbook.xlsx.write(res).then(() => res.end());
});

// Admin: JSON-Rohdaten
app.get('/admin/data', (req, res) => {
  if (!auth(req, res)) return;
  res.json(loadData());
});

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log('Server: http://localhost:' + PORT);
  console.log('Export: http://localhost:' + PORT + '/admin/export?token=admin123');
  console.log('Löschen: DELETE http://localhost:' + PORT + '/admin/clear?token=admin123');
});
