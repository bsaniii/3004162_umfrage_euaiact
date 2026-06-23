const express = require('express');
const path = require('path');
const { MongoClient } = require('mongodb');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;
const MONGO_URI = process.env.MONGO_URI;
const ADMIN_TOKEN = process.env.ADMIN_TOKEN || 'admin123';

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

let db;

async function connectDB() {
  const client = new MongoClient(MONGO_URI);
  await client.connect();
  db = client.db('umfrage');
  console.log('MongoDB verbunden');
}

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Antworten speichern (nach jeder Seite)
app.post('/api/save', async (req, res) => {
  const { sessionId, step, answers, completed, language } = req.body;
  if (!sessionId) return res.status(400).json({ error: 'Kein sessionId' });
  try {
    const col = db.collection('responses');
    const existing = await col.findOne({ sessionId });
    const entry = {
      sessionId,
      language: language || 'de',
      step: step || 0,
      answers: answers || {},
      completed: !!completed,
      lastUpdated: new Date(),
      createdAt: existing ? existing.createdAt : new Date()
    };
    await col.replaceOne({ sessionId }, entry, { upsert: true });
    res.json({ ok: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Datenbankfehler' });
  }
});

function auth(req, res) {
  if (req.query.token !== ADMIN_TOKEN) {
    res.status(401).send('Nicht autorisiert');
    return false;
  }
  return true;
}

// Admin: alle Einträge löschen (für Testdaten)
app.delete('/admin/clear', async (req, res) => {
  if (!auth(req, res)) return;
  try {
    await db.collection('responses').deleteMany({});
    res.json({ ok: true, message: 'Alle Einträge gelöscht.' });
  } catch (e) {
    res.status(500).json({ error: 'Datenbankfehler' });
  }
});

// Admin: einzelnen Eintrag löschen
app.delete('/admin/delete/:sessionId', async (req, res) => {
  if (!auth(req, res)) return;
  try {
    await db.collection('responses').deleteOne({ sessionId: req.params.sessionId });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: 'Datenbankfehler' });
  }
});

// Admin: JSON-Rohdaten
app.get('/admin/data', async (req, res) => {
  if (!auth(req, res)) return;
  try {
    const data = await db.collection('responses').find({}).toArray();
    res.json(data);
  } catch (e) {
    res.status(500).json({ error: 'Datenbankfehler' });
  }
});

// Admin: Excel-Export
app.get('/admin/export', async (req, res) => {
  if (!auth(req, res)) return;
  try {
    const data = await db.collection('responses').find({}).toArray();
    const wb = new ExcelJS.Workbook();

    // Sheet 1: Alle Antworten
    const ws = wb.addWorksheet('Alle Antworten');

    const felder = [
      ['email',      'E-Mail (Gewinnspiel)'],
      ['size',       'Unternehmensgröße'],
      ['role',       'Funktion'],
      ['role_other', 'Funktion (Sonstige)'],
      ['sector',     'Branche'],
      ['sector_other','Branche (Sonstige)'],
      ['ai',         'KI-Einsatz'],
      ['know',       'Kenntnisstand EU AI Act'],
      ['riskclass',  'Risikoklasse bekannt'],
      ['compliance', 'Compliance-Status'],
      ['bbawar',     'Black-Box-Bewusstsein'],
      ['xai',        'Genutzte XAI-Methoden'],
      ['xai_other',  'XAI-Methoden (Sonstige)'],
      ['barrier',    'Größte Hürde'],
      ['support',    'Gewünschte Unterstützung'],
      ['feas',       'Technische Umsetzbarkeit'],
      ['opinion',    'Bewertung EU AI Act'],
      ['comment',    'Offener Kommentar']
    ];

    // Likert-Items exakt wie in index.html gespeichert
    const likert = {
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

    const likertKurz = {
      lk1: ['lk1: Intransparenz erschwert EU AI Act','lk1: XAI-Methoden reichen aus','lk1: Black-Box mindert Vertrauen','lk1: Priorisierung Leistung > Interpretierbarkeit'],
      lk2: ['lk2: Fehlende Standards','lk2: Unklare Anforderungen','lk2: Fehlendes Fachwissen','lk2: Hoher Dokumentationsaufwand','lk2: Fehlende Tools','lk2: Wirtschaftl. Kosten','lk2: Fehlende Ressourcen'],
      lk3: ['lk3: XAI wird Standard 3J','lk3: EU AI Act verbessert KI','lk3: Wettbewerbsnachteil ohne Compliance','lk3: Anforderungen in 5J erfüllbar','lk3: [KONTROLLFRAGE] Bereit Leistung für Interpretierbarkeit opfern']
    };

    const headers = ['Session-ID','Sprache','Erstellt','Abgeschlossen','Schritt'];
    felder.forEach(([,l]) => headers.push(l));
    Object.values(likertKurz).forEach(arr => arr.forEach(l => headers.push(l)));

    ws.addRow(headers);
    ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F2D5E' } };

    data.forEach(r => {
      const a = r.answers || {};
      const row = [
        r.sessionId,
        r.language || 'de',
        r.createdAt ? new Date(r.createdAt).toLocaleString('de-DE') : '',
        r.completed ? 'Ja' : 'Nein',
        r.step
      ];
      felder.forEach(([k]) => {
        const v = a[k];
        row.push(Array.isArray(v) ? v.join('; ') : (v || ''));
      });
      Object.keys(likert).forEach(lkId => {
        const obj = a[lkId] || {};
        likert[lkId].forEach(item => row.push(obj[item] || ''));
      });
      ws.addRow(row);
    });
    ws.columns.forEach(c => { c.width = 32; });

    // Sheet 2: Zusammenfassung
    const ws2 = wb.addWorksheet('Zusammenfassung');
    const done = data.filter(r => r.completed);
    ws2.addRow(['Kennzahl', 'Wert']);
    ws2.getRow(1).font = { bold: true };
    ws2.addRow(['Gesamt Einträge', data.length]);
    ws2.addRow(['Davon abgeschlossen', done.length]);
    ws2.addRow(['Abbruchquote', data.length ? Math.round((1 - done.length / data.length) * 100) + '%' : '-']);
    ws2.addRow(['Export erstellt', new Date().toLocaleString('de-DE')]);
    ws2.columns = [{ width: 36 }, { width: 20 }];

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=umfrage_' + new Date().toISOString().slice(0, 10) + '.xlsx');
    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    console.error(e);
    res.status(500).send('Fehler beim Export');
  }
});

connectDB().then(() => {
  app.listen(PORT, () => console.log('Server läuft auf Port ' + PORT));
}).catch(err => {
  console.error('MongoDB Verbindungsfehler:', err);
  process.exit(1);
});
