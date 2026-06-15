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

app.post('/api/save', async (req, res) => {
  const { sessionId, step, answers, completed } = req.body;
  if (!sessionId) return res.status(400).json({ error: 'Kein sessionId' });
  try {
    const col = db.collection('responses');
    const existing = await col.findOne({ sessionId });
    const entry = {
      sessionId,
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

app.get('/admin/data', async (req, res) => {
  if (req.query.token !== ADMIN_TOKEN) return res.status(401).send('Nicht autorisiert');
  try {
    const data = await db.collection('responses').find({}).toArray();
    res.json(data);
  } catch (e) {
    res.status(500).json({ error: 'Datenbankfehler' });
  }
});

app.get('/admin/export', async (req, res) => {
  if (req.query.token !== ADMIN_TOKEN) return res.status(401).send('Nicht autorisiert');
  try {
    const data = await db.collection('responses').find({}).toArray();
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Antworten');
    const ws2 = wb.addWorksheet('Zusammenfassung');

    const felder = [
      ['email','E-Mail (Gewinnspiel)'],
      ['size','Unternehmensgroesse'],
      ['role','Funktion'],
      ['role_other','Funktion (Sonstige)'],
      ['sector','Branche'],
      ['sector_other','Branche (Sonstige)'],
      ['ai','KI-Einsatz'],
      ['know','Kenntnisstand EU AI Act'],
      ['riskclass','Risikoklassifizierung'],
      ['compliance','Compliance-Status'],
      ['bbawar','Black-Box-Bewusstsein'],
      ['xai','XAI-Methoden'],
      ['xai_other','XAI-Methoden (Sonstige)'],
      ['barrier','Groesste Hurde'],
      ['support','Gewuenschte Unterstuetzung'],
      ['feas','Technische Umsetzbarkeit'],
      ['opinion','Bewertung EU AI Act'],
      ['comment','Kommentar']
    ];

    const likert = {
      lk1: [
        'Die Intransparenz unserer KI-Systeme erschwert die Erfullung der EU AI Act-Anforderungen erheblich.',
        'Verfugbare XAI-Methoden reichen aus, um die Transparenzanforderungen des EU AI Acts zu erfullen.',
        'Das Black-Box-Problem beeintrachtigt das Vertrauen unserer Kunden und Stakeholder in unsere KI-Systeme.',
        'Wir priorisieren Modellleistung gegenuber Interpretierbarkeit, wenn beides nicht gleichzeitig erreichbar ist.'
      ],
      lk2: [
        'Fehlende technische Normen fur Erklarbarkeit (z.B. ISO/IEC, CEN)',
        'Unklare oder interpretationsoffene rechtliche Anforderungen im Gesetzestext',
        'Mangelndes internes Fachwissen zu KI-Regulierung und Compliance',
        'Hoher Aufwand fur technische Dokumentation und Konformitatsbewertung',
        'Fehlende Software-Tools zur Unterstutzung der Compliance',
        'Wirtschaftliche Kosten der Compliance im Verhaltnis zum Nutzen',
        'Fehlende Ressourcen (Personal, Budget, Zeit) fur die Umsetzung'
      ],
      lk3: [
        'Erklarbare KI (XAI) wird in den nachsten 3 Jahren zum Industriestandard in unserer Branche.',
        'Der EU AI Act wird langfristig die Qualitat und Sicherheit von KI-Systemen in der EU verbessern.',
        'Unternehmen ohne EU AI Act-Compliance werden erhebliche Wettbewerbsnachteile haben.',
        'Die Anforderungen des EU AI Acts werden in 5 Jahren technisch vollstandig erfullbar sein.'
      ]
    };

    const headers = ['Session-ID','Erstellt','Aktualisiert','Schritt','Abgeschlossen'];
    felder.forEach(([,l]) => headers.push(l));
    Object.values(likert).forEach(items => items.forEach(i => headers.push(i)));

    ws.addRow(headers);
    ws.getRow(1).font = { bold: true };
    ws.getRow(1).fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF0F2D5E' }};
    ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' }};

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
    ws.columns.forEach(c => { c.width = 35; });

    const done = data.filter(r => r.completed);
    ws2.addRow(['Kennzahl','Wert']);
    ws2.getRow(1).font = { bold: true };
    ws2.addRow(['Gesamt Teilnahmen', data.length]);
    ws2.addRow(['Abgeschlossen', done.length]);
    ws2.addRow(['Abgebrochen', data.length - done.length]);
    ws2.addRow(['Abschlussquote', data.length > 0 ? Math.round(done.length / data.length * 100) + '%' : '0%']);
    ws2.addRow(['Export erstellt', new Date().toLocaleString('de-DE')]);
    ws2.columns = [{ width: 30 }, { width: 20 }];

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=umfrage_' + new Date().toISOString().slice(0,10) + '.xlsx');
    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    console.error(e);
    res.status(500).send('Fehler beim Export');
  }
});

connectDB().then(() => {
  app.listen(PORT, () => console.log('Server laeuft auf Port ' + PORT));
}).catch(err => {
  console.error('MongoDB Verbindungsfehler:', err);
  process.exit(1);
});
