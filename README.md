# Umfrage: EU AI Act in der Praxis

## Setup (5 Minuten)

### 1. Node.js installieren
Falls noch nicht vorhanden: https://nodejs.org (LTS-Version)

### 2. Abhängigkeiten installieren
```
npm install
```

### 3. Server starten
```
npm start
```
Die Umfrage ist jetzt unter http://localhost:3000 erreichbar.

---

## Hosting (kostenlos, öffentlich erreichbar)

### Option A: Railway (empfohlen, einfachster Weg)
1. Konto anlegen auf https://railway.app
2. "New Project" → "Deploy from GitHub" oder Dateien hochladen
3. Fertig – du bekommst eine öffentliche URL

### Option B: Render
1. Konto anlegen auf https://render.com
2. "New Web Service" → Dateien hochladen
3. Start command: `npm start`

---

## Daten abrufen

### Excel-Export (alle Antworten)
```
http://deine-url.com/admin/export?token=admin123
```
→ Lädt eine Excel-Datei herunter mit:
- Sheet 1: Alle Antworten (vollständig + abgebrochene)
- Sheet 2: Zusammenfassung (Anzahl, Abschlussquote)

### Admin-Token ändern (empfohlen!)
Setze die Umgebungsvariable ADMIN_TOKEN auf einen eigenen geheimen Wert,
z. B. in Railway unter "Variables": ADMIN_TOKEN=meinGeheimesPasswort123

Dann abrufbar unter:
http://deine-url.com/admin/export?token=meinGeheimesPasswort123

---

## Zwischenspeicherung
- Jede Seite wird automatisch an den Server gesendet (auch abgebrochene Umfragen)
- Im Excel-Export siehst du bei jeder Antwort, ob sie abgeschlossen wurde
- Die Spalte "Schritt" zeigt, wie weit jemand gekommen ist

## Datenstruktur
Alle Rohdaten liegen in: `data/responses.json`
