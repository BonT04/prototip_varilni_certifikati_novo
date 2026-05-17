# Palfinger – Vnos podatkov varilnih izvorov

Spletna aplikacija za hitri vnos meritev varilnih izvorov in izvoz v Excel format.

---

## 📁 Struktura projekta

```
palfinger-vnos/
├── index.html
├── package.json
├── vite.config.js
├── palfinger_export.py      ← Python skripta za izvoz v XLSX
└── src/
    ├── main.jsx
    └── App.jsx              ← glavna aplikacija
```

---

## 🚀 Lokalni zagon (razvoj)

Potrebujete: **Node.js** → prenesite na https://nodejs.org (izberite LTS verzijo)

```bash
# 1. V terminalu pojdite v mapo projekta
cd palfinger-vnos

# 2. Namestite odvisnosti (samo prvič)
npm install

# 3. Zaženite razvojni strežnik
npm run dev
```

Aplikacija se odpre na http://localhost:5173

---

## ☁️ Objava na Vercel (brezplačno)

### Korak 1 — GitHub
1. Pojdite na https://github.com → **New repository**
2. Ime: `palfinger-vnos`, nastavite na **Private**
3. Kliknite **Create repository**
4. Naložite vse datoteke tega projekta (**Add file → Upload files**)

### Korak 2 — Vercel
1. Pojdite na https://vercel.com → **Sign up with GitHub**
2. Kliknite **Add New Project**
3. Izberite vaš `palfinger-vnos` repozitorij
4. Vercel samodejno zazna Vite → kliknite **Deploy**
5. Čez ~1 minuto dobite javni URL, npr. `palfinger-vnos.vercel.app`

Vsaka sprememba na GitHubu se samodejno objavi na Vercel.

---

## 📤 Izvoz v Excel

### V aplikaciji:
1. Vnesite podatke in kliknite **{ } PRIKAŽI JSON**
2. Kliknite **⧉ KOPIRAJ JSON**
3. Odprite Beležnico (Notepad), prilepite, shranite kot `"palfinger_podatki.json"`
   - Ime datoteke: `"palfinger_podatki.json"` (z narekovaji!)
   - Vrsta: **Vse datoteke**
   - Kodiranje: **UTF-8**

### Python skripta:
```bash
# Namestite openpyxl (samo prvič)
pip install openpyxl

# Zaženite izvoz
python palfinger_export.py palfinger_podatki.json palfinger_glavni.xlsx izhod.xlsx
```

---

## ⬆️ Uvoz iz Excel

V aplikaciji kliknite **⬆ UVOZI XLSX** in izberite `palfinger_glavni.xlsx`.  
Aplikacija samodejno prebere vse naprave in splošne podatke.
