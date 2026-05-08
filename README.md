# GIS Standards Landscape

A lightweight Explorer for EN/ISO geographic information standards.  
**Stack:** Node.js · Express · vanilla HTML/CSS/JavaScript · SheetJS · vis-network

---

## Setup

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/gis-standards-landscape.git
cd gis-standards-landscape

# 2. Place your Excel file in the project root
#    (must be named exactly:)
cp /path/to/GIS-standards-landscape.xlsx .

# 3. Install dependencies
npm install

# 4. Start
npm start
```

Then open **http://localhost:3000**

---

## Features

| Feature | Details |
|---|---|
| **Catalogue** | Searchable card grid with text, body (CEN/ISO) and committee filters |
| **Network graph** | Force-directed graph of all standards and their references (vis-network) |
| **Detail dialog** | Scope, years, ICS, type, web link, outgoing and incoming references |
| **Navigation** | Click any reference chip in the dialog to jump to that standard |

---

## Project structure

```
├── server.js                 Express server – reads Excel, serves /api/*
├── package.json
├── GIS-standards-landscape.xlsx   ← your data file (required, not in repo)
└── public/
    ├── index.html
    ├── style.css
    └── app.js
```

> **Note:** The Excel file is excluded from version control by default (see `.gitignore`).  
> Add it manually after cloning, or remove it from `.gitignore` if you want to track it.

---

## Environment variables

| Variable | Default | Purpose |
|---|---|---|
| `PORT` | `3000` | HTTP port |
