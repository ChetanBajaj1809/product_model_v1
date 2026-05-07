# PM Explorer — Vercel Deployment Guide

A web app to explore Product Model coverages, terms, options, and business rules by uploading your Excel file.

---

## Deploy in 3 steps

### 1. Install dependencies (one-time)
```bash
npm install -g vercel
```

### 2. Install project dependencies
```bash
cd pm-explorer
npm install
```

### 3. Deploy
```bash
vercel --prod
```

Vercel will ask a few questions on first deploy:
- **Set up and deploy?** → Y
- **Which scope?** → your account
- **Link to existing project?** → N
- **Project name** → pm-explorer (or anything you like)
- **In which directory is your code?** → `.` (current directory)
- **Want to modify settings?** → N

Your app will be live at `https://pm-explorer-xxxx.vercel.app`

---

## Project structure

```
pm-explorer/
├── public/
│   └── index.html        # Frontend (upload screen + explorer)
├── api/
│   └── parse.js          # Serverless function: parses the Excel upload
├── package.json
└── vercel.json
```

---

## How it works

1. You visit the URL and upload your `.xlsx` file
2. The file is sent to `/api/parse` (a Vercel serverless function)
3. The API reads the **Clauses**, **Terms**, **Options**, and **PM Business Rules** tabs
4. It joins everything by coverage name and returns structured JSON
5. The frontend renders the unified coverage explorer

Sheet names are detected flexibly — as long as the tabs contain the words "Clause", "Term", "Option", and "Business Rule", it will find them.

---

## Local development

```bash
npm install
vercel dev
```

Then open `http://localhost:3000`
