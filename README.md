## Result Card Generator

Node.js + Express app that reads an Excel file and fills a Word template to generate result cards.

### Folder structure

- `templates/Result Template.docx` (or place `Result Template.docx` at project root)
- `data/class.xlsx` (or place `class.xlsx` at project root)
- `public/index.html`
- `server.js`
- `package.json`

### Install and run

1. Install Node.js (v16+)
2. Install deps:
   ```bash
   npm install
   ```
3. Start:
   ```bash
   npm start
   ```
4. Open `http://localhost:3000`

### Excel columns

`Roll No, Name, Class, Father Name, English, Math, Urdu, Science, Islamiat, Computer, S.st, Nazra, Summer Task`

Notes:
- `S.st` → placeholder `{{Sst}}`
- `Father Name` → `{{FatherName}}`
- `Summer Task` → `{{SummerTask}}`

### Word template placeholders

Insert these tokens exactly in your `Result Template.docx` where values should appear:

```
{{RollNo}}
{{Name}}
{{Class}}
{{FatherName}}
{{English}}
{{Math}}
{{Urdu}}
{{Science}}
{{Islamiat}}
{{Computer}}
{{Sst}}
{{Nazra}}
{{SummerTask}}
{{Total}}
{{Percentage}}
{{Grade}}
```

Keep your template layout, fonts, and spacing unchanged; only replace the variable text with the placeholders above.

### Calculations

- Total = sum of subjects
- Percentage = (Total / 565) × 100, rounded to 2 decimals
- Grade: A+ ≥ 80, A ≥ 70, B ≥ 60, C ≥ 50, else F

### API

- `GET /student/:roll` → preview JSON
- `POST /generate` (form field `roll`) → downloads `Result_<RollNo>_<Name>.docx`

### Frontend usage

1. Enter Roll No
2. Click "Fetch & Preview"
3. Click "Download .docx" (optional: Print)

### Editing placeholders in Word

Open `Result Template.docx` → locate each field (e.g., Name) → replace the existing value with the corresponding placeholder token (e.g., `{{Name}}`). Do not add extra spaces or line breaks inside the curly braces.


