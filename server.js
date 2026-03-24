// server.js – FULL CODE WITH EXCEL UPLOAD & MERGE (SAFE, NO CHANGES TO ORIGINAL LOGIC)

const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const pdfMake = require('pdfmake');

const app = express();
const upload = multer({
    dest: 'uploads/'
});

// CONFIG
const DATA_PATH = path.join(__dirname, 'data', 'class.xlsx');
const TOTAL_MARKS_PATH = path.join(__dirname, 'data', 'total.xlsx');
const DONOR_DATA_PATH = path.join(__dirname, 'data', 'donor-students.xlsx');
const TEMPLATE_PATH = path.join(__dirname, 'templates', 'Result Template.docx');
const PHOTO_DIR = path.join(__dirname, 'public', 'photos');

fs.mkdirSync(path.dirname(DATA_PATH), {
    recursive: true
});
fs.mkdirSync(path.dirname(DONOR_DATA_PATH), {
    recursive: true
});
fs.mkdirSync(path.dirname(TEMPLATE_PATH), {
    recursive: true
});
fs.mkdirSync(PHOTO_DIR, {
    recursive: true
});

// pdfmake fonts
const fonts = {
    Roboto: {
        normal: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.12/fonts/Roboto/Roboto-Regular.ttf',
        bold: 'https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.12/fonts/Roboto/Roboto-Medium.ttf',
    }
};
const printer = new pdfMake(fonts);

// === EXCEL → FRONTEND FIELD MAPPING ===
const EXCEL_TO_FRONTEND = {
    'Roll No': 'roll',
    'Term': 'Term',
    'Name': 'Name',
    'Class': 'Class',
    'Father Name': 'FatherName',
    'Admission Date': 'AdmissionDate',
    'Session': 'Session',
    'Semester': 'Semester',
    'Tajveed': 'Tajveed',
    'Remarks': 'Remarks',
    'Total': 'Total',
    'Percentage': 'Percentage',
    'Grade': 'Grade',
    'Geography': 'Geography',
    'English': 'English',
    'Math': 'Math',
    'Urdu': 'Urdu',
    'General Knowledge': 'GeneralKnowledge',
    'Science': 'Science',
    'Islamiat': 'Islamiat',
    'Computer/Biology': 'Computer',
    'S.st': 'Sst',
    'Quraan Pak': 'Nazra',
    'Summer Task': 'SummerTask',
    'Chemistry': 'Chemistry',
    'Physics': 'Physics',
    'Islamiat Elective': 'IslamiatElective',
    'Biology': 'Biology',
    'Genral Math': 'GenralMath',
    'Genral Science': 'GenralScience',
    'Homeconomics': 'Homeconomics',
    'Statistics': 'Statistics',
    'Pakistan Studies': 'PakistanStudies',
    'Tarjama-tul-Quran': 'Tarjama',
    'English (Oral)': 'EnglishOral',
    'English (Written)': 'EnglishWritten',
    'Math (Oral)': 'MathOral',
    'Math (Written)': 'MathWritten',
    'Urdu (Oral)': 'UrduOral',
    'Urdu (Written)': 'UrduWritten',
    'Drawing': 'Drawing'
};

const FRONTEND_TO_EXCEL = Object.fromEntries(
    Object.entries(EXCEL_TO_FRONTEND).map(([k, v]) => [v, k])
);

// HELPERS
function findStudentPhoto(roll) {
    const extensions = ['.jpg', '.jpeg', '.png'];
    for (const ext of extensions) {
        const photoName = `${roll}${ext}`;
        const photoPath = path.join(PHOTO_DIR, photoName);
        if (fs.existsSync(photoPath)) {
            return `/photos/${photoName}`;
        }
    }
    return null;
}

function formatExcelDate(serial) {
    if (!serial || isNaN(serial) || typeof serial !== 'number') return serial;
    // Excel dates are days since 1899-12-30
    const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
    const day = date.getUTCDate();
    const month = date.getUTCMonth() + 1;
    const year = date.getUTCFullYear();
    return `${day}/${month}/${year}`;
}

function ensureExcel() {
    if (!fs.existsSync(DATA_PATH)) {
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([
            []
        ]);
        XLSX.utils.book_append_sheet(wb, ws, 'All Students');
        XLSX.writeFile(wb, DATA_PATH);
    }
}

function loadWorkbook() {
    ensureExcel();
    return XLSX.readFile(DATA_PATH);
}

function saveWorkbook(wb) {
    XLSX.writeFile(wb, DATA_PATH);
}

// MIDDLEWARE
app.use(express.static('public'));
app.use('/photos', express.static(PHOTO_DIR));

// DEBUG
app.get('/debug/paths', (req, res) => {
    res.json({
        dataExists: fs.existsSync(DATA_PATH),
        donorExists: fs.existsSync(DONOR_DATA_PATH),
        templateExists: fs.existsSync(TEMPLATE_PATH),
    });
});

// === FETCH TOTAL MARKS FROM total.xlsx ===
app.get('/totals/:className', (req, res) => {
    const className = req.params.className.trim();
    if (!className) return res.status(400).json({
        error: 'Class required'
    });

    if (!fs.existsSync(TOTAL_MARKS_PATH)) {
        return res.status(404).json({
            error: 'total.xlsx not found'
        });
    }

    try {
        const wb = XLSX.readFile(TOTAL_MARKS_PATH);
        const sheetName = wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws);

        if (rows.length === 0) return res.status(404).json({
            error: 'No data in total.xlsx'
        });

        // Find row where 'Class' column matches className
        const row = rows.find(r => String(r['Class']).trim() === className);
        if (!row) return res.status(404).json({
            error: 'Class not found in total.xlsx'
        });

        // Map Excel keys to frontend keys
        const mappedRow = {};
        Object.keys(row).forEach(excelKey => {
            const frontKey = EXCEL_TO_FRONTEND[excelKey] || excelKey;
            let val = row[excelKey];
            if (frontKey === 'AdmissionDate' && typeof val === 'number') {
                val = formatExcelDate(val);
            }
            mappedRow[frontKey] = val;
        });

        res.json(mappedRow);
    } catch (err) {
        console.error(err);
        res.status(500).json({
            error: 'Error reading total.xlsx'
        });
    }
});

// === FETCH STUDENT FROM SINGLE SHEET ===
app.get('/student/:roll', (req, res) => {
    const roll = req.params.roll.trim();
    if (!roll) return res.status(400).json({
        error: 'Roll required'
    });

    const wb = loadWorkbook();
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, {
        header: 1
    });
    if (rows.length === 0) return res.status(404).json({
        error: 'No data'
    });

    const header = rows[0].map(h => String(h).trim());
    const rollIdx = header.indexOf('Roll No');
    if (rollIdx === -1) return res.status(500).json({
        error: 'Roll No column missing'
    });

    const row = rows.find(r => String(r[rollIdx]).trim() === roll);
    if (!row) return res.status(404).json({
        error: 'Student not found'
    });

    const student = {};
    header.forEach((h, i) => {
        if (h) student[h] = row[i] || '';
    });

    const frontendData = {};
    Object.keys(student).forEach(excelKey => {
        const frontKey = EXCEL_TO_FRONTEND[excelKey] || excelKey;
        let val = student[excelKey];
        if (frontKey === 'AdmissionDate' && typeof val === 'number') {
            val = formatExcelDate(val);
        }
        frontendData[frontKey] = val;
    });

    const photoPath = findStudentPhoto(roll);
    if (photoPath) {
        frontendData.ImagePath = `${photoPath}?t=${Date.now()}`;
    }

    res.json(frontendData);
});

// === SAVE DATA ===
app.post('/save', upload.single('photo'), (req, res) => {
    try {
        const roll = req.body.roll ? req.body.roll.trim() : '';
        if (!roll) throw new Error('Roll No required');

        const wb = loadWorkbook();
        const sheetName = wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, {
            header: 1
        });
        let header = rows[0] || [];

        const allColumns = Object.keys(EXCEL_TO_FRONTEND);
        allColumns.forEach(col => {
            if (!header.includes(col)) {
                const idx = header.length;
                header.push(col);
                rows.forEach(r => r[idx] = '');
            }
        });

        const rollIdx = header.indexOf('Roll No');
        let rowIdx = rows.findIndex(r => String(r[rollIdx]) === roll);
        if (rowIdx === -1) {
            rowIdx = rows.length;
            rows.push(new Array(header.length).fill(''));
            rows[rowIdx][rollIdx] = roll;
        }

        Object.keys(req.body).forEach(frontKey => {
            const excelKey = FRONTEND_TO_EXCEL[frontKey] || frontKey;
            const idx = header.indexOf(excelKey);
            if (idx !== -1) {
                rows[rowIdx][idx] = req.body[frontKey];
            }
        });

        if (req.file) {
            const ext = path.extname(req.file.originalname) || '.jpg';
            // Delete any existing photo with different extension to avoid conflicts
            ['.jpg', '.jpeg', '.png'].forEach(e => {
                const oldPath = path.join(PHOTO_DIR, `${roll}${e}`);
                if (fs.existsSync(oldPath)) fs.unlinkSync(oldPath);
            });
            const dest = path.join(PHOTO_DIR, `${roll}${ext}`);
            fs.renameSync(req.file.path, dest);
        }

        const newWs = XLSX.utils.aoa_to_sheet(rows);
        wb.Sheets[sheetName] = newWs;
        saveWorkbook(wb);

        res.json({
            ok: true
        });
    } catch (err) {
        console.error(err);
        res.status(500).json({
            ok: false,
            error: err.message
        });
    }
});

// === GENERATE DOCX ===
app.post('/generate', upload.none(), async (req, res) => {
    const roll = req.body.roll ? req.body.roll.trim() : '';
    if (!roll) return res.status(400).json({
        error: 'Roll required'
    });

    const studentRes = await fetch(`${req.protocol}://${req.get('host')}/student/${roll}`);
    if (!studentRes.ok) return res.status(404).json({
        error: 'Student not found'
    });
    const student = await studentRes.json();

    if (!fs.existsSync(TEMPLATE_PATH)) {
        return res.status(500).json({
            error: 'DOCX template missing'
        });
    }

    const content = fs.readFileSync(TEMPLATE_PATH, 'binary');
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip);
    doc.setData(student);

    try {
        doc.render();
    } catch (e) {
        return res.status(500).json({
            error: 'Render failed: ' + e.message
        });
    }

    const buf = doc.getZip().generate({
        type: 'nodebuffer'
    });
    res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.set('Content-Disposition', `attachment; filename=Result_${roll}.docx`);
    res.send(buf);
});

// === GENERATE PDF ===
app.post('/generate-pdf', upload.none(), async (req, res) => {
    const roll = req.body.roll ? req.body.roll.trim() : '';
    if (!roll) return res.status(400).json({
        error: 'Roll required'
    });

    const studentRes = await fetch(`${req.protocol}://${req.get('host')}/student/${roll}`);
    if (!studentRes.ok) return res.status(404).json({
        error: 'Student not found'
    });
    const student = await studentRes.json();

    const subjects = Object.keys(student)
        .filter(k => !['roll', 'Name', 'Class', 'FatherName', 'Session', 'Semester', 'Remarks', 'ImagePath', 'Total', 'Percentage', 'Grade'].includes(k))
        .map(k => [k, '75', student[k] || '']);

    const docDefinition = {
        pageSize: 'A4',
        pageMargins: [40, 60, 40, 60],
        content: [{
                text: 'Progress Report',
                style: 'header',
                alignment: 'center'
            },
            {
                text: '\n'
            },
            {
                columns: [{
                        text: `Roll No: ${student.roll || ''}`,
                        style: 'info'
                    },
                    {
                        text: `Name: ${student.Name || ''}`,
                        style: 'info',
                        alignment: 'right'
                    }
                ]
            },
            {
                columns: [{
                        text: `Class: ${student.Class || ''}`,
                        style: 'info'
                    },
                    {
                        text: `Father Name: ${student.FatherName || ''}`,
                        style: 'info',
                        alignment: 'right'
                    }
                ]
            },
            {
                text: '\n'
            },
            {
                table: {
                    headerRows: 1,
                    widths: ['*', 'auto', 'auto'],
                    body: [
                        ['Subject', 'Total', 'Obtained'],
                        ...subjects,
                        [{
                            text: 'Total',
                            bold: true
                        }, {
                            text: student.Total || '',
                            bold: true
                        }, '']
                    ]
                },
                layout: 'lightHorizontalLines'
            },
            {
                text: '\n'
            },
            {
                columns: [{
                        text: `Percentage: ${student.Percentage || ''}%`,
                        style: 'info'
                    },
                    {
                        text: `Grade: ${student.Grade || ''}`,
                        style: 'info',
                        alignment: 'right'
                    }
                ]
            },
            {
                text: `Remarks: ${student.Remarks || ''}`,
                style: 'info',
                margin: [0, 10, 0, 0]
            },
            {
                text: 'Unique School System, Lehtrar Road, Punjgran, Islamabad',
                style: 'footer',
                alignment: 'center',
                margin: [0, 30, 0, 0]
            }
        ],
        styles: {
            header: {
                fontSize: 18,
                bold: true
            },
            info: {
                fontSize: 12
            },
            footer: {
                fontSize: 10,
                italics: true
            }
        }
    };

    const pdfDoc = printer.createPdfKitDocument(docDefinition);
    const chunks = [];
    pdfDoc.on('data', c => chunks.push(c));
    pdfDoc.on('end', () => {
        const pdfBuffer = Buffer.concat(chunks);
        res.set('Content-Type', 'application/pdf');
        res.set('Content-Disposition', `attachment; filename=Result_${roll}.pdf`);
        res.send(pdfBuffer);
    });
    pdfDoc.end();
});

// ===================== DONOR CARD FUNCTIONALITY =====================
app.get('/donor-card', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'donor-card.html'));
});

app.get('/donor-student/:roll', (req, res) => {
    const roll = req.params.roll.trim();
    if (!roll) return res.status(400).json({
        error: 'Roll required'
    });

    if (!fs.existsSync(DONOR_DATA_PATH)) {
        return res.status(404).json({
            error: 'Donor file not found'
        });
    }

    const wb = XLSX.readFile(DONOR_DATA_PATH);
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, {
        header: 1
    });
    if (rows.length === 0) return res.status(404).json({
        error: 'No donor data'
    });

    const header = rows[0].map(h => String(h).trim());
    const rollIdx = header.indexOf('roll') || header.indexOf('Roll No');
    if (rollIdx === -1) return res.status(500).json({
        error: 'Roll column missing'
    });

    const row = rows.find(r => String(r[rollIdx]).trim() === roll);
    if (!row) return res.status(404).json({
        error: 'Donor student not found'
    });

    const student = {};
    header.forEach((h, i) => {
        if (h) student[h.trim()] = row[i] || '';
    });

    const photoPath = findStudentPhoto(roll);
    if (photoPath) {
        student.ImagePath = `${photoPath}?t=${Date.now()}`;
    }

    res.json(student);
});

app.post('/save-donor', upload.single('photo'), (req, res) => {
    try {
        const roll = req.body.roll ? req.body.roll.trim() : '';
        if (!roll) throw new Error('Roll No required');

        let wb;
        if (!fs.existsSync(DONOR_DATA_PATH)) {
            wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet([
                []
            ]);
            XLSX.utils.book_append_sheet(wb, ws, 'Donors');
        } else {
            wb = XLSX.readFile(DONOR_DATA_PATH);
        }

        const sheetName = wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, {
            header: 1
        });
        let header = rows[0] || [];

        const required = ['roll', 'Name', 'Class', 'Session', 'AdmissionDate', 'Donor ID', 'Donor Name'];
        required.forEach(col => {
            if (!header.includes(col)) {
                header.push(col);
                rows.forEach(r => r.push(''));
            }
        });

        const rollIdx = header.indexOf('roll');
        let rowIdx = rows.findIndex(r => String(r[rollIdx]) === roll);
        if (rowIdx === -1) {
            rowIdx = rows.length;
            rows.push(new Array(header.length).fill(''));
            rows[rowIdx][rollIdx] = roll;
        }

        Object.keys(req.body).forEach(key => {
            const idx = header.indexOf(key);
            if (idx !== -1) rows[rowIdx][idx] = req.body[key];
        });

        if (req.file) {
            const ext = path.extname(req.file.originalname) || '.jpg';
            // Delete any existing photo with different extension to avoid conflicts
            ['.jpg', '.jpeg', '.png'].forEach(e => {
                const oldPath = path.join(PHOTO_DIR, `${roll}${e}`);
                if (fs.existsSync(oldPath)) fs.unlinkSync(oldPath);
            });
            const dest = path.join(PHOTO_DIR, `${roll}${ext}`);
            fs.renameSync(req.file.path, dest);
        }

        const newWs = XLSX.utils.aoa_to_sheet(rows);
        wb.Sheets[sheetName] = newWs;
        XLSX.writeFile(wb, DONOR_DATA_PATH);

        res.json({
            ok: true
        });
    } catch (err) {
        console.error(err);
        res.status(500).json({
            ok: false,
            error: err.message
        });
    }
});

// ===================== NEW: UPLOAD & MERGE EXCEL FILE =====================
app.post('/upload-excel', upload.single('excel'), (req, res) => {
    if (!req.file) {
        return res.status(400).json({
            ok: false,
            error: 'No file uploaded'
        });
    }

    try {
        const uploadedPath = req.file.path;
        const uploadedWb = XLSX.readFile(uploadedPath);
        const uploadedSheetName = uploadedWb.SheetNames[0];
        const uploadedRows = XLSX.utils.sheet_to_json(uploadedWb.Sheets[uploadedSheetName], {
            header: 1
        });

        if (uploadedRows.length === 0) {
            fs.unlinkSync(uploadedPath);
            return res.status(400).json({
                ok: false,
                error: 'Uploaded file is empty'
            });
        }

        const header = uploadedRows[0].map(h => String(h).trim());
        const rollIdx = header.indexOf('Roll No');
        if (rollIdx === -1) {
            fs.unlinkSync(uploadedPath);
            return res.status(400).json({
                ok: false,
                error: 'Uploaded file must have "Roll No" column'
            });
        }

        ensureExcel();
        const mainWb = XLSX.readFile(DATA_PATH);
        const mainSheetName = mainWb.SheetNames[0];
        let mainRows = XLSX.utils.sheet_to_json(mainWb.Sheets[mainSheetName], {
            header: 1
        });
        if (mainRows.length === 0) mainRows = [header];

        // Ensure main file has all columns
        header.forEach(col => {
            if (!mainRows[0].includes(col)) {
                const idx = mainRows[0].length;
                mainRows[0][idx] = col;
                for (let i = 1; i < mainRows.length; i++) {
                    mainRows[i][idx] = '';
                }
            }
        });

        // Merge uploaded rows
        for (let i = 1; i < uploadedRows.length; i++) {
            const uploadedRow = uploadedRows[i];
            const rollValue = String(uploadedRow[rollIdx]).trim();

            const existingIdx = mainRows.findIndex((r, idx) => idx > 0 && String(r[mainRows[0].indexOf('Roll No')]).trim() === rollValue);

            if (existingIdx !== -1) {
                // Update existing
                header.forEach((col, colIdx) => {
                    const mainColIdx = mainRows[0].indexOf(col);
                    if (mainColIdx !== -1) {
                        mainRows[existingIdx][mainColIdx] = uploadedRow[colIdx] || '';
                    }
                });
            } else {
                // Add new
                const newRow = new Array(mainRows[0].length).fill('');
                header.forEach((col, colIdx) => {
                    const mainColIdx = mainRows[0].indexOf(col);
                    if (mainColIdx !== -1) newRow[mainColIdx] = uploadedRow[colIdx] || '';
                });
                mainRows.push(newRow);
            }
        }

        const newWs = XLSX.utils.aoa_to_sheet(mainRows);
        mainWb.Sheets[mainSheetName] = newWs;
        XLSX.writeFile(mainWb, DATA_PATH);

        fs.unlinkSync(uploadedPath);
        res.json({
            ok: true,
            message: `Successfully merged ${uploadedRows.length - 1} students!`
        });

    } catch (err) {
        console.error('Excel upload error:', err);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({
            ok: false,
            error: err.message || 'Failed to process file'
        });
    }
});

// =======================================================================
// START SERVER
// =======================================================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
    console.log(`Main Form       → http://localhost:${PORT}`);
    console.log(`Donor Card      → http://localhost:${PORT}/donor-card`);
    console.log(`Excel: ${DATA_PATH}`);
    console.log(`Donor Excel: ${DONOR_DATA_PATH}`);
});
// ===================== STUDENT VIEW PAGE =====================


// ===================== STUDENT RESULT PAGES =====================
app.get('/student-search', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'student-search.html'));
});

app.get('/result-display', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'result-display.html'));
});

// Optional: Redirect root to student search
app.get('/', (req, res) => {
    res.redirect('/student-search');
});