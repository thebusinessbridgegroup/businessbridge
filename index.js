const express = require('express');
const multer = require('multer');
const mysql = require('mysql2');
const cors = require('cors');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3002;

app.use(cors());
app.use(express.json());

// Multer configuration for handling file uploads
const upload = multer({ dest: 'uploads/' }); // Destination folder for uploaded files

// MySQL connection setup
const db = mysql.createConnection({
  host: 'localhost',
  user: 'root',
  password: '',
  database: 'blue_bridge',
  port: 4306
});

db.connect(err => {
  if (err) {
    console.error('Error connecting to the database:', err);
    return;
  }
  console.log('Connected to the MySQL database');
});

// Root route
app.get('/', (req, res) => {
  res.send('Welcome to the Blue Bridge Group API');
});
// Route to fetch marked dates
app.get('/api/marked-dates', (req, res) => {
  const sql = 'SELECT date FROM calendar';
  db.query(sql, (err, results) => {
    if (err) {
      res.status(500).json({ error: 'Internal server error' });
    } else {
      const markedDates = results.map(result => result.date);
      res.json(markedDates);
    }
  });
});

// Search route
app.get('/search', (req, res) => {
  const searchQuery = req.query.q;
  console.log('Received search query:', searchQuery);
  
  if (!searchQuery) {
    return res.status(400).json({ error: 'No search query provided' });
  }

  const sql = `
    SELECT \`Query\`, \`PDF File\`, \`Software Link\`
    FROM searchers
    WHERE query LIKE ?
  `;
  console.log('SQL query:', sql);
  
  db.query(sql, [`%${searchQuery}%`], (err, results) => {
    if (err) {
      console.error('Error executing SQL query:', err);
      return res.status(500).json({ error: err.message });
    }
    console.log('Search results:', results);
    res.json(results);
  });
});

// Route for handling file uploads
app.post('/upload', upload.single('file'), async (req, res) => {
  const file = req.file;
  if (!file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }

  const originalFilename = file.originalname;
  const filePath = file.path;
  const targetPath = path.join(__dirname, 'uploads', originalFilename);

  // Rename the file to keep its original extension
  fs.rename(filePath, targetPath, async (err) => {
    if (err) {
      console.error('Error renaming file:', err);
      return res.status(500).json({ error: 'Error processing file' });
    }

    // Process the uploaded file
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(targetPath);

      // Assuming the data is in the first sheet
      const worksheet = workbook.getWorksheet(1);
      const data = [];

      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const rowData = row.values;
        data.push(rowData);
        console.log(`Row ${rowNumber}: ${JSON.stringify(rowData)}`);
      });

      // Here you can save the data to the database or perform other operations

      res.json({ message: 'File uploaded and processed successfully', data });
    } catch (error) {
      console.error('Error processing file:', error);
      res.status(500).json({ error: 'Error processing file' });
    }
  });
});

// Download reconciliation format route
app.get('/download-reconciliation-format', async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Database');

    // Add columns
    worksheet.columns = [
      { header: 'Date', key: 'date', width: 15 },
      { header: 'GST No.', key: 'gst_no', width: 20 },
      { header: 'Suppliers name', key: 'suppliers_name', width: 30 },
      { header: 'Invoice no.', key: 'invoice_no', width: 20 },
      { header: 'Taxable Value', key: 'taxable_value', width: 15 },
      { header: 'IGST', key: 'igst', width: 15 },
      { header: 'CGST', key: 'cgst', width: 15 },
      { header: 'SGST', key: 'sgst', width: 15 },
      { header: 'Remarks', key: 'remarks', width: 30 },
      { header: 'Type', key: 'type', width: 20 }
    ];

    // Sample data for testing
    worksheet.addRow({ date: '2024-01-01', gst_no: '123456789', suppliers_name: 'Supplier A', invoice_no: 'INV001', taxable_value: 1000, igst: 180, cgst: 90, sgst: 90, remarks: '', type: '' });
    worksheet.addRow({ date: '2024-01-02', gst_no: '987654321', suppliers_name: 'Supplier B', invoice_no: 'INV002', taxable_value: 2000, igst: 360, cgst: 180, sgst: 180, remarks: '', type: '' });

    // Apply border to header row
    worksheet.getRow(1).eachCell(cell => {
      cell.border = {
        bottom: { style: 'thin', color: { argb: '000000' } }
      };
    });

    // Generate a unique filename
    const filename = 'reconciliation_format.xlsx';
    const dirPath = path.join(__dirname, 'temp');

    // Ensure directory exists
    if (!fs.existsSync(dirPath)) {
      fs.mkdirSync(dirPath, { recursive: true });
    }

    const filepath = path.join(dirPath, filename);

    // Write the workbook to a file
    await workbook.xlsx.writeFile(filepath);

    // Check if the file exists
    if (!fs.existsSync(filepath)) {
      console.error('Error: File does not exist:', filepath);
      return res.status(500).send('Error generating Excel file');
    }

    // Set headers and send the file
    res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    fs.createReadStream(filepath).pipe(res).on('finish', () => {
      // Optionally delete the file after sending it
      fs.unlinkSync(filepath);
    });
  } catch (error) {
    console.error('Error generating Excel file:', error);
    res.status(500).send('Error generating Excel file');
  }
});

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
