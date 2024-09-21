const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3000;

// Set up multer for file uploads
const upload = multer({ dest: 'uploads/' });

// Serve the static HTML and JavaScript files
app.use(express.static('public'));

// Helper function to check if a worksheet is empty
const isSheetEmpty = (worksheet) => {
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  return jsonData.length === 0 || jsonData.every(row => row.length === 0);
};

// Helper function to merge Excel files into different sheets in the correct order
const mergeExcelFiles = (files) => {
  const newWorkbook = XLSX.utils.book_new();

  let sheetIndex = 1; // To track the new sheet numbering

  files.forEach((file) => {
    const workbook = XLSX.readFile(file.path);
    const sheetNames = workbook.SheetNames;

    sheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];

      // Skip the sheet if it's empty
      if (!isSheetEmpty(worksheet)) {
        const newSheetName = `Sheet${sheetIndex}`; // Rename sheet to "Sheet1", "Sheet2", etc.
        XLSX.utils.book_append_sheet(newWorkbook, worksheet, newSheetName);
        sheetIndex++; // Increment the sheet index for the next sheet
      }
    });
  });

  // Save merged workbook to a file
  const mergedFilePath = path.join(__dirname, 'uploads', 'merged_file.xlsx');
  XLSX.writeFile(newWorkbook, mergedFilePath);

  return mergedFilePath;
};

// API endpoint to handle the file upload and merging
app.post('/merge', upload.fields([
  { name: 'file1', maxCount: 1 },
  { name: 'file2', maxCount: 1 },
  { name: 'file3', maxCount: 1 },
  { name: 'file4', maxCount: 1 }
]), (req, res) => {
  try {
    const files = [
      req.files.file1[0],
      req.files.file2[0],
      req.files.file3[0],
      req.files.file4[0]
    ];

    const mergedFilePath = mergeExcelFiles(files);

    // Send the merged file as a response
    res.download(mergedFilePath, 'merged_file.xlsx', (err) => {
      if (err) {
        console.error('Error downloading merged file:', err);
        res.status(500).json({ error: 'Error downloading merged file' });
      }

      // Clean up uploaded and merged files after the download
      files.forEach((file) => {
        fs.unlinkSync(file.path);
      });
      fs.unlinkSync(mergedFilePath);
    });
  } catch (error) {
    console.error('Error merging files:', error);
    res.status(500).json({ error: 'Error merging files' });
  }
});

// Start the server
app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
