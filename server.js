const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static('public')); // Serve static files

const filePath = path.join(__dirname, 'admit_cards.xlsx'); // Path for the Excel file

// Route to save data
app.post('/save', (req, res) => {
    const { name, rollNo } = req.body;

    // Check if the Excel file already exists
    let workbook;
    if (fs.existsSync(filePath)) {
        // Read existing workbook
        workbook = xlsx.readFile(filePath);
    } else {
        // Create a new workbook
        workbook = xlsx.utils.book_new();
    }

    // Prepare data
    const sheetData = [[name.first, name.last, rollNo]];
    const worksheet = xlsx.utils.aoa_to_sheet(sheetData);

    // Append data to the existing sheet or create a new sheet
    if (workbook.Sheets['AdmitCards']) {
        // If the sheet exists, append new data
        xlsx.utils.sheet_add_aoa(workbook.Sheets['AdmitCards'], sheetData, { origin: -1 });
    } else {
        // Create a new sheet
        xlsx.utils.book_append_sheet(workbook, worksheet, 'AdmitCards');
    }

    // Write the workbook to file
    xlsx.writeFile(workbook, filePath);

    res.json({ message: 'Data saved successfully!' });
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
