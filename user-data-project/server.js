const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(bodyParser.json());
app.use(express.static('public')); // To serve the index.html

// Initialize or load existing Excel file
const filePath = path.join(__dirname, 'user_data.xlsx');

const initializeExcelFile = () => {
    if (!fs.existsSync(filePath)) {
        const workBook = xlsx.utils.book_new();
        const workSheet = xlsx.utils.json_to_sheet([]);
        xlsx.utils.book_append_sheet(workBook, workSheet, 'Users');
        xlsx.writeFile(workBook, filePath);
    }
};

// Route to handle form submission
app.post('/submit', (req, res) => {
    try {
        const userData = req.body;

        // Load the existing workbook
        const workBook = xlsx.readFile(filePath);
        const workSheet = workBook.Sheets['Users'];

        // Get existing data and add the new entry
        const existingData = xlsx.utils.sheet_to_json(workSheet);
        existingData.push({
            Name: userData.name,
            Email: userData.email,
            Gender: userData.gender,
        });

        // Convert back to sheet and write to the file
        const updatedSheet = xlsx.utils.json_to_sheet(existingData);
        workBook.Sheets['Users'] = updatedSheet;
        xlsx.writeFile(workBook, filePath);

        res.json({ message: 'Data saved successfully' });
    } catch (error) {
        res.status(500).json({ message: 'Error saving data' });
    }
});

// Route to download Excel file
app.get('/admin/download', (req, res) => {
    try {
        res.download(filePath, 'user_data.xlsx');
    } catch (error) {
        res.status(500).json({ message: 'Error downloading file' });
    }
});

// Start the server and initialize Excel file
app.listen(PORT, () => {
    initializeExcelFile();
    console.log(`Server is running on http://localhost:${PORT}`);
});
