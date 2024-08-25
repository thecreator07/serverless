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

// In-memory storage for user data
let userData = [];

// Route to handle form submission
app.post('/submit', (req, res) => {
    try {
        const user = req.body;
        userData.push({
            Name: user.name,
            Email: user.email,
            Gender: user.gender,
        });
        res.json({ message: 'Data saved successfully' });
    } catch (error) {
        res.status(500).json({ message: 'Error saving data' });
    }
});

// Route to download Excel file
app.get('/admin/download', (req, res) => {
    try {
        // Generate the Excel file dynamically
        const workBook = xlsx.utils.book_new();
        const workSheet = xlsx.utils.json_to_sheet(userData);
        xlsx.utils.book_append_sheet(workBook, workSheet, 'Users');

        // Write the workbook to a buffer
        const buffer = xlsx.write(workBook, { type: 'buffer', bookType: 'xlsx' });

        // Set the headers and send the buffer as a file download
        res.setHeader('Content-Disposition', 'attachment; filename="user_data.xlsx"');
        res.setHeader('Content-Type', 'application/octet-stream');
        res.send(buffer);
    } catch (error) {
        res.status(500).json({ message: 'Error generating file' });
    }
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
