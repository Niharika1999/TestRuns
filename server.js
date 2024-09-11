const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const cors = require('cors');

const app = express();
app.use(bodyParser.json());
app.use(cors());

const filePath = path.join(__dirname, 'assessment_data.xlsx');

app.post('/update-excel', async (req, res) => {
    try {
        const formData = req.body;
        formData.services = formData.services.join(', '); // Format services as a comma-separated string

        let workbook = new ExcelJS.Workbook();
        let worksheet;

        if (fs.existsSync(filePath)) {
            await workbook.xlsx.readFile(filePath);
            worksheet = workbook.getWorksheet('Assessment Data');

            // If the worksheet doesn't exist, create it
            if (!worksheet) {
                worksheet = workbook.addWorksheet('Assessment Data');
                worksheet.addRow(['Full Name', 'Date of Birth', 'Gender', 'Phone Number', 'Email Address', 'Address', 'Services Interested', 'Questions']);
            }
        } else {
            // Create a new workbook and worksheet if the file doesn't exist
            worksheet = workbook.addWorksheet('Assessment Data');
            worksheet.addRow(['Full Name', 'Date of Birth', 'Gender', 'Phone Number', 'Email Address', 'Address', 'Services Interested', 'Questions']);
        }

        // Prepare the new data to be appended
        const data = [
            formData.fullName,
            formData.dob,
            formData.gender,
            formData.phoneNumber,
            formData.email,
            formData.address,
            formData.services,
            formData.questions
        ];

        // Append the data to the worksheet
        worksheet.addRow(data);

        // Write the updated workbook back to the file
        await workbook.xlsx.writeFile(filePath);

        res.json({ success: true });
    } catch (error) {
        console.error('Error updating Excel file:', error);
        res.status(500).json({ success: false, message: 'Error updating Excel file.' });
    }
});

app.listen(3000, () => {
    console.log('Server is running on port 3000');
});
