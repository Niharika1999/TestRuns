const express = require('express');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const cors = require('cors');

const app = express();
app.use(bodyParser.json());
app.use(cors());

const filePath = path.join(__dirname, 'assessment_data.xlsx');

app.post('/update-excel', (req, res) => {
    try {
        const formData = req.body;
        formData.services = formData.services.join(', '); // Format services as a comma-separated string

        let workbook;
        let worksheet;

        if (fs.existsSync(filePath)) {
            workbook = XLSX.readFile(filePath);
            worksheet = workbook.Sheets['Assessment Data'];

            // If the worksheet doesn't exist, create it
            if (!worksheet) {
                worksheet = XLSX.utils.aoa_to_sheet([
                    ['Full Name', 'Date of Birth', 'Gender', 'Phone Number', 'Email Address', 'Address', 'Services Interested', 'Questions']
                ]);
                XLSX.utils.book_append_sheet(workbook, worksheet, "Assessment Data");
            }
        } else {
            // Create a new workbook and worksheet if the file doesn't exist
            workbook = XLSX.utils.book_new();
            worksheet = XLSX.utils.aoa_to_sheet([
                ['Full Name', 'Date of Birth', 'Gender', 'Phone Number', 'Email Address', 'Address', 'Services Interested', 'Questions']
            ]);
            XLSX.utils.book_append_sheet(workbook, worksheet, "Assessment Data");
        }

        // Prepare the new data to be appended
        const data = [
            [
                formData.fullName,
                formData.dob,
                formData.gender,
                formData.phoneNumber,
                formData.email,
                formData.address,
                formData.services,
                formData.questions
            ]
        ];

        // Append the data to the worksheet
        XLSX.utils.sheet_add_aoa(worksheet, data, { origin: -1 });

        // Write the updated workbook back to the file
        XLSX.writeFile(workbook, filePath);

        res.json({ success: true });
    } catch (error) {
        console.error('Error updating Excel file:', error);
        res.status(500).json({ success: false, message: 'Error updating Excel file.' });
    }
});

app.listen(3000, () => {
    console.log('Server is running on port 3000');
});
