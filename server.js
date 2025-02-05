const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const fs = require("fs");
const XLSX = require("xlsx");

const app = express();
app.use(bodyParser.json());
app.use(cors());

const FILE_PATH = "data.xlsx";

// Function to read existing Excel file or create a new one
function readExcelFile() {
    if (fs.existsSync(FILE_PATH)) {
        const workbook = XLSX.readFile(FILE_PATH);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        return XLSX.utils.sheet_to_json(sheet);
    }
    return [];
}

// Function to write data to Excel file
function writeExcelFile(data) {
    console.log("Saving data to Excel:", data); // Debugging log
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Users");
    XLSX.writeFile(workbook, FILE_PATH);
    console.log("Excel file updated successfully!"); // Debugging log
}

// API to add user data
app.post("/add", (req, res) => {
    const { name, age } = req.body;
    if (!name || !age) {
        return res.status(400).json({ message: "Name and Age are required" });
    }

    const users = readExcelFile();
    users.push({ Name: name, Age: age });

    writeExcelFile(users);
    res.json({ message: "Data saved successfully to Excel" });
});

app.listen(5000, () => {
    console.log("Server is running on port 5000");
});
