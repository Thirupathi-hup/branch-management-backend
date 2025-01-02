const express = require("express");
const sqlite3 = require("sqlite3").verbose();
const cors = require("cors");
const multer = require("multer");
const xlsx = require("xlsx");
const bodyParser = require("body-parser");
const fs = require("fs");

const app = express();
const port = 5000;

app.use(cors());
app.use(bodyParser.json());

// Database setup
const db = new sqlite3.Database("./branches.db", (err) => {
    if (err) return console.error("Error opening database:", err.message);
    db.run(
        `CREATE TABLE IF NOT EXISTS branches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            location TEXT,
            manager TEXT
        )`
    );
});

// CRUD APIs
app.get("/api/branches", (req, res) => {
    db.all("SELECT * FROM branches", [], (err, rows) => {
        if (err) {
            console.error("Error fetching branches:", err.message);
            return res.status(500).json({ error: "Failed to fetch branches" });
        }
        res.json(rows);
    });
});

app.post("/api/branches", (req, res) => {
    const { name, location, manager } = req.body;
    db.run(
        `INSERT INTO branches (name, location, manager) VALUES (?, ?, ?)`,
        [name, location, manager],
        function (err) {
            if (err) {
                console.error("Error inserting branch:", err.message);
                return res.status(500).json({ error: "Failed to add branch" });
            }
            res.json({ id: this.lastID });
        }
    );
});

app.put("/api/branches/:id", (req, res) => {
    const { name, location, manager } = req.body;
    const { id } = req.params;
    db.run(
        `UPDATE branches SET name = ?, location = ?, manager = ? WHERE id = ?`,
        [name, location, manager, id],
        function (err) {
            if (err) {
                console.error("Error updating branch:", err.message);
                return res.status(500).json({ error: "Failed to update branch" });
            }
            res.sendStatus(200);
        }
    );
});

app.delete("/api/branches/:id", (req, res) => {
    const { id } = req.params;
    db.run(`DELETE FROM branches WHERE id = ?`, [id], function (err) {
        if (err) {
            console.error("Error deleting branch:", err.message);
            return res.status(500).json({ error: "Failed to delete branch" });
        }
        res.sendStatus(200);
    });
});

// Excel Import/Export
const upload = multer({ dest: "uploads/" });

app.post("/api/import", upload.single("file"), (req, res) => {
    const workbook = xlsx.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);
    data.forEach(({ name, location, manager }) => {
        db.run(
            `INSERT INTO branches (name, location, manager) VALUES (?, ?, ?)`,
            [name, location, manager]
        );
    });

    // Clean up the uploaded file
    fs.unlink(req.file.path, (err) => {
        if (err) {
            console.error("Error deleting file:", err.message);
        }
    });

    res.sendStatus(200);
});

app.get("/api/export", (req, res) => {
    db.all("SELECT * FROM branches", [], (err, rows) => {
        if (err) {
            console.error("Error exporting branches:", err.message);
            return res.status(500).json({ error: "Failed to export branches" });
        }
        const worksheet = xlsx.utils.json_to_sheet(rows);
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, "Branches");
        const filePath = "./branches.xlsx";
        xlsx.writeFile(workbook, filePath);
        res.download(filePath);
    });
});

app.listen(port, () => console.log(`Server running on port ${port}`));
