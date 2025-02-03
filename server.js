const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");
const { createServer } = require("@vercel/node"); // Vercel server wrapper

const app = express();
const upload = multer({ dest: "/tmp/uploads/" }); // Use /tmp for serverless environments

app.use(express.urlencoded({ extended: true }));

app.get("/", (req, res) => {
    res.send(` 
        <h1>تم النشر على Vercel بنجاح</h1>
        <p>اذهب إلى <strong>/upload</strong> لرفع ملف Excel</p>
    `);
});

app.post("/upload", upload.single("excelFile"), (req, res) => {
    if (!req.file) {
        return res.status(400).send("يرجى رفع ملف Excel");
    }

    const userMessage = encodeURIComponent(req.body.message);
    const filePath = req.file.path;
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    let waLinks = [];
    let names = [];
    data.forEach(row => {
        let phone = String(row["رقم التواصل"]).trim();
        let name = String(row["اسم الطالب"]).trim();
        if (phone.match(/^\d+$/)) {
            let waLink = `https://web.whatsapp.com/send/?phone=${phone}&text=${userMessage}&type=phone_number&app_absent=0`;
            waLinks.push(waLink);
        }
        names.push(name);
    });

    fs.unlinkSync(filePath);

    let htmlLinks = waLinks
        .map((link, i) => `<div style="display:flex;justify-content:center">
            <a style="color:#fff;text-decoration:none" href="${link}" target="_blank">
                <div style="height:150px;width: 250px;border-radius:10px;display:flex;justify-content:center;font-size:20px;align-items:center; background:#6495ed">
                    <h4 style="text-align:center">${names[i]}</h4>
                </div>
            </a>
        </div>`).join("");

    res.send(`
        <h3 style="display:flex;justify-content:center;background:#6bce83">تم إنشاء الروابط بنجاح</h3>
        <p style="text-align:right">:WhatsApp اضغط على الروابط أدناه لإرسال الرسائل عبر </p>
        <div style="display:grid;grid-template-columns: auto auto auto;gap:30px;padding-top:30px">${htmlLinks}</div>
        <br><br>
        <h3 style="display:flex;justify-content:center">Created By: KaramDargham</h3>
    `);
});

// Export the serverless function for Vercel
module.exports = app;
