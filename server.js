const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
const upload = multer({ dest: "/tmp/uploads/" });

app.use(express.urlencoded({ extended: true }));

app.get("/", (req, res) => {
    res.send(`
    <form dir="rtl" action="/upload" method="POST" enctype="multipart/form-data" 
          style="background: white; padding: 20px;margin-top:30px !important; border-radius: 10px; 
                 box-shadow: 0px 0px 10px rgba(0, 0, 0, 0.1); max-width: 400px; margin: auto;text-align:right;">
        
        <label style="font-weight: bold; font-size: 16px;"> اختر ملف Excel</label>
        <input type="file" text-align:center; name="excelFile" accept=".xlsx, .xls" required 
               style="width: 100%; padding: 10px; margin: 10px 0; 
                      border: 1px solid #ccc; border-radius: 5px; font-size: 16px;">

        <label style="font-weight: bold; font-size: 16px;">أدخل الرسالة التي تريد إرسالها</label>
        <textarea name="message" rows="4" required 
                  style="width: 100%; padding: 10px; margin: 10px 0; text-align:right;
                         border: 1px solid #ccc; border-radius: 5px; font-size: 16px;"></textarea>

        <button type="submit" 
                style="width: 100%; padding: 10px; margin: 10px 0; border: none; 
                       border-radius: 5px; font-size: 16px; background-color: #28a745; 
                       color: white; cursor: pointer;">
            رفع الملف وإنشاء الروابط
        </button>

    </form>
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
        // استخدام الأسماء العربية للأعمدة كما في ملف Excel
        let phone = String(row["رقم التواصل"] || row["D"] || "").trim();
        let name = String(row["اسم الطالب"] || row["C"] || "").trim();
        
        // تنظيف رقم الهاتف من الأحرف غير الرقمية
        phone = phone.replace(/[^\d+]/g, '');
        
        // التأكد من أن الرقم يبدأ بـ + أو 00
        if (phone && !phone.startsWith('+') && !phone.startsWith('00')) {
            phone = '+'+phone;
        }
        
        if (phone && phone.length > 5) {
            let waLink = `https://web.whatsapp.com/send/?phone=${phone}&text=${userMessage}&type=phone_number&app_absent=0`;
            waLinks.push(waLink);
            names.push(name);
        }
    });

    fs.unlinkSync(filePath);

    let htmlLinks = waLinks.map((link, i) => `
        <div style="display:flex;justify-content:center">
            <a style="color:#fff;text-decoration:none" href="${link}" target="_blank">
                <div onclick="this.style.background='purple'" 
                     style="height:150px;width: 250px;border-radius:10px;display:flex;
                            justify-content:center;font-size:20px;align-items:center; 
                            background:#6495ed; cursor: pointer;">
                    <h4 style="text-align:center">${names[i] || 'طالب'}</h4>
                </div>
            </a>
        </div>`).join("");
    
    res.send(`
        <h3 style="display:flex;justify-content:center;background:#6bce83">تم إنشاء الروابط بنجاح</h3>
        <p style="text-align:right">اضغط على الروابط أدناه لإرسال الرسائل عبر WhatsApp:</p>
        <div style="display:grid;grid-template-columns: auto auto auto;gap:30px;padding-top:30px"> 
            ${htmlLinks} 
        </div>
        <br><br>
        <h3 style="display:flex;justify-content:center">Created By: KaramDargham</h3>
    `);
});

module.exports = app;
