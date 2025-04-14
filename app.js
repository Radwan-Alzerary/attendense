// Import required modules
const express = require('express');
const nodemailer = require('nodemailer');
const XLSX = require('xlsx');
const path = require('path');
const crypto = require('crypto');
const mongoose = require('mongoose');

const app = express();
const port = 3110;
// ------------------------------
// Connect to MongoDB using Mongoose
// ------------------------------
mongoose.connect('mongodb://admin:fgdfds432rtegf4wtesfdvxwefsd@localhost:27017/attendees?authSource=attendees', {
  useNewUrlParser: true,
  useUnifiedTopology: true,
})
.then(() => console.log("Connected to MongoDB database"))
.catch((err) => console.error("Failed to connect to MongoDB:", err.message));

// ------------------------------
// Define the Registration Schema and Model
// ------------------------------
const registrationSchema = new mongoose.Schema({
  name: String,
  email: String,
  phone: String,
  token: String,
  field:String,
  confirmed: { type: Boolean, default: false },
});

const Registration = mongoose.model('Registration', registrationSchema);

// ------------------------------
// Setup Nodemailer Transporter
// ------------------------------
// Replace these details with your own Gmail credentials and app password.
const transporter = nodemailer.createTransport({
  service: "Gmail",
  host: "smtp.gmail.com",
  port: 465,
  secure: true, // Use true for port 465
  auth: {
    user: "mdociniraqinfo@gmail.com",
    pass: "tqnl ppbs gaix joeg", // Your 16-character App Password WITHOUT spaces
  },
});

// ------------------------------
// Helper Function: Send Confirmation Email
// ------------------------------
/**
 * Sends a confirmation email to a recipient.
 *
 * @param {Object} details - Contains name, email, and token.
 */
function sendConfirmationEmail({ name, email, token }) {
  // Generate unique confirmation URL
  const confirmUrl = `https://attendance.niuraiq.com:${port}/confirm/${token}`;

  const mailOptions = {
    from: '"radwan" <mdociniraqinfo@gmail.com>',
    to: email,
    subject: 'دعوة حضور دورة MDOC',
    text: `من فضلك قم بتأكيد حضورك عبر الرابط التالي: ${confirmUrl}`,
    html: `<!DOCTYPE html>
<html lang="ar">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>دعوة الفعالية</title>
  <style>
    body { font-family: Arial, sans-serif; background-color: #f4f4f4; margin: 0; padding: 0; direction: rtl; text-align: right; }
    .container { width: 90%; max-width: 600px; background-color: #ffffff; margin: 20px auto; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); }
    .header { background-color: #007BFF; color: #ffffff; padding: 20px; text-align: center; direction: ltr; }
    .header h1 { margin: 0; font-size: 22px; }
    .header p { margin: 5px 0 0; font-size: 16px; }
    .content { padding: 20px; color: #333333; }
    .content h2 { color: #007BFF; font-size: 20px; margin-bottom: 15px; }
    .content p { line-height: 1.6; margin: 10px 0; }
    .button { display: inline-block; background-color: #28a745; color: #ffffff; text-decoration: none; padding: 10px 20px; border-radius: 5px; margin-top: 20px; font-weight: bold; text-align: center; }
    .footer { background-color: #f1f1f1; color: #777777; padding: 10px; text-align: center; font-size: 12px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>4th Multi-Disciplinary Oncology Course (MDOC) Iraq Series</h1>
      <p>Toward Better Care For Women’s Cancers</p>
    </div>
    <div class="content">
      <h2>تأكيد الحضور</h2>
      <p>مرحباً ${name},</p>
      <p>يرجى تأكيد حضورك بالنقر على الزر أدناه:</p>
      <a href="${confirmUrl}" class="button">تأكيد الحضور</a>
      <p>مع أطيب التحيات ونتطلع إلى لقائك.</p>
    </div>
    <div class="footer">
      <p>&copy; 2025 MDOC Iraq Series. جميع الحقوق محفوظة.</p>
    </div>
  </div>
</body>
</html>`
  };

  transporter.sendMail(mailOptions, (error, info) => {
    if (error) {
      console.error(`Error sending email to ${email}:`, error);
    } else {
      console.log(`Email sent successfully to ${email}. Message ID: ${info.messageId}`);
    }
  });
}

// ------------------------------
// Process the Excel File and Save Records to MongoDB
// ------------------------------
/**
 * Reads an Excel file and processes each row.
 *
 * Expects the Excel file to have columns (or similar) for:
 * - Name (or name)
 * - Email (or email)
 * - Phone (or phone / Number)
 *
 * @param {string} filePath - Path to the Excel file.
 */
function processExcelFile(filePath) {
  // Read the workbook
  const workbook = XLSX.readFile(filePath);
  // Assume data is in the first sheet
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  // Convert sheet to JSON format
  const data = XLSX.utils.sheet_to_json(sheet);

  data.forEach(async (row) => {
    // Adjust the property names based on your Excel file headers
    const name = row.NameFirst || row.name;
    const email = row.Email || row.email;
    const phone = row.Phone    || row.phone || row.Number || row.number;
    const field    = row.Field    || row.phone || row.Number || row.number;
    
    // Generate a unique token (32 characters long hex string)
    const token = crypto.randomBytes(16).toString('hex');

    // Create a new registration document
    const registration = new Registration({ name, email, phone,field, token });

    try {
      await registration.save();
      console.log(`Inserted record for ${name} with email ${email}.`);
      // After saving, send the confirmation email with the unique URL.
      sendConfirmationEmail({ name, email, token });
    } catch (err) {
      console.error("Error inserting data into MongoDB:", err.message);
    }
  });
}

// ------------------------------
// Express Route for Confirming Attendance
// ------------------------------
app.get('/confirm/:token', async (req, res) => {
  const token = req.params.token;

  try {
    const registration = await Registration.findOne({ token });
    if (!registration) {
      return res.status(404).send("الرابط غير صالح.");
    }

    if (registration.confirmed) {
      return res.send("لقد تم تأكيد الحضور مسبقاً.");
    }

    // Mark as confirmed
    registration.confirmed = true;
    await registration.save();

    res.send("تم تاكيد الحضور");
  } catch (err) {
    console.error("Error updating registration:", err.message);
    res.status(500).send("خطأ في الخادم.");
  }
});

// ------------------------------
// Start the Server and Process the Excel File
// ------------------------------
app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
  // Replace 'attendees.xlsx' with the path to your Excel file
//   processExcelFile(path.join(__dirname, 'attendis.xlsx'));
});
