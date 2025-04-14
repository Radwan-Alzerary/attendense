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
// Added field: massageSendSusses to track if the email was sent successfully.
const registrationSchema = new mongoose.Schema({
    name: String,
    email: String,
    phone: String,
    token: String,
    field: String,
    confirmed: { type: Boolean, default: false },
    massageSendSusses: { type: Boolean, default: false } // new field
});

const Registration = mongoose.model('Registration', registrationSchema);

// ------------------------------
// Setup Nodemailer Transporter
// ------------------------------
const transporter = nodemailer.createTransport({
    service: "Gmail",
    host: "smtp.gmail.com",
    port: 465,
    secure: true,
    auth: {
        user: "mdociniraqinfo@gmail.com",
        pass: "tqnl ppbs gaix joeg",
    },
    pool: true,                // Enable pooling
    maxConnections: 5,         // Limit the number of connections
    maxMessages: 100,          // Limit the number of messages per connection
    rateDelta: 5000,          // Rate period in milliseconds (1 minute)
    rateLimit: 50              // Limit the number of messages per rate period
});

// ------------------------------
// Helper Function: Send Confirmation Email
// ------------------------------
/**
 * Sends a confirmation email to a recipient.
 * Updates the registration document's `massageSendSusses` field to true when the email is sent.
 *
 * @param {Object} details - Contains name, email, and token.
 * @returns {Promise} - Resolves when the email is sent and the DB updated.
 */
function sendConfirmationEmail({ name, email, token }) {
    // Generate unique confirmation URL
    const confirmUrl = `https://attendance.niuraiq.com/confirm/${token}`;

    const mailOptions = {
        from: '"mdociniraq" <mdociniraqinfo@gmail.com>',
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
    .content h2 { color: #ffffff; font-size: 20px; margin-bottom: 15px; }
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
    <div class="content" style="display: flex;flex-direction: column;align-items: center;">
      <h2>تأكيد الحضور</h2>
      <h2>مرحبا ${name}</h2>
      <p>تحية طيبة</p>
      <p>يسرنا دعوتكم للمشاركة في الدورة الرابعة لكورس سرطانات النساء في العراق، والذي يجمع نخبة الخبراء والمتخصصين لتبادل أحدث التجارب والرؤى في مجال رعاية المرضى. إن حضوركم الكريم يُعد دليلاً على التزامكم بتعزيز مستوى الرعاية الصحية وتقديم أفضل الخدمات.</p>
      <p>نرجو منكم تأكيد حضوركم عبر النقر على الزر أدناه:</p>
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

    // Return a Promise that resolves when the email is sent and the database updated.
    return new Promise((resolve, reject) => {
        transporter.sendMail(mailOptions, async (error, info) => {
            if (error) {
                console.error(`Error sending email to ${email}:`, error);
                return reject(error);
            } else {
                console.log(`Email sent successfully to ${email}. Message ID: ${info.messageId}`);
                try {
                    // Update the record: mark message as sent successfully.
                    await Registration.updateOne({ token }, { massageSendSusses: true });
                    resolve(info);
                } catch (updateError) {
                    console.error(`Error updating send status for ${email}:`, updateError);
                    reject(updateError);
                }
            }
        });
    });
}

// ------------------------------
// Delay function for sequential processing
// ------------------------------
function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// ------------------------------
// Process the Excel File and Save Records to MongoDB
// ------------------------------
/**
 * Reads an Excel file and processes each row.
 *
 * For each row:
 *  - It checks whether the email already exists in the database.
 *  - If it exists and the email has not been sent (massageSendSusses is false), it attempts to resend the confirmation email.
 *  - If it does not exist, it creates a new registration and sends the confirmation email.
 *  - A 3-second delay is inserted between processing rows.
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

    (async function processRows() {
        for (const row of data) {
            // Adjust the property names based on your Excel file headers.
            const name = `${row.NameFirst || ''} ${row.NameLast || ''}`.trim();
            const email = row.Email || row.email;
            const phone = row.Phone || row.phone || row.Number || row.number;
            const field = row.Field || ''; // Adjust as needed

            if (!email) {
                console.warn("No email found for row, skipping.");
                continue;
            }

            try {
                // Check if the registration already exists by email.
                let registration = await Registration.findOne({ email });
                if (registration) {
                    console.log(`Record already exists for ${email}.`);
                    // If the email exists but the message hasn't been sent, resend the confirmation email.
                    if (!registration.massageSendSusses) {
                        console.log(`Resending confirmation email to ${email}...`);
                        await sendConfirmationEmail({ name: registration.name, email: registration.email, token: registration.token });
                    }
                } else {
                    // Email does not exist; create a new registration.
                    const token = crypto.randomBytes(16).toString('hex');
                    registration = new Registration({ name, email, phone, field, token });
                    await registration.save();
                    console.log(`Inserted record for ${name} with email ${email}.`);
                    await sendConfirmationEmail({ name, email, token });
                }
            } catch (err) {
                console.error(`Error processing row for email ${email}:`, err.message);
            }

            // Wait 3 seconds before processing the next record.
            await delay(3000);
        }
    })();
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

        // Mark as confirmed and save the update.
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
    // Replace 'attendis.xlsx' with the correct path to your Excel file.
    processExcelFile(path.join(__dirname, 'attendis.xlsx'));
});
