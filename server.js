const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");
const ExcelJS = require("exceljs");
const multer = require("multer");
const path = require("path");
const { v4: uuidv4 } = require("uuid");
const QRCode = require("qrcode");
const nodemailer = require("nodemailer");
const fs = require("fs");

const app = express();
app.use(cors());
app.use(bodyParser.json());
app.use("/uploads", express.static("uploads"));

const ADMIN_EMAIL = "bujji6728@gmail.com";
let registrations = [];

/* ===== Multer Setup ===== */
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, "uploads/"),
    filename: (req, file, cb) =>
        cb(null, Date.now() + path.extname(file.originalname))
});
const upload = multer({ storage });

/* ===== Email Setup ===== */
const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
        user: "bujji6728@gmail.com",
        pass: "Prasanth@19"
    }
});

/* ===== Send QR Code to Participant ===== */
async function sendQRMail(to, qrDataUrl, scanUrl) {
    await transporter.sendMail({
        from: '"Event Team" <bujji6728@gmail.com>',
        to,
        subject: "Your Event QR Code",
        html: `
            <p>Thank you for registering.</p>
            <p>Please show this QR code at the event.</p>
            <img src="${qrDataUrl}" /><br><br>
            <a href="${scanUrl}">${scanUrl}</a>
        `
    });
}

/* ===== Generate Excel + Send to Admin ===== */
async function generateExcelAndSendToAdmin() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Registrations");

    sheet.columns = [
        { header: "Team Name", key: "team", width: 20 },
        { header: "Leader Name", key: "lname", width: 20 },
        { header: "Leader Email", key: "lemail", width: 25 },
        { header: "Leader Mobile", key: "lmobile", width: 15 },
        { header: "Leader Reg No", key: "lreg", width: 18 },
        { header: "Member Names", key: "members", width: 30 },
        { header: "Transaction ID", key: "txn", width: 25 },
        { header: "Scanned", key: "scanned", width: 10 }
    ];

    registrations.forEach(t => {
        sheet.addRow({
            team: t.teamName,
            lname: t.leader.name,
            lemail: t.leader.email,
            lmobile: t.leader.mobile,
            lreg: t.leader.reg,
            members: t.members.map(m => m.name).join(", "),
            txn: t.txn,
            scanned: t.scanned ? "Yes" : "No"
        });
    });

    const filePath = path.join(__dirname, "Registrations.xlsx");
    await workbook.xlsx.writeFile(filePath);

    await transporter.sendMail({
        from: '"Event Admin" <bujji6728@gmail.com>',
        to: ADMIN_EMAIL,
        subject: "Updated Registrations Excel",
        text: "Latest participant registrations are attached.",
        attachments: [
            {
                filename: "Registrations.xlsx",
                path: filePath
            }
        ]
    });
}

/* ===== Register Team ===== */
app.post("/register", upload.single("screenshot"), async (req, res) => {
    try {
        const data = JSON.parse(req.body.data);

        // Leader + 4 members = 5
        if (data.members.length !== 4) {
            return res
                .status(400)
                .send("Team must have exactly 1 leader and 4 members");
        }

        const regId = uuidv4();
        const scanUrl = `http://localhost:3000/scan?id=${regId}`;
        const qrDataUrl = await QRCode.toDataURL(scanUrl);

        registrations.push({
            id: regId,
            teamName: data.team,
            leader: data.leader,
            members: data.members,
            txn: data.txn,
            screenshot: req.file ? req.file.filename : "",
            scanned: false
        });

        await sendQRMail(data.leader.email, qrDataUrl, scanUrl);
        await generateExcelAndSendToAdmin();

        res.json({ message: "Registration successful" });
    } catch (err) {
        console.error(err);
        res.status(500).send("Registration failed");
    }
});

/* ===== Scan QR ===== */
app.get("/scan", (req, res) => {
    const { id } = req.query;
    const reg = registrations.find(r => r.id === id);

    if (!reg) return res.send("Invalid QR code");
    if (reg.scanned) return res.send("Already scanned");

    reg.scanned = true;
    res.send("Attendance marked successfully");
});

/* ===== Download Excel (Admin) ===== */
app.get("/download", async (req, res) => {
    if (req.query.email !== ADMIN_EMAIL) {
        return res.status(403).send("Access denied");
    }

    const filePath = path.join(__dirname, "Registrations.xlsx");
    if (!fs.existsSync(filePath)) {
        return res.send("No registrations yet");
    }

    res.download(filePath);
});

/* ===== Start Server ===== */
app.listen(3000, () => {
    console.log("Server running at http://localhost:3000");
});
