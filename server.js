const express = require('express');
const mysql = require('mysql2');
const cors = require('cors');
const bodyParser = require('body-parser');
const speakeasy = require("speakeasy");
const multer = require('multer');
const csv = require("csv-parser");
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const axios = require('axios');
const XLSX = require("xlsx");
const sql = require("mssql");
const { formatNumbersDeep } = require("./formatter");
const { formatWithMeanings } = require("./formatter");
const https = require("https");

const app = express();
const PORT = process.env.PORT || 5000;

const httpsOptions = {
  key: fs.readFileSync("C:/ssl/coastal.bank.in.key"),
  cert: fs.readFileSync("C:/ssl/coastal.bank.in.crt"),
  ca: fs.readFileSync("C:/ssl/ca_bundle.crt")
};

// LAN IP
const LAN_IP = "40.80.79.26";

// BASE URL (now HTTPS)
const BASE_URL = process.env.BASE_URL || `https://mobile.coastal.bank.in:5000`;
app.set("BASE_URL", BASE_URL);

// Middleware
app.use(cors());
app.use(bodyParser.json());


// ===========================================================
// 🪄 Global middleware to format all numeric values in Indian style (safe version)
// ===========================================================
app.use((req, res, next) => {
  const originalJson = res.json;

  res.json = function (data) {
    try {
      // ✅ Skip formatting for contact-related routes
      const skipFormatting =
        req.originalUrl.includes("get-branch-contacts") ||
        req.originalUrl.includes("admin-contacts") ||
        req.originalUrl.includes("employee-contacts") ||
        req.originalUrl.includes("get-branch-manager") ||
        req.originalUrl.includes("get-call-details");

      if (skipFormatting) {
        return originalJson.call(this, data); // 🚫 No formatting
      }

      // ✅ Otherwise, apply number formatting safely
      const clonedData = JSON.parse(JSON.stringify(data));
      const formattedData = formatNumbersDeep(clonedData);
      return originalJson.call(this, formattedData);
    } catch (err) {
      console.error("⚠️ Formatting error:", err);
      // In case of any issue, send original data
      return originalJson.call(this, data);
    }
  };

  next();
});


//========================= MSSQL (Local/UIT Server) Connection =========================//
const mssqlConfig = {
  user: "AdministratorDev",
  password: "Clab@@230830",
  server: "10.0.0.4",
  database: "MIS",
  options: {
    encrypt: false,
    trustServerCertificate: true,
  },
};

const poolPromise = new sql.ConnectionPool(mssqlConfig)
  .connect()
  .then(pool => {
    console.log("✅ Connected directly to MSSQL Database");
    return pool;
  })
  .catch(err => console.error("❌ MSSQL Connection Failed:", err));
  

//========================= Helper: Query Function =========================//
async function queryUTIDatabase(query, params = []) {
  try {
    const pool = await poolPromise;
    const request = pool.request();

    // ✅ Add parameters correctly for MSSQL
    params.forEach((value, i) => {
      request.input(`param${i + 1}`, value);
    });

    // Replace '?' with @param1, @param2, etc.
    let formattedQuery = query;
    params.forEach((_, i) => {
      formattedQuery = formattedQuery.replace('?', `@param${i + 1}`);
    });

    const result = await request.query(formattedQuery);
    return result.recordset;
  } catch (err) {
    console.error("❌ SQL Query Failed:", err);
    throw err;
  }
}

app.get("/backend", (req, res) => {
  res.send("✅ Backend server is running successfully!");
});

//========================= MSSQL MISUAT Connection =========================//
const mssqlConfigMISUAT = {
  user: "AdministratorDev",
  password: "Clab@@230830",
  server: "10.0.0.4",
  database: "MISUAT",
  options: {
    encrypt: false,
    trustServerCertificate: true,
  },
};

const poolMISUAT = new sql.ConnectionPool(mssqlConfigMISUAT)
  .connect()
  .then(pool => {
    console.log("✅ Connected to MISUAT Database");
    return pool;
  })
  .catch(err => console.error("❌ MISUAT Connection Failed:", err));
async function queryMISUAT(query, params = []) {
  const pool = await poolMISUAT;
  const request = pool.request();

  params.forEach((param, i) => {
    request.input(`param${i}`, param);
  });

  // Convert ? to @param0, @param1, etc.
  let parsedQuery = query;
  params.forEach((_, i) => {
    parsedQuery = parsedQuery.replace('?', `@param${i}`);
  });

  const result = await request.query(parsedQuery);
  return result.recordset;
}

// ===========================================================
// 🧹 Helper Function: Strip commas from numeric-like values
// ===========================================================
function stripCommas(value) {
  if (value === null || value === undefined) return value;

  // If it's already a number, leave it
  if (typeof value === "number") return value;

  // If it's a string like "1,009" or "-12,34,567.89"
  if (typeof value === "string") {
    const numericLike = /^-?[\d,]+(\.\d+)?$/.test(value.trim());
    if (numericLike) {
      // Convert safely to a number after removing commas
      const num = Number(value.replace(/,/g, ""));
      return isNaN(num) ? value : num;
    }
  }

  // For non-numeric values, return as is
  return value;
}

// ===========================================================
// ✅ Log Activity Function (Safe for SQL)
// ===========================================================
async function logActivity(userId, role, action, details) {
  try {
    // 🧹 Step 1: Clean all incoming parameters
    const cleanUserId = stripCommas(userId);
    const cleanRole = stripCommas(role);
    const cleanAction = stripCommas(action);

    let cleanDetails = details;

    // 🧹 Step 2: If details is object → stringify
    if (typeof cleanDetails === "object") {
      cleanDetails = JSON.stringify(cleanDetails);
    }

    // 🧹 Step 3: Remove commas in the details string (if any)
    if (typeof cleanDetails === "string") {
      cleanDetails = cleanDetails.replace(/,/g, "");
    }

    // 🧱 Step 4: SQL query (no formatted numbers will reach DB)
    const query = `
      INSERT INTO [dbo].[activity_logs]
      (user_id, role, action, details, created_at)
      VALUES (?, ?, ?, ?, GETDATE())
    `;

    await queryUTIDatabase(query, [
      cleanUserId,
      cleanRole,
      cleanAction,
      cleanDetails,
    ]);

    console.log(`🔹 Activity logged: ${action} by ${userId}`);
  } catch (err) {
    console.error(
      "❌ Error logging activity: UTI Database access failed",
      err.message
    );
    // ❌ Do not throw, just continue
  }
}

// ===================== NORMALIZE PHONE =====================
function normalizePhoneOrBranch(input) {
  if (!input) return "";

  // Remove non-digit characters
  let cleaned = input.replace(/\D/g, "");

  // ✅ Branch code normalization (remove leading zeros)
  if (/^0+\d+$/.test(cleaned)) {
    cleaned = String(parseInt(cleaned, 10)); // Convert to number to drop leading zeros
  }

  // ✅ Phone normalization
  if (cleaned.length === 10) return cleaned; // already a valid 10-digit phone
  if (cleaned.length === 12 && cleaned.startsWith("91")) return cleaned.slice(2); // handle 91 prefix

  return cleaned;
}


//============================================================================================
//                                    REGISTER
//=============================================================================================
app.post("/register", async (req, res) => {
  const { employeeId, phone, password, securityQuestions, deviceId } = req.body;
  console.log("📩 Register request received:", employeeId);

  if (!employeeId || !phone || !password || !deviceId) {
    return res.status(400).json({ message: "All fields are required." });
  }

  if (
    !Array.isArray(securityQuestions) ||
    securityQuestions.length !== 3 ||
    securityQuestions.some((q) => !q.question || !q.answer)
  ) {
    return res.status(400).json({
      message: "Exactly 3 security questions with answers are required.",
    });
  }

  const normalizedPhone = normalizePhoneOrBranch(phone);

  try {
    // 1️⃣ Verify employee exists
    const empQuery = `
      SELECT Password, device_id, GA_Secret
      FROM [dbo].[employees]
      WHERE [Emp No.] = ? AND [Mobile number] = ?
    `;

    const empRows = await queryUTIDatabase(empQuery, [
      employeeId,
      normalizedPhone,
    ]);

    if (empRows.length === 0) {
      return res.status(404).json({
        message: "Employee ID or Phone not found.",
      });
    }

    const emp = empRows[0];

    // 🚫 Already registered (same employee)
    if (emp.Password || emp.GA_Secret) {
      return res.status(409).json({
        message:
          "Registration already completed for this employee. Contact admin if reset is required.",
      });
    }

    // 🚫 Device already bound to another employee
    const deviceCheckQuery = `
      SELECT [Emp No.]
      FROM [dbo].[employees]
      WHERE device_id = ?
    `;

    const deviceRows = await queryUTIDatabase(deviceCheckQuery, [deviceId]);

    if (deviceRows.length > 0) {
      return res.status(409).json({
        message:
          "This device is already registered to another employee. Only one registration is allowed per device.",
      });
    }

    // 2️⃣ Generate GA secret ONCE
    const secret = speakeasy.generateSecret({ length: 32 });

    // 3️⃣ Save registration
    const updateQuery = `
      UPDATE [dbo].[employees]
      SET
        [Password] = ?,
        [Approval status] = 'pending',
        [GA_Secret] = ?,
        [security_q1] = ?, [security_a1] = ?,
        [security_q2] = ?, [security_a2] = ?,
        [security_q3] = ?, [security_a3] = ?,
        [device_id] = ?
      WHERE [Emp No.] = ? AND [Mobile number] = ?
    `;

    await queryUTIDatabase(updateQuery, [
      password,
      secret.base32,
      securityQuestions[0].question,
      securityQuestions[0].answer,
      securityQuestions[1].question,
      securityQuestions[1].answer,
      securityQuestions[2].question,
      securityQuestions[2].answer,
      deviceId,
      employeeId,
      normalizedPhone,
    ]);

    logActivity(employeeId, "User", "Register", `Device:${deviceId}`);

    return res.status(200).json({
      message: "Registration successful. Awaiting admin approval.",
      googleAuthKey: secret.base32,
    });
  } catch (err) {
    console.error("❌ /register error:", err);
    return res.status(500).json({
      message: "Server error during registration.",
    });
  }
});

//============================================================================================
//                                           LOGIN
//=============================================================================================
//========================= /login Route =========================//
app.post("/login", async (req, res) => {
  try {
    const { employeeId, password, deviceId, token } = req.body;

    if (!employeeId || !password || !deviceId || !token) {
      return res.status(400).json({ message: "All fields are required." });
    }

    const results = await queryUTIDatabase(`
      SELECT * FROM [dbo].[employees]
      WHERE [Emp No.] = ? 
    `, [employeeId]);

    // INVALID EMPLOYEE ID
    if (results.length === 0) {
      logActivity(
  String(employeeId || "UNKNOWN"),
  "Unknown",
  "Login Failed",
  `Employee ${employeeId} does not exist`
);
      return res.status(404).json({ message: "Employee not found." });
    }

    const user = results[0];
    const role = user.Role || user.role || "user";
    const userId = user["Emp No."].toString();

    // ACCOUNT LOCKED
    if (user.account_locked) {
      logActivity(userId, role, "Login Failed", "Account locked");
      return res.status(403).json({ message: "Account locked. Contact admin." });
    }

    // ADMIN NOT APPROVED
    if (user["Approval status"] !== "approved") {
      logActivity(userId, role, "Login Failed", "Account not approved");
      return res.status(403).json({ message: "Account not approved by admin yet." });
    }

    // PASSWORD RESET REQUIRED
    if (user.reset_required) {
      logActivity(userId, role, "Login Failed", "Password reset required");
      return res.status(403).json({ message: "Password reset required before login." });
    }

    // WRONG PASSWORD
    if (user.Password !== password) {
      const newAttempts = (user.failed_attempts || 0) + 1;

      await queryUTIDatabase(
        "UPDATE [dbo].[employees] SET failed_attempts = ? WHERE [Emp No.] = ?",
        [newAttempts, employeeId]
      );

      logActivity(userId, role, "Login Failed", "Invalid password");

      if (newAttempts >= 3) {
        await queryUTIDatabase(
          "UPDATE [dbo].[employees] SET reset_required = 1 WHERE [Emp No.] = ?",
          [employeeId]
        );

        logActivity(userId, role, "Login Failed", "Account locked due to 3 failed attempts");
        return res.status(403).json({
          message: "Too many failed attempts. Password reset required.",
        });
      }

      return res.status(401).json({ message: "Invalid password." });
    }

    // WRONG GOOGLE AUTH TOKEN
    const verified = speakeasy.totp.verify({
      secret: user.GA_Secret,
      encoding: "base32",
      token,
      window: 1,
    });

    if (!verified) {
      logActivity(userId, role, "Login Failed", "Invalid Google Authenticator code");
      return res.status(401).json({ message: "Invalid Google Authenticator code." });
    }

    // WRONG DEVICE
    if (!user.device_id) {
      await queryUTIDatabase(
        "UPDATE [dbo].[employees] SET device_id = ? WHERE [Emp No.] = ?",
        [deviceId, employeeId]
      );
    } else if (user.device_id !== deviceId) {
      logActivity(userId, role, "Login Failed", "Device mismatch");
      return res.status(403).json({
        message: "Login denied: Account bound to another device.",
      });
    }

    // SUCCESS LOGIN
    await queryUTIDatabase(
      "UPDATE [dbo].[employees] SET failed_attempts = 0 WHERE [Emp No.] = ?",
      [employeeId]
    );

    logActivity(
      userId,
      role,
      "Login",
      `User ${user['Employee Name']} logged in`
    );

    return res.status(200).json({
      success: true,
      message: `Welcome, ${user["Employee Name"] || "User"}!`,
      role,
      userId,
      level: user.Level,
      branchCode: user["Br Code"],
      designation: user.Designation,
      employeeName: user["Employee Name"],
      clusterName:
        user.Level === "Level 2" && user.Designation?.includes("Cluster Head-")
          ? user.Designation.replace("Cluster Head-", "").trim()
          : null,
    });

  } catch (err) {
    console.error("❌ /login error:", err);

    // Log server error too
    logActivity("UNKNOWN", "Unknown", "Login Failed", `Server error: ${err.message}`);

    return res.status(500).json({
      message: "Server error during login.",
      error: err.message,
    });
  }
});


//============================================================================================
//                                          LOGOUT ENDPOINT
//=============================================================================================

app.post("/logout", async (req, res) => {
  const { userId, role, lastActivity } = req.body;

  try {
    await logActivity(
      userId,
      role,
      "Logout",
      `Last activity: ${lastActivity || "N/A"}`
    );

    res.json({ success: true, message: "Logged out successfully" });
  } catch (err) {
    console.error("❌ Logout error:", err.message);
    res.status(500).json({ success: false, message: "Logout failed" });
  }
});

//============================================================================================
//                                 FORGOT PASSWORD
//=============================================================================================

// ✅ Forgot Password - Step 1: Send back user's stored security questions
// ✅ Step 1: Forgot Password Start — Fetch Security Questions
app.post("/forgot-password/start", async (req, res) => {
  try {
    const { employeeId, phone } = req.body;

    if (!employeeId || !phone) {
      return res
        .status(400)
        .json({ message: "Employee ID and phone are required." });
    }

    const query = `
      SELECT [security_q1], [security_q2], [security_q3], [Role]
      FROM [dbo].[employees]
      WHERE [Emp No.] = ? AND [Mobile number] = ?
    `;
    const results = await queryUTIDatabase(query, [employeeId, phone]);

    if (results.length === 0) {
      return res.status(404).json({ message: "Employee not found" });
    }

    const user = results[0];
    const role = user.Role || "user";

    logActivity(
      employeeId,
      role,
      "Forgot Password Start",
      "Requested security questions"
    );

    // Utility: Map IDs (q1..q10) to text
    function mapQuestion(id) {
      const questionBank = {
        q1: "What was the name of your first school?",
        q2: "What is the name of your favorite childhood teacher?",
        q3: "What is the title of your favorite childhood book?",
        q4: "What is the name of your favorite childhood friend?",
        q5: "What was the make of your first bicycle?",
        q6: "Where did you go on your first school trip?",
        q7: "What was the first film you saw in a theater?",
        q8: "What is the name of your first pet?",
        q9: "What is your favorite subject in school?",
        q10: "What was the nickname given to you in childhood?",
      };
      return questionBank[id] || "Unknown question";
    }

    const securityQuestions = [
      { id: "q1", text: mapQuestion(user.security_q1) },
      { id: "q2", text: mapQuestion(user.security_q2) },
      { id: "q3", text: mapQuestion(user.security_q3) },
    ];

    res.json({ questions: securityQuestions });
  } catch (err) {
    console.error("❌ Error in /forgot-password/start:", err);
    res.status(500).json({ message: "Server error" });
  }
});


// ✅ Step 2: Verify Security Answers
app.post("/verify-security-answers", async (req, res) => {
  try {
    const { employeeId, phone, answers } = req.body;

    const query = `
      SELECT [security_a1], [security_a2], [security_a3], [reset_attempts],
             [account_locked], [reset_required], [lock_until], [Role]
      FROM [dbo].[employees]
      WHERE [Emp No.] = ? AND [Mobile number] = ?
    `;
    const results = await queryUTIDatabase(query, [employeeId, phone]);

    if (results.length === 0)
      return res.status(404).json({ message: "Employee not found." });

    const user = results[0];
    const role = user.Role || "user";

    // 🔹 Account locked check
    if (user.account_locked) {
      logActivity(
        employeeId,
        role,
        "Forgot Password Attempt",
        "Account locked, cannot verify answers"
      );
      return res
        .status(403)
        .json({ message: "Account locked. Contact admin to unlock." });
    }

    // 🔹 Cooldown check
    if (
      !user.reset_required &&
      user.lock_until &&
      new Date(user.lock_until) > new Date()
    ) {
      const waitMinutes = Math.ceil(
        (new Date(user.lock_until) - new Date()) / (60 * 1000)
      );
      return res.status(403).json({
        message: `❌ Too many wrong attempts. Please try again after ${waitMinutes} minutes.`,
      });
    }

    // 🔹 Check answers (case-insensitive)
    let correctCount = 0;
    if (answers.q1?.toLowerCase() === user.security_a1?.toLowerCase())
      correctCount++;
    if (answers.q2?.toLowerCase() === user.security_a2?.toLowerCase())
      correctCount++;
    if (answers.q3?.toLowerCase() === user.security_a3?.toLowerCase())
      correctCount++;

    if (correctCount >= 2) {
      await queryUTIDatabase(
        "UPDATE [dbo].[employees] SET reset_attempts=0, lock_until=NULL WHERE [Emp No.]=?",
        [employeeId]
      );
      logActivity(
        employeeId,
        role,
        "Forgot Password Verified",
        "Security answers verified successfully"
      );
      return res.status(200).json({ message: "✅ Verification successful" });
    }

    // ❌ Wrong answers
    const newAttempts = (user.reset_attempts || 0) + 1;
    await queryUTIDatabase(
      "UPDATE [dbo].[employees] SET reset_attempts=? WHERE [Emp No.]=?",
      [newAttempts, employeeId]
    );
    logActivity(
      employeeId,
      role,
      "Forgot Password Failed Verification",
      `Wrong answers: ${JSON.stringify(answers)}`
    );

    if (user.reset_required && newAttempts >= 3) {
      await queryUTIDatabase(
        "UPDATE [dbo].[employees] SET account_locked=1 WHERE [Emp No.]=?",
        [employeeId]
      );
      return res.status(403).json({
        message: "❌ Too many wrong answers. Account locked. Contact admin.",
      });
    }

    if (!user.reset_required && newAttempts >= 3) {
      const lockUntil = new Date(Date.now() + 15 * 60 * 1000); // +15 min
      await queryUTIDatabase(
        "UPDATE [dbo].[employees] SET reset_attempts=?, lock_until=? WHERE [Emp No.]=?",
        [newAttempts, lockUntil, employeeId]
      );
      return res.status(403).json({
        message: "❌ Too many wrong attempts. Try again after 15 minutes.",
      });
    }

    return res
      .status(401)
      .json({ message: `❌ Wrong answers. Attempts left: ${3 - newAttempts}` });
  } catch (err) {
    console.error("❌ Error in /verify-security-answers:", err);
    res.status(500).json({ message: "Server error." });
  }
});


// ✅ Step 3: Reset Password
app.post("/forgot-password/reset", async (req, res) => {
  try {
    const { employeeId, phone, newPassword } = req.body;

    if (!employeeId || !phone || !newPassword) {
      return res.status(400).json({ message: "All fields required." });
    }

    const getUserQuery = `
      SELECT [Role]
      FROM [dbo].[employees]
      WHERE [Emp No.] = ? AND [Mobile number] = ?
    `;
    const userResults = await queryUTIDatabase(getUserQuery, [
      employeeId,
      phone,
    ]);

    if (userResults.length === 0) {
      return res.status(404).json({ message: "Employee not found." });
    }

    const role = userResults[0].Role || "user";

    const updateQuery = `
      UPDATE [dbo].[employees]
      SET [Password] = ?, [reset_required] = 0, [reset_attempts] = 0, [lock_until] = NULL
      WHERE [Emp No.] = ? AND [Mobile number] = ?
    `;
    await queryUTIDatabase(updateQuery, [newPassword, employeeId, phone]);

    logActivity(
      employeeId,
      role,
      "Password Reset",
      "Successfully reset password"
    );

    return res.status(200).json({
      message: "✅ Password reset successful. You can login now.",
    });
  } catch (err) {
    console.error("❌ Error in /forgot-password/reset:", err);
    res
      .status(500)
      .json({ message: "Server error while resetting password." });
  }
});

//============================================================================================
//                                       CIRCULAR MANAGEMENT
//=============================================================================================

/**
 * POST /circulars
 * Body: { empNo, title, message }
 */
app.post("/circulars", async (req, res) => {
  console.log("📩 Incoming request: /circulars");

  try {
    let empNo = parseInt(String(req.body.empNo).split(".")[0]);  // FIX
    let title = req.body.title;
    let message = req.body.message;

    if (!empNo || !title || !message) {
      return res.status(400).json({ message: "empNo, title and message required" });
    }

    const insertCircularSQL = `
  INSERT INTO circulars (emp_no, title, message, type, created_at)
  VALUES (@param1, @param2, @param3, @param4, GETDATE());
`;


    await queryUTIDatabase(insertCircularSQL, [
  empNo,
  title,
  message,
  "text"      // 👈 DEFAULT TYPE
]);

    // Log admin action
    const logSQL = `
      INSERT INTO activity_logs (user_id, role, action, details, created_at)
      VALUES (@param1, @param2, @param3, @param4, GETDATE());
    `;

    await queryUTIDatabase(logSQL, [
      empNo,
      "admin",
      "Create Circular",
      `Title: ${title}`
    ]);

    res.json({ ok: true, message: "Circular created successfully" });
  } catch (err) {
    console.error("❌ Error creating circular:", err);
    res.status(500).json({ message: "Error creating circular", error: err.message });
  }
});


/**
 * POST /circulars/delete
 * Body: { circularId, empNo }
 */
app.post("/circulars/delete", async (req, res) => {
  try {
    let circularId = parseInt(req.body.circularId);
    let empNo = String(req.body.empNo || "").replace(/\.0+$/, "");

    if (!circularId || !empNo) {
      return res.status(400).json({ message: "circularId and empNo required" });
    }

    // Check if exists
    const checkQuery = `SELECT id FROM circulars WHERE id = @param1`;
    const rows = await queryUTIDatabase(checkQuery, [circularId]);

    if (!rows.length) {
      return res.status(404).json({ message: "Circular not found" });
    }

    // Delete
    const deleteQuery = `DELETE FROM circulars WHERE id = @param1`;
    await queryUTIDatabase(deleteQuery, [circularId]);

    // Log deletion
    const logQuery = `
      INSERT INTO activity_logs (user_id, role, action, details, created_at)
      VALUES (@param1, 'admin', 'Delete Circular', @param2, GETDATE())
    `;
    await queryUTIDatabase(logQuery, [
      empNo,
      `Deleted circular ID: ${circularId}`
    ]);

    return res.json({ success: true, message: "Circular deleted successfully" });
  } catch (err) {
    console.error("❌ Error deleting circular:", err);
    return res.status(500).json({ message: "Error deleting circular", error: err.message });
  }
});



/**
 * POST /circulars/fetch
 * Body: { empNo }
 * Returns only pending (unacknowledged) circulars for an employee
 */
app.post("/circulars/fetch", async (req, res) => {

  try {
    const { empNo } = req.body;

    if (!empNo) {
      return res.status(400).json({ message: "empNo required" });
    }
	
	const cleanEmpNo = parseInt(empNo);


    // ✅ MSSQL query (GETDATE instead of NOW, ? replaced with @param1)
    const sql = `
      SELECT c.id, c.title, c.message, c.created_at
      FROM circulars c
      WHERE NOT EXISTS (
        SELECT 1 FROM circular_acknowledgements ca
        WHERE ca.circular_id = c.id AND ca.emp_no = @param1
      )
      ORDER BY c.created_at DESC;
    `;

    const rows = await queryUTIDatabase(sql, [cleanEmpNo]);

    console.log(`✅ ${rows.length} circulars fetched for empNo: ${empNo}`);
    return res.json(rows);
  } catch (err) {
    console.error("❌ Error fetching circulars:", err);
    return res.status(500).json({ message: "Error fetching circulars" });
  }
});

/**
 * POST /circulars/acknowledge
 * Body: { circularId, empNo }
 */
app.post("/circulars/acknowledge", async (req, res) => {
  try {
    const { circularId, empNo } = req.body;

    if (!circularId || !empNo) {
      return res.status(400).json({ message: "Missing circularId or empNo" });
    }

    // CLEAN both values
    const cleanCircularId = parseInt(circularId);
    const cleanEmpNo = parseInt(empNo);

    if (isNaN(cleanCircularId) || isNaN(cleanEmpNo)) {
      return res.status(400).json({ message: "Invalid circularId or empNo" });
    }

    // MERGE UPSERT
    const sql = `
      MERGE INTO circular_acknowledgements AS target
      USING (SELECT @param1 AS circular_id, @param2 AS emp_no) AS source
      ON (target.circular_id = source.circular_id AND target.emp_no = source.emp_no)
      WHEN MATCHED THEN
        UPDATE SET acknowledged_at = GETDATE()
      WHEN NOT MATCHED THEN
        INSERT (circular_id, emp_no, acknowledged_at)
        VALUES (source.circular_id, source.emp_no, GETDATE());
    `;

    await queryUTIDatabase(sql, [cleanCircularId, cleanEmpNo]);

    return res.json({ success: true });

  } catch (err) {
    console.error("❌ Error acknowledging circular:", err);
    return res.status(500).json({ message: "Error acknowledging circular" });
  }
});


/**
 * POST /circulars/history
 * Admin endpoint: Returns all circulars with acknowledgement count
 * Body: { adminId }
 */
app.post("/circulars/history", async (req, res) => {
  try {
    const { empNo } = req.body || {};
	const cleanEmpNo = parseInt(empNo || 0);


    const sql = `
      SELECT
        c.id,
        c.title,
        c.message,
        c.created_at,
        c.emp_no,
        e.[Employee Name] AS uploader_name,
        (
          SELECT COUNT(*) 
          FROM circular_acknowledgements ca 
          WHERE ca.circular_id = c.id
        ) AS acknowledged_count
      FROM circulars c
      LEFT JOIN employees e ON e.[Emp No.] = c.emp_no
      ORDER BY c.created_at DESC;
    `;

    const result = await queryUTIDatabase(sql);

    if (cleanEmpNo) {
      const logSQL = `
        INSERT INTO activity_logs (user_id, role, action, details, created_at)
        VALUES (@param1, 'admin', 'View Circulars History', @param2, GETDATE());
      `;
      await queryUTIDatabase(logSQL, [
        cleanEmpNo,
        `Fetched ${result.length} circulars`,
      ]);
    }

    return res.json(result);
  } catch (err) {
    console.error("❌ Error fetching circulars history:", err);
    return res.status(500).json({
      message: "Error fetching circular history",
      error: err.message,
    });
  }
});



// ===========================================================================================================
//                                            USER STATS
// ===========================================================================================================
app.post("/user-stats", async (req, res) => {
  try {
    const query = `
      SELECT
        CAST(SUM(CASE WHEN [Approval status] = 'approved' THEN 1 ELSE 0 END) AS INT) AS approved,
        CAST(SUM(CASE WHEN [Approval status] = 'pending' AND Password IS NOT NULL THEN 1 ELSE 0 END) AS INT) AS pending,
        CAST(SUM(CASE WHEN [Approval status] = 'rejected' THEN 1 ELSE 0 END) AS INT) AS rejected,
        CAST(SUM(CASE WHEN account_locked = 1 THEN 1 ELSE 0 END) AS INT) AS locked,
        CAST(COUNT(*) AS INT) AS total_registered,
        CAST(
          (SELECT COUNT(*) FROM [dbo].[employees] WHERE Password IS NULL OR Password = '')
          AS INT
        ) AS not_registered
      FROM [dbo].[employees]
    `;

    const rows = await queryUTIDatabase(query, []);
    const row = rows[0];

    return res.status(200).json({
      approved: Number(row.approved),
      pending: Number(row.pending),
      rejected: Number(row.rejected),
      locked: Number(row.locked),
      total_registered: Number(row.total_registered),
      not_registered: Number(row.not_registered)
    });

  } catch (err) {
    console.error("❌ Error fetching user stats:", err);
    return res.status(500).json({ message: "Server error while fetching user stats" });
  }
});


// ======================================================================================================
//                                ✅ Get Users (with pagination & status)
// ======================================================================================================
app.post("/get-users", async (req, res) => {
  try {
    const { status, page = 1, limit = 10 } = req.body;
    const offset = (page - 1) * limit;

    let whereClause = "";

if (status) {

  if (status === "pending") {
    whereClause = " WHERE [Approval status] = 'pending' AND Password IS NOT NULL AND Password <> ''";

  } else if (status === "approved") {
    whereClause = " WHERE [Approval status] = 'approved'";

  } else if (status === "rejected") {
    whereClause = " WHERE [Approval status] = 'rejected'";

  } else if (status === "locked") {
    whereClause = " WHERE account_locked = 1";

  } else if (status === "not_registered") {
    // ⭐ FIX: Only fetch users whose Password is NULL or empty
    whereClause = " WHERE (Password IS NULL OR Password = '')";
  }
}

    const countQuery = `SELECT COUNT(*) AS total FROM [dbo].[employees] ${whereClause}`;
    const totalResult = await queryUTIDatabase(countQuery, []);
    const total = totalResult[0]?.total || 0;

    const usersQuery = `
      SELECT * FROM [dbo].[employees]
      ${whereClause}
      ORDER BY [Emp No.]
      OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY
    `;
    const results = await queryUTIDatabase(usersQuery, []);

    res.json({
      total,
      page: parseInt(page),
      limit: parseInt(limit),
      totalPages: Math.ceil(total / limit),
      users: results,
    });
  } catch (err) {
    console.error("❌ Error fetching users (UTI):", err);
    return res.status(500).json({ message: "Server error while fetching users" });
  }
});

// =============================================================================================
//                                     ✅ Update User Status
// =============================================================================================
app.post("/update-user-status", async (req, res) => {
  try {
    let { empNo, status, adminId, adminRole } = req.body;

    if (!empNo || !status) {
      return res.status(400).json({ message: "Employee number and status are required" });
    }

    // Convert empNo to integer (removes commas like 1,011)
    const cleanEmpNo = parseInt(String(empNo).replace(/[, ]/g, ""), 10);

    if (isNaN(cleanEmpNo)) {
      return res.status(400).json({ message: "Invalid Employee Number" });
    }

    // -------------------------
    // 1️⃣ RUN UPDATE QUERY
    // -------------------------
    const updateQuery = `
      UPDATE [dbo].[employees]
      SET [Approval status] = ?
      WHERE [Emp No.] = ?
    `;

    await queryUTIDatabase(updateQuery, [status, cleanEmpNo]);

    // -------------------------
    // 2️⃣ VERIFY UPDATE SUCCESS USING SELECT
    // -------------------------
    const checkQuery = `
      SELECT [Emp No.], [Approval status]
      FROM [dbo].[employees]
      WHERE [Emp No.] = ?
    `;

    const rows = await queryUTIDatabase(checkQuery, [cleanEmpNo]);

    if (!rows || rows.length === 0) {
      return res.status(404).json({
        message: `❌ Employee ${cleanEmpNo} not found — record not updated`
      });
    }

    // (Optional) verify that status actually updated
    const updatedStatus = rows[0]["Approval status"];
    if (updatedStatus !== status) {
      return res.status(500).json({
        message: `❌ Update failed — status did not change`
      });
    }

    // -------------------------
    // 3️⃣ LOG ACTIVITY
    // -------------------------
    logActivity(
      adminId || "system",
      adminRole || "admin",
      "Update User Status",
      `Set user ${cleanEmpNo} to ${status}`
    );

    return res.json({ message: `User ${status} successfully` });

  } catch (err) {
    console.error("❌ Error updating user status (UTI):", err);
    return res.status(500).json({ message: "Server error while updating user status" });
  }
});


// ===================================================================================================
//                                                ✅ Unlock User
// ===================================================================================================
// Utility to remove commas, decimals, and formatting
// Utility to remove commas, decimals, and formatting
function cleanNumber(value) {
  if (!value) return null;
  return value.toString().replace(/,/g, "").replace(/\.00$/, "").trim();
}

app.post("/unlock-user", async (req, res) => {
  try {
    let { empNo, adminId, adminRole } = req.body;

    if (!empNo) {
      return res.status(400).json({
        success: false,
        message: "Employee number is required",
      });
    }

    // Clean formatting "1,009.00" → "1009"
    const cleanEmpNo = cleanNumber(empNo);
    const cleanAdminId = cleanNumber(adminId);

    const query = `
      UPDATE [dbo].[employees]
      SET failed_attempts = 0,
          account_locked = 0,
          attempts = 0,
          reset_attempts = 0,
          lock_until = NULL,
          reset_required = 0
      WHERE [Emp No.] = ?
    `;

    const result = await queryUTIDatabase(query, [cleanEmpNo]);

    // 🛑 SAFETY CHECK — result must exist
    if (!result) {
      console.error("❌ Unlock-user: queryUTIDatabase returned undefined");
      return res.status(500).json({
        success: false,
        message: "Database error: No response received",
      });
    }

    console.log("Unlock-user SQL result:", result);

    // 🛑 SAFETY: rowsAffected may not exist in your wrapper
    const rows =
      result.rowsAffected?.[0] ??
      result.rowsAffected ??
      result.affectedRows ??
      0;

    if (rows > 0) {
      // Log activity safely (never crash main route)
      try {
        logActivity(
          cleanAdminId || "system",
          adminRole || "admin",
          "Unlock User",
          `Unlocked user ${cleanEmpNo}`
        );
      } catch (logErr) {
        console.error("⚠ logActivity failed:", logErr);
      }

      return res.json({
        success: true,
        message: "User unlocked successfully",
      });
    }

    // If update returned 0 rows
    return res.status(404).json({
      success: false,
      message: "User not found",
    });

  } catch (err) {
    console.error("❌ Unlock user error (UTI):", err);
    return res.status(500).json({
      success: false,
      message: "Server error while unlocking user",
    });
  }
});


// --------------------------------------------------------------------------------------------
//          FETCH DISTINCT DISTRICTS & CLUSTERS (Mapping-based, NOT Hardcoded)
//---------------------------------------------------------------------------------------------

app.post("/get-districts-clusters", async (req, res) => {
  try {
    const { level, designation } = req.body;

    const section = "Deposits";   // ALWAYS safe source
    console.log("🔍 Loading districts/clusters from section:", section);

    // 1️⃣ Load meaning mapping
    const mapping = await queryUTIDatabase(`
      SELECT col_number, meaning
      FROM MIS.dbo.MIS_Column_Mapping
      WHERE section = ?
    `, [section]);

    // Identify which column numbers are District & Cluster
    const districtMap = mapping.find(m => m.meaning === "district");
    const clusterMap  = mapping.find(m => m.meaning === "cluster");

    const districtCol = districtMap ? `col${districtMap.col_number}` : `col3`;
    const clusterCol  = clusterMap ? `col${clusterMap.col_number}` : `col4`;

    console.log("📌 District Column:", districtCol);
    console.log("📌 Cluster Column:", clusterCol);

    // 2️⃣ Base queries
    let baseDistrictQuery = `
      SELECT DISTINCT LTRIM(RTRIM([${districtCol}])) AS District
      FROM [dbo].[Deposits]
      WHERE [${districtCol}] IS NOT NULL AND LTRIM(RTRIM([${districtCol}])) <> ''
    `;

    let baseClusterQuery = `
      SELECT DISTINCT LTRIM(RTRIM([${clusterCol}])) AS Cluster
      FROM [dbo].[Deposits]
      WHERE [${clusterCol}] IS NOT NULL AND LTRIM(RTRIM([${clusterCol}])) <> ''
    `;

    let params = [];

    // 3️⃣ Restrict for Level 2 (Cluster Head)
    let effectiveCluster = null;

    if (level === "Level 2" && designation?.includes("Cluster Head-")) {
      effectiveCluster = designation.replace("Cluster Head-", "").trim();

      baseDistrictQuery += ` AND LTRIM(RTRIM([${clusterCol}])) = ?`;
      baseClusterQuery  += ` AND LTRIM(RTRIM([${clusterCol}])) = ?`;

      params.push(effectiveCluster);
    }

    // 4️⃣ Execute Queries
    const districtsRaw = await queryUTIDatabase(baseDistrictQuery, params);
    const clustersRaw  = await queryUTIDatabase(baseClusterQuery, params);

    const districts = districtsRaw
      .map(d => d.District?.trim())
      .filter(Boolean);

    const clusters = clustersRaw
      .map(c => c.Cluster?.trim())
      .filter(c => c && c.toUpperCase() !== "CO");

    // 5️⃣ Return
    return res.status(200).json({
      success: true,
      districts,
      clusters,
      fixedCluster: effectiveCluster
    });

  } catch (err) {
    console.error("❌ Error in get-districts-clusters:", err);
    return res.status(500).json({
      success: false,
      message: "Error while fetching districts & clusters",
      error: err.message
    });
  }
});
// ============================================================================================
//                        DEPOSITS (Dynamic + Meaning-Based Sorting)
// ============================================================================================

app.post("/get-deposits", async (req, res) => {
  try {
    let {
      branchCode,
      branchName,
      clusterName,
      districtName,
      userId,
      role,
      level,
      designation,
      page = 1,
      pageSize = 10,
      sortBy,                 // meaning (e.g., "gdm_deposits")
      sortOrder = "ASC",
    } = req.body;

    const SECTION = "Deposits";

    // ---------------------------------------------
    // Normalize inputs
    // ---------------------------------------------
    branchCode = branchCode !== undefined && branchCode !== null
      ? branchCode.toString().trim()
      : "";
    branchName = (branchName || "").trim();
    districtName = (districtName || "").trim();
    clusterName = (clusterName || "").trim();
    level = level || "Level 1";

    // ✅ LEVEL-1 FIX: branchCode like "2.00" → "2"
    if (level === "Level 1" && branchCode) {
      const numeric = parseFloat(branchCode);
      if (!Number.isNaN(numeric)) {
        branchCode = parseInt(numeric, 10).toString();
      }
    }

    page = parseInt(page);
    pageSize = parseInt(pageSize);

    const offset = (page - 1) * pageSize;

    // ============================================================================================
    // 1️⃣ LOAD COLUMN MAPPING (meaning → colX)
    // ============================================================================================
    const mappingRows = await queryUTIDatabase(
      `
      SELECT col_number, display_label, meaning, detected_type
      FROM MIS.dbo.MIS_Column_Mapping
      WHERE section = ?
      ORDER BY col_number ASC
      `,
      [SECTION]
    );

    let mapping = mappingRows.map((r) => ({
      colNumber: Number(r.col_number),
      colName: `col${r.col_number}`,
      displayLabel: r.display_label,
      meaning: r.meaning,
      type: r.detected_type || "string",
    }));

    // fallback minimal mapping (very rare)
    if (!mapping.length) {
      mapping = [];
      for (let i = 1; i <= 30; i++) {
        mapping.push({
          colNumber: i,
          colName: `col${i}`,
          displayLabel: `col${i}`,
          meaning:
            i === 1 ? "branch_code"
            : i === 2 ? "branch_name"
            : i === 3 ? "district"
            : i === 4 ? "cluster"
            : null,
        });
      }
    }

    // Build meaning lookup
    const meaningToCol = {};
    mapping.forEach((m) => {
      if (m.meaning) meaningToCol[m.meaning] = m.colName;
    });

    const fallback = {
      branch_code: "col1",
      branch_name: "col2",
      district: "col3",
      cluster: "col4",
    };

    const getCol = (meaning) => meaningToCol[meaning] || fallback[meaning];

    const colList = mapping.map((m) => `[${m.colName}]`).join(",");

    // ============================================================================================
    // 2️⃣ LOAD SORTABLE COLUMNS (meaning-only list)
    // ============================================================================================
    const sortableRows = await queryUTIDatabase(
      `
      SELECT display_label, meaning
      FROM MIS.dbo.MIS_Sortable_Columns
      WHERE section = ?
      ORDER BY display_label ASC
      `,
      [SECTION]
    );

    const allowedSortableMeaning = sortableRows
      .map((r) => r.meaning)
      .filter((m) => m);

    // ============================================================================================
    // 3️⃣ BUILD BASE QUERY (with filters)
    // ============================================================================================
    let query = `SELECT ${colList} FROM MIS.dbo.[Deposits] WHERE 1=1`;
    let countQuery = `SELECT COUNT(*) AS totalRecords FROM MIS.dbo.[Deposits] WHERE 1=1`;

    const params = [];
    const countParams = [];

    const addFilter = (meaning, value, exact = true) => {
      if (value === undefined || value === null || value === "") return;

      const col = getCol(meaning);
      if (!col) return;

      // ✅ SPECIAL CASE: branch_code compared NUMERICALLY (handles "2", "2.0", "2.00")
      if (meaning === "branch_code" && exact) {
        query += ` AND TRY_CAST(LTRIM(RTRIM([${col}])) AS INT) = TRY_CAST(? AS INT)`;
        countQuery += ` AND TRY_CAST(LTRIM(RTRIM([${col}])) AS INT) = TRY_CAST(? AS INT)`;
        params.push(value);
        countParams.push(value);
        return;
      }

      if (exact) {
        query += ` AND LTRIM(RTRIM([${col}])) = ?`;
        countQuery += ` AND LTRIM(RTRIM([${col}])) = ?`;
        params.push(value.trim());
        countParams.push(value.trim());
      } else {
        query += ` AND LOWER(LTRIM(RTRIM([${col}]))) LIKE ?`;
        countQuery += ` AND LOWER(LTRIM(RTRIM([${col}]))) LIKE ?`;
        params.push(`%${value.trim().toLowerCase()}%`);
        countParams.push(`%${value.trim().toLowerCase()}%`);
      }
    };

    // ============================================================================================
    // 4️⃣ LEVEL RESTRICTIONS (Meaning-based RLS)
    // ============================================================================================
    let effectiveCluster = clusterName;

    if (level === "Level 2" && designation) {
      const m = designation.match(/Cluster\s*Head\s*[-:]\s*(.*)/i);
      if (m && m[1]) effectiveCluster = m[1].trim();
    }

    if (level === "Level 1" && branchCode) {
      addFilter("branch_code", branchCode);
    }

    if (level === "Level 2") {
      if (!effectiveCluster)
        return res.status(400).json({
          success: false,
          message: "Cluster missing for Level 2 user",
        });

      addFilter("cluster", effectiveCluster);
    }

    // Normal filters
    if (branchCode && level !== "Level 1") addFilter("branch_code", branchCode);
    if (branchName && level !== "Level 1") addFilter("branch_name", branchName, false);
    if (districtName) addFilter("district", districtName, false);
    if (clusterName && level !== "Level 2") addFilter("cluster", clusterName, false);

    // ============================================================================================
    // 5️⃣ MEANING-BASED SORTING (FINAL)
    // ============================================================================================
    let orderCol = getCol("branch_code");
    let orderDir = sortOrder.toUpperCase() === "DESC" ? "DESC" : "ASC";

    if (sortBy && allowedSortableMeaning.includes(sortBy)) {
      const col = getCol(sortBy);
      if (col) orderCol = col;
    }

    // ✅ sort numerically by branch_code (and other numeric cols)
    query += ` ORDER BY TRY_CAST(LTRIM(RTRIM([${orderCol}])) AS INT) ${orderDir}`;

    // ============================================================================================
    // 6️⃣ PAGINATION
    // ============================================================================================
    query += " OFFSET ? ROWS FETCH NEXT ? ROWS ONLY";
    params.push(offset, pageSize);

    const dbRows = await queryUTIDatabase(query, params);

    // OUT OF SCOPE CHECK — BEFORE formatting and before sending success:true
    // ❗ Level 1 users should NOT trigger out-of-scope check
    if (
      level !== "Level 1" &&
      dbRows.length === 0 &&
      branchCode &&
      (clusterName || districtName)
    ) {
      return res.status(403).json({
        success: false,
        message: "Branch is out of scope",
      });
    }

    // ============================================================================================
    // 7️⃣ COUNT RECORDS (fixed — uses SAME filters)
    // ============================================================================================
    const totalRecords =
      (await queryUTIDatabase(countQuery, countParams))[0]?.totalRecords || 0;

    // ============================================================================================
    // 8️⃣ FORMAT OUTPUT (React-safe)
    // ============================================================================================
    const meaningMap = {};
    mapping.forEach((m) => {
      if (m.meaning) meaningMap[m.displayLabel] = m.meaning;
    });

    const formattedRaw = dbRows.map((row) => {
      const obj = {};
      mapping.forEach((m) => {
        let val = row[m.colName];
        if (val === null || val === "" || val === undefined) val = "-";
        if (typeof val === "object") val = "-";
        obj[m.displayLabel] = val;
      });
      return obj;
    });

    const formatted = formatWithMeanings(formattedRaw, meaningMap);


    // ============================================================================================
    // 🔟 SEND RESPONSE
    // ============================================================================================
    return res.json({
      success: true,
      headers: mapping.map((m) => m.displayLabel),
      groupedHeaders: [],
      mapping,
      data: formatted,
      sortableColumns: sortableRows,
      totalRecords,
      totalPages: Math.ceil(totalRecords / pageSize),
      fixedCluster: effectiveCluster || null,
    });
  } catch (err) {
    console.error("❌ Deposits Error:", err);
    return res.status(500).json({
      success: false,
      message: "Server error fetching deposits",
      error: err.message,
    });
  }
});



//============================================================================================
//     DEPOSITS ACCOUNTS OPENED (Meaning-based Dynamic Sorting)
//============================================================================================

app.post("/get-deposits-accounts", async (req, res) => {
  console.log("🔹 /get-deposits-accounts called");

  try {
    let {
      branchCode,
      branchName,
      districtName,
      clusterName,
      userId,
      role,
      level,
      designation,
      sortBy,
      sortOrder = "ASC",
      fetchAll = false,
      page = 1,
      pageSize = 10,
    } = req.body;

    // ------------------- Normalize Inputs -------------------
    branchCode =
      branchCode !== undefined && branchCode !== null
        ? branchCode.toString().trim()
        : "";

    branchName = (branchName || "").trim();
    districtName = (districtName || "").trim();
    clusterName = (clusterName || "").trim();

    page = parseInt(page, 10);
    pageSize = parseInt(pageSize, 10);
    const offset = (page - 1) * pageSize;
    level = level || "Level 1";

    //===========================================================
    //   🔥 LEVEL-1 FIX — convert "2.00" → "2"
    //===========================================================
    if (level === "Level 1" && branchCode) {
      const numeric = parseFloat(branchCode);
      if (!Number.isNaN(numeric)) {
        branchCode = parseInt(numeric, 10).toString();   // "2.00" → "2"
      }
    }

    const SECTION = "deposits_accounts_opened";

    //========================================================================================
    // 1️⃣ LOAD MEANING → colX MAPPING
    //========================================================================================
    const mappingRows = await queryUTIDatabase(
      `
      SELECT col_number, display_label, meaning, detected_type
      FROM MIS.dbo.MIS_Column_Mapping
      WHERE section = ?
      ORDER BY col_number ASC
      `,
      [SECTION]
    );

    let mapping = mappingRows.map((m) => ({
      colNumber: Number(m.col_number),
      colName: `col${m.col_number}`,
      displayLabel: m.display_label,
      meaning: m.meaning,
      type: m.detected_type || "string",
    }));

    if (!mapping.length) {
      mapping = [];
      for (let i = 1; i <= 30; i++) {
        mapping.push({
          colNumber: i,
          colName: `col${i}`,
          displayLabel: `col${i}`,
          meaning:
            i === 1
              ? "branch_code"
              : i === 2
              ? "branch_name"
              : i === 3
              ? "district"
              : i === 4
              ? "cluster"
              : null,
        });
      }
    }

    const meaningToCol = {};
    mapping.forEach((m) => {
      if (m.meaning) meaningToCol[m.meaning] = m.colName;
    });

    const fallback = {
      branch_code: "col1",
      branch_name: "col2",
      district: "col3",
      cluster: "col4",
    };

    const getCol = (meaning) => meaningToCol[meaning] || fallback[meaning];

    const colNames = mapping.map((m) => `[${m.colName}]`);

    //========================================================================================
    // 2️⃣ LOAD SORTABLE COLUMNS
    //========================================================================================
    const sortableRows = await queryUTIDatabase(
      `
      SELECT meaning
      FROM MIS.dbo.MIS_Sortable_Columns
      WHERE section = ?
      ORDER BY meaning ASC
      `,
      [SECTION]
    );

    const allowedSortableMeaning = sortableRows
      .map((r) => r.meaning)
      .filter((m) => m);

    //========================================================================================
    // 3️⃣ BUILD BASE QUERY + FILTER HELPER
    //========================================================================================
    let query = `SELECT ${colNames.join(", ")} FROM MIS.dbo.[${SECTION}] WHERE 1=1`;
    let countQuery = `SELECT COUNT(*) AS totalRecords FROM MIS.dbo.[${SECTION}] WHERE 1=1`;

    const params = [];
    const countParams = [];

    // ---------------------------------------------------------
    // FILTER HELPER
    // ---------------------------------------------------------
    const addFilter = (meaning, value, exact = true) => {
      if (!value) return;
      const col = getCol(meaning);
      if (!col) return;

      // ============================================
      // 🔥 SPECIAL CASE: branch_code numeric compare
      // ============================================
      if (meaning === "branch_code" && exact) {
        query += ` AND TRY_CAST(LTRIM(RTRIM([${col}])) AS INT) = TRY_CAST(? AS INT)`;
        countQuery += ` AND TRY_CAST(LTRIM(RTRIM([${col}])) AS INT) = TRY_CAST(? AS INT)`;
        params.push(value);
        countParams.push(value);
        return;
      }

      if (exact) {
        query += ` AND LTRIM(RTRIM([${col}])) = ?`;
        countQuery += ` AND LTRIM(RTRIM([${col}])) = ?`;
        params.push(value.trim());
        countParams.push(value.trim());
      } else {
        query += ` AND LOWER(LTRIM(RTRIM([${col}]))) LIKE ?`;
        countQuery += ` AND LOWER(LTRIM(RTRIM([${col}]))) LIKE ?`;
        params.push(`%${value.trim().toLowerCase()}%`);
        countParams.push(`%${value.trim().toLowerCase()}%`);
      }
    };

    //========================================================================================
    // 4️⃣ LEVEL-BASED ACCESS CONTROL
    //========================================================================================
    let effectiveCluster = clusterName;

    if (level === "Level 2" && designation) {
      const m = designation.match(/Cluster\s*Head\s*[-:]\s*(.*)/i);
      if (m && m[1]) effectiveCluster = m[1].trim();
    }

    if (level === "Level 1" && branchCode) {
      addFilter("branch_code", branchCode);
    }

    if (level === "Level 2") {
      if (!effectiveCluster)
        return res.status(400).json({
          success: false,
          message: "Cluster missing for Level 2 user.",
        });

      addFilter("cluster", effectiveCluster, true);
    }

    //========================================================================================
    // 5️⃣ NORMAL FILTERS
    //========================================================================================
    if (branchCode && level !== "Level 1") addFilter("branch_code", branchCode);
    if (branchName) addFilter("branch_name", branchName, false);
    if (districtName) addFilter("district", districtName, false);
    if (clusterName && level !== "Level 2") addFilter("cluster", clusterName);

    //========================================================================================
    // 6️⃣ SORTING (NUMERIC-SAFE)
    //========================================================================================
    let sortCol = getCol("branch_code") || "col1";

    if (sortBy && allowedSortableMeaning.includes(sortBy)) {
      const c = getCol(sortBy);
      if (c) sortCol = c;
    }

    const sortDir = sortOrder.toUpperCase() === "DESC" ? "DESC" : "ASC";

    const orderBySQL = ` ORDER BY TRY_CAST(LTRIM(RTRIM([${sortCol}])) AS INT) ${sortDir}`;

    if (fetchAll) {
      query += orderBySQL;
    } else {
      query += `${orderBySQL} OFFSET ? ROWS FETCH NEXT ? ROWS ONLY`;
      params.push(offset, pageSize);
    }

    //========================================================================================
    // 7️⃣ EXECUTE
    //========================================================================================
    const dbRows = await queryUTIDatabase(query, params);
    const totalRecords =
      (await queryUTIDatabase(countQuery, countParams))[0]?.totalRecords || 0;

    //========================================================================================
    // 8️⃣ FORMAT RESULT
    //========================================================================================
    const data = dbRows.map((row) => {
      const obj = {};
      mapping.forEach((m) => {
        let val = row[m.colName];
        if (val === null || val === undefined || val === "") val = "-";
        if (val && typeof val === "object") val = "-";
        obj[m.displayLabel] = val;
      });
      return obj;
    });

    const headers = mapping.map((m) => m.displayLabel);
    const totalPages = fetchAll ? 1 : Math.ceil(totalRecords / pageSize);

    //========================================================================================
    // 🔟 RETURN RESPONSE
    //========================================================================================
    return res.json({
      success: true,
      headers,
      mapping,
      data,
      totalRecords,
      totalPages,
      sortableColumns: allowedSortableMeaning,
      fixedCluster: effectiveCluster || null,
    });
  } catch (err) {
    console.error("❌ DAO ERROR:", err);
    return res.status(500).json({
      success: false,
      message: "Server error fetching deposits accounts opened",
      error: err.message,
    });
  }
});


//============================================================================================
//                               ADVANCES (Meaning-based Sorting)
//============================================================================================

app.post("/get-advances", async (req, res) => {
  try {
    let {
      branchCode,
      branchName,
      clusterName,
      districtName,
      userId,
      role,
      level,
      designation,
      page = 1,
      pageSize = 10,
      sortBy,             // meaning
      sortOrder = "DESC",
    } = req.body;

    // ---------- Normalize ----------
    branchCode = (branchCode || "").trim();
    branchName = (branchName || "").trim();
    clusterName = (clusterName || "").trim();
    districtName = (districtName || "").trim();
    level = level || "Level 1";
	
	// ===============================
// FIX: Level 1 branchCode "2.00" → "2"
// ===============================
if (level === "Level 1" && branchCode) {
  const num = parseFloat(branchCode);
  if (!isNaN(num)) {
    branchCode = parseInt(num, 10).toString();   // "2.00" → "2"
  }
}


    page = parseInt(page, 10) || 1;
    pageSize = parseInt(pageSize, 10) || 10;

    const SECTION = "Advances";

    //========================================================================================
    // 1️⃣ LOAD MEANING → colX mapping
    //========================================================================================
    const mappingRows = await queryUTIDatabase(
      `
        SELECT col_number, display_label, meaning, detected_type
        FROM MIS.dbo.MIS_Column_Mapping
        WHERE section = ?
        ORDER BY col_number ASC
      `,
      [SECTION]
    );

    let mapping = mappingRows.map((m) => ({
      colNumber: Number(m.col_number),
      colName: `col${m.col_number}`,
      displayLabel: m.display_label,
      meaning: m.meaning || null,
      type: m.detected_type || "string",
    }));

    // Fallback (rare)
    if (!mapping.length) {
      mapping = [];
      for (let i = 1; i <= 30; i++) {
        mapping.push({
          colNumber: i,
          colName: `col${i}`,
          displayLabel: `col${i}`,
          meaning:
            i === 1 ? "branch_code" :
            i === 2 ? "branch_name" :
            i === 3 ? "district" :
            i === 4 ? "cluster" : null,
          type: "string",
        });
      }
    }

    const meaningToCol = {};
    mapping.forEach((m) => {
      if (m.meaning) meaningToCol[m.meaning] = m.colName;
    });

    const fallback = {
      branch_code: "col1",
      branch_name: "col2",
      district: "col3",
      cluster: "col4",
    };

    const getCol = (meaning) => meaningToCol[meaning] || fallback[meaning];

    const colNames = mapping.map((m) => `[${m.colName}]`);

    //========================================================================================
    // 2️⃣ LOAD SORTABLE MEANINGS (from MIS_Sortable_Columns)
//========================================================================================
    const sortableRows = await queryUTIDatabase(
      `
        SELECT meaning
        FROM MIS.dbo.MIS_Sortable_Columns
        WHERE section = ?
      `,
      [SECTION]
    );

    const allowedSortableMeaning = sortableRows
      .map((r) => r.meaning)
      .filter((m) => m);

    //========================================================================================
    // 3️⃣ BUILD QUERY + FILTER HELPER
    //========================================================================================
    let query = `SELECT ${colNames.join(", ")} FROM [dbo].[Advances] WHERE 1=1`;
    let params = [];

    const addFilter = (meaning, value, exact = true) => {
      if (!value) return;

      const col = getCol(meaning);
      if (!col) return;

      if (meaning === "branch_code") {
  query += ` AND TRY_CAST(LTRIM(RTRIM([${col}])) AS INT) = TRY_CAST(? AS INT)`;
  params.push(value);
  return;
}
 else {
        query += ` AND LOWER(LTRIM(RTRIM([${col}]))) LIKE ?`;
        params.push(`%${value.trim().toLowerCase()}%`);
      }
    };

    //========================================================================================
    // 4️⃣ LEVEL-BASED ACCESS CONTROL
    //========================================================================================
    let effectiveCluster = clusterName;

    if (level === "Level 2" && designation) {
      const m = designation.match(/Cluster\s*Head\s*[-:]\s*(.*)/i);
      if (m && m[1]) effectiveCluster = m[1].trim();
    }

    if (level === "Level 1" && branchCode) {
      addFilter("branch_code", branchCode, true);
    }

    if (level === "Level 2") {
      if (!effectiveCluster)
        return res.status(400).json({ success: false, message: "Cluster missing." });

      addFilter("cluster", effectiveCluster, true);
    }

    // NORMAL FILTERS
    if (branchCode && level !== "Level 1") addFilter("branch_code", branchCode, true);
    if (branchName) addFilter("branch_name", branchName, false);
    if (districtName) addFilter("district", districtName, false);
    if (clusterName && level !== "Level 2") addFilter("cluster", clusterName, false);

    //========================================================================================
    // 5️⃣ SORTING (MEANING-BASED)
    //========================================================================================
    let sortCol = getCol("branch_code") || "col1";

    if (sortBy && allowedSortableMeaning.includes(sortBy)) {
      const c = getCol(sortBy);
      if (c) sortCol = c;
    }

    const sortDir = sortOrder.toUpperCase() === "ASC" ? "ASC" : "DESC";

    query += ` ORDER BY TRY_CAST(LTRIM(RTRIM([${sortCol}])) AS DECIMAL(18,2)) ${sortDir}`;

    //========================================================================================
    // 6️⃣ PAGINATION
    //========================================================================================
    const offset = (page - 1) * pageSize;
    query += " OFFSET ? ROWS FETCH NEXT ? ROWS ONLY";
    params.push(offset, pageSize);

    const dbRows = await queryUTIDatabase(query, params);

    //========================================================================================
    // 7️⃣ FORMAT OUTPUT — (React-safe)
//========================================================================================
    const data = dbRows.map((row) => {
      const obj = {};
      mapping.forEach((m) => {
        const raw = row[m.colName];
        if (raw && typeof raw === "object") obj[m.displayLabel] = "-";
        else obj[m.displayLabel] = raw ?? "-";
      });
      return obj;
    });

    const headers = mapping.map((m) => m.displayLabel);

    //========================================================================================
    // 8️⃣ TOTAL COUNT
    //========================================================================================
    let countQuery = `SELECT COUNT(*) AS totalRecords FROM [dbo].[Advances] WHERE 1=1`;
    let countParams = [];

    const addCountFilter = (meaning, value, exact = true) => {
      if (!value) return;
      const col = getCol(meaning);
      if (!col) return;

      if (exact) {
        countQuery += ` AND LTRIM(RTRIM([${col}])) = ?`;
        countParams.push(value.trim());
      } else {
        countQuery += ` AND LOWER(LTRIM(RTRIM([${col}]))) LIKE ?`;
        countParams.push(`%${value.trim().toLowerCase()}%`);
      }
    };

    // replicate same filters
    if (level === "Level 1" && branchCode) addCountFilter("branch_code", branchCode);
    if (level === "Level 2" && effectiveCluster) addCountFilter("cluster", effectiveCluster);

    if (branchCode && level !== "Level 1") addCountFilter("branch_code", branchCode);
    if (branchName) addCountFilter("branch_name", branchName, false);
    if (districtName) addCountFilter("district", districtName, false);
    if (clusterName && level !== "Level 2")
      addCountFilter("cluster", clusterName, false);

    const totalRecords =
      (await queryUTIDatabase(countQuery, countParams))[0]?.totalRecords || 0;

    //========================================================================================
    // 9️⃣ GROUPED HEADERS
    //========================================================================================
    const skipMeanings = new Set(["branch_code", "branch_name", "district", "cluster"]);
    const groupMap = new Map();

    mapping.forEach((m) => {
      if (skipMeanings.has(m.meaning)) return;

      const parts = m.displayLabel.split(" ");
      let title, child;

      if (parts.length === 1) {
        title = m.displayLabel;
        child = m.displayLabel;
      } else {
        title = parts.slice(0, -1).join(" ");
        child = parts[parts.length - 1];
      }

      if (!groupMap.has(title)) groupMap.set(title, { title, cols: [] });

      groupMap.get(title).cols.push({
        key: m.displayLabel,
        label: child,
      });
    });

    const groupedHeaders = Array.from(groupMap.values());

    //========================================================================================
    // 🔟 RETURN RESPONSE
    //========================================================================================
    await logActivity(
      userId,
      role,
      "View Advances",
      `Filters: Branch ${branchCode}, Cluster ${effectiveCluster}`
    );

    // Build meaningMap: displayLabel -> meaning (for frontend graphs)
    const meaningMap = {};
    mapping.forEach((m) => {
      if (m.meaning) {
        meaningMap[m.displayLabel] = m.meaning;
      }
    });

    //========================================================================================
    // 🔟 FINAL RESPONSE
    //========================================================================================
           return res.json({
      success: true,
      headers,
      groupedHeaders,
      data,
      totalRecords,
      totalPages: Math.ceil(totalRecords / pageSize),
      sortableColumns: allowedSortableMeaning,
      fixedCluster: effectiveCluster || null,
      meaningMap,  // 👈 added
    });


  } catch (err) {
    console.error("❌ ADVANCES ERROR:", err);
    return res.status(500).json({
      success: false,
      message: "Server error fetching advances",
      error: err.message,
    });
  }
});


//============================================================================================
//                             NPA REPORT (Meaning-based Sorting)
//============================================================================================

app.post("/get-npa", async (req, res) => {
  try {
    let {
      branchCode,
      branchName,
      clusterName,
      districtName,
      userId,
      role,
      level,
      designation,
      page = 1,
      pageSize = 10,
      sortBy,        // 👉 MEANING (not label)
      sortOrder = "DESC",
    } = req.body;

    // Normalize
    branchCode = (branchCode || "").trim();
    branchName = (branchName || "").trim();
    clusterName = (clusterName || "").trim();
    districtName = (districtName || "").trim();
    level = level || "Level 1";
	
	// ===============================
// FIX: Level 1 branchCode "2.00" → "2"
// ===============================
if (level === "Level 1" && branchCode) {
  const num = parseFloat(branchCode);
  if (!isNaN(num)) {
    branchCode = parseInt(num, 10).toString();   // "2.00" → "2"
  }
}


    page = parseInt(page, 10) || 1;
    pageSize = parseInt(pageSize, 10) || 10;

    const SECTION = "NPA";

    //========================================================================================
    // 1️⃣ LOAD MAPPING — meaning → colX (from upload MIS)
//========================================================================================
    const mappingRows = await queryUTIDatabase(
      `
      SELECT col_number, display_label, meaning, detected_type
      FROM MIS.dbo.MIS_Column_Mapping
      WHERE section = ?
      ORDER BY col_number ASC
      `,
      [SECTION]
    );

    let mapping = (mappingRows || []).map((m) => ({
      colNumber: Number(m.col_number),
      colName: `col${m.col_number}`,
      displayLabel: m.display_label,
      meaning: m.meaning || null,
      detectedType: m.detected_type || "string",
    }));

    // Fallback (rare)
    if (!mapping.length) {
      mapping = [];
      for (let i = 1; i <= 30; i++) {
        mapping.push({
          colNumber: i,
          colName: `col${i}`,
          displayLabel: `col${i}`,
          meaning:
            i === 1
              ? "branch_code"
              : i === 2
              ? "branch_name"
              : i === 3
              ? "district"
              : i === 4
              ? "cluster"
              : null,
          detectedType: "string",
        });
      }
    }

    const meaningToCol = {};
    mapping.forEach((m) => {
      if (m.meaning) meaningToCol[m.meaning] = m.colName;
    });

    const fallback = {
      branch_code: "col1",
      branch_name: "col2",
      district: "col3",
      cluster: "col4",
    };

    const getCol = (meaning) =>
      meaningToCol[meaning] || fallback[meaning] || null;

    const colNames = mapping.map((m) => `[${m.colName}]`);

    //========================================================================================
    // 2️⃣ LOAD SORTABLE MEANINGS — from MIS_Sortable_Columns
//========================================================================================
    const sortableRows = await queryUTIDatabase(
      `
      SELECT meaning
      FROM MIS.dbo.MIS_Sortable_Columns
      WHERE section = ?
      `,
      [SECTION]
    );

    const allowedSortableMeaning = (sortableRows || [])
      .map((r) => r.meaning)
      .filter((m) => m);

    //========================================================================================
    // 3️⃣ BUILD BASE QUERY
//========================================================================================
    let query = `SELECT ${colNames.join(", ")} FROM [dbo].[${SECTION}] WHERE 1=1`;
    let countQuery = `SELECT COUNT(*) AS totalRecords FROM [dbo].[${SECTION}] WHERE 1=1`;

    let params = [];
    let countParams = [];

    const addFilter = (meaning, value, exact = true) => {
  if (!value) return;

  const col = getCol(meaning);
  if (!col) return;

  // ================================
  // ⭐ FIX FOR BRANCH CODE (ALL LEVELS)
  // numeric-safe comparison
  // ================================
  if (meaning === "branch_code") {
    query += ` AND TRY_CAST(LTRIM(RTRIM([${col}])) AS INT) = TRY_CAST(? AS INT)`;
    countQuery += ` AND TRY_CAST(LTRIM(RTRIM([${col}])) AS INT) = TRY_CAST(? AS INT)`;
    params.push(value);
    countParams.push(value);
    return;
  }

  // ================================
  // Normal filters
  // ================================
  if (exact) {
    query += ` AND LTRIM(RTRIM([${col}])) = ?`;
    countQuery += ` AND LTRIM(RTRIM([${col}])) = ?`;
    params.push(value.trim());
    countParams.push(value.trim());
  } else {
    query += ` AND LOWER(LTRIM(RTRIM([${col}]))) LIKE ?`;
    countQuery += ` AND LOWER(LTRIM(RTRIM([${col}]))) LIKE ?`;
    params.push(`%${value.trim().toLowerCase()}%`);
    countParams.push(`%${value.trim().toLowerCase()}%`);
  }
};
;

    //========================================================================================
    // 4️⃣ LEVEL-BASED ACCESS CONTROL (RLS)
//========================================================================================
    let effectiveCluster = clusterName;

    if (level === "Level 2" && designation) {
      const match = designation.match(/Cluster\s*Head\s*[-:]\s*(.*)/i);
      if (match && match[1]) effectiveCluster = match[1].trim();
    }

    if (level === "Level 1" && branchCode) {
      addFilter("branch_code", branchCode, true);
    }

    if (level === "Level 2" && effectiveCluster) {
      addFilter("cluster", effectiveCluster, true);
    }

    // Other filters
    if (branchCode && level !== "Level 1") addFilter("branch_code", branchCode);
    if (branchName && level !== "Level 1") addFilter("branch_name", branchName, false);
    if (districtName) addFilter("district", districtName, false);
    if (clusterName && level !== "Level 2") addFilter("cluster", clusterName, false);

    //========================================================================================
    // 5️⃣ MEANING-BASED SORTING
//========================================================================================
    let sortCol = getCol("branch_code") || "col1";

    if (sortBy && allowedSortableMeaning.includes(sortBy)) {
      const col = getCol(sortBy);
      if (col) sortCol = col;
    }

    const sortDir =
      sortOrder && sortOrder.toUpperCase() === "ASC" ? "ASC" : "DESC";

    query += ` ORDER BY TRY_CAST(LTRIM(RTRIM([${sortCol}])) AS DECIMAL(18,2)) ${sortDir}`;

    //========================================================================================
    // Pagination
//========================================================================================
    const offset = (page - 1) * pageSize;
    query += " OFFSET ? ROWS FETCH NEXT ? ROWS ONLY";
    params.push(offset, pageSize);

    const dbRows = await queryUTIDatabase(query, params);
    const totalRecords =
      (await queryUTIDatabase(countQuery, countParams))[0]?.totalRecords || 0;

    //========================================================================================
    // 6️⃣ MAP colX → displayLabel (safe for React)
//========================================================================================
    const data = dbRows.map((row) => {
      const obj = {};
      mapping.forEach((m) => {
        const raw = row[m.colName];
        if (raw && typeof raw === "object") {
          obj[m.displayLabel] = "-"; // prevent React crash
        } else {
          obj[m.displayLabel] =
            raw === null || raw === undefined || raw === "" ? "-" : raw;
        }
      });
      return obj;
    });

    const headers = mapping.map((m) => m.displayLabel);

    //========================================================================================
    // 7️⃣ GROUP HEADERS
//========================================================================================
    const ignoreMeanings = new Set([
      "branch_code",
      "branch_name",
      "district",
      "cluster",
    ]);

    const groupMap = new Map();

    mapping.forEach((m) => {
      if (!m.displayLabel) return;
      if (ignoreMeanings.has(m.meaning)) return;

      const parts = m.displayLabel.split(/\s+/);

      let title, child;
      if (parts.length === 1) {
        title = m.displayLabel;
        child = m.displayLabel;
      } else {
        title = parts.slice(0, -1).join(" "); // e.g. STAMPED NPA As on Date
        child = parts[parts.length - 1];      // e.g. Balance / A/Cs
      }

      if (!groupMap.has(title)) groupMap.set(title, { title, cols: [] });

      groupMap.get(title).cols.push({
        key: m.displayLabel,
        label: child,
      });
    });

    const groupedHeaders = Array.from(groupMap.values());

    //========================================================================================
    // 8️⃣ RESPONSE
//========================================================================================
    await logActivity(
      userId,
      role,
      "View NPA",
      `Filters → Branch: ${branchCode}, Cluster: ${effectiveCluster}`
    );

   // Build meaningMap: displayLabel -> meaning (for frontend graphs)
    const meaningMap = {};
    mapping.forEach((m) => {
      if (m.meaning) {
        meaningMap[m.displayLabel] = m.meaning;
      }
    });

    //========================================================================================
    // 🔟 FINAL RESPONSE
    //========================================================================================
           return res.json({
      success: true,
      headers,
      groupedHeaders,
      data,
      totalRecords,
      totalPages: Math.ceil(totalRecords / pageSize),
      sortableColumns: allowedSortableMeaning,
      fixedCluster: effectiveCluster || null,
      meaningMap,  // 👈 added
    });

  } catch (err) {
    console.error("❌ NPA ERROR:", err);
    return res.status(500).json({
      success: false,
      message: "Server error fetching NPA",
      error: err.message,
    });
  }
});

//============================================================================================
//                         LOANS SANCTIONED (Meaning-based Sorting)
//============================================================================================

app.post("/get-loans-sanctioned", async (req, res) => {
  try {
    let {
      branchCode,
      branchName,
      clusterName,
      districtName,
      userId,
      role,
      level,
      designation,
      page = 1,
      pageSize = 10,
      sortBy,        // 👉 THIS IS MEANING (over_all_fy_accounts, over_all_fy_amount, etc.)
      sortOrder = "DESC",
    } = req.body;

    console.log("🔹 /get-loans-sanctioned (meaning-based) called");

    // ------------------------------------------------------
    // Normalize values
    // ------------------------------------------------------
    branchCode = (branchCode || "").trim();
    branchName = (branchName || "").trim();
    clusterName = (clusterName || "").trim();
    districtName = (districtName || "").trim();
    level = level || "Level 1";
	
	// ===============================
// FIX: Level 1 branchCode "2.00" → "2"
// ===============================
if (level === "Level 1" && branchCode) {
  const num = parseFloat(branchCode);
  if (!isNaN(num)) {
    branchCode = parseInt(num, 10).toString();
  }
}


    page = parseInt(page, 10) || 1;
    pageSize = parseInt(pageSize, 10) || 10;

    const SECTION = "loanssanctioned";

    //========================================================================================
    // 1️⃣ LOAD COLUMN MAPPING (MIS_Column_Mapping → meaning → colX)
    //========================================================================================
    const mappingRows = await queryUTIDatabase(
      `
      SELECT col_number, display_label, meaning, detected_type
      FROM MIS.dbo.MIS_Column_Mapping
      WHERE section = ?
      ORDER BY col_number ASC
      `,
      [SECTION]
    );

    let mapping = (mappingRows || []).map((m) => ({
      colNumber: Number(m.col_number),
      colName: `col${m.col_number}`,
      displayLabel: m.display_label,
      meaning: m.meaning || null,
      detectedType: (m.detected_type || "").toLowerCase(),
    }));

    // Fallback mapping if no MIS_Column_Mapping yet (very rare safety net)
    if (!mapping.length) {
      mapping = [];
      for (let i = 1; i <= 30; i++) {
        mapping.push({
          colNumber: i,
          colName: `col${i}`,
          displayLabel: `col${i}`,
          meaning:
            i === 1
              ? "branch_code"
              : i === 2
              ? "branch_name"
              : i === 3
              ? "district"
              : i === 4
              ? "cluster"
              : null,
          detectedType: "string",
        });
      }
    }

    // Helper maps
    const colNames = mapping.map((m) => `[${m.colName}]`);
    const meaningToCol = {};
    const labelToCol = {};

    mapping.forEach((m) => {
      if (m.meaning) {
        meaningToCol[m.meaning] = m.colName;
      }
      if (m.displayLabel) {
        labelToCol[m.displayLabel.trim().toLowerCase()] = m.colName;
      }
    });

    // Fallback for base columns
    const fallback = {
      branch_code: "col1",
      branch_name: "col2",
      district: "col3",
      cluster: "col4",
    };

    const getCol = (meaning) =>
      meaningToCol[meaning] || fallback[meaning] || null;

    //========================================================================================
    // 2️⃣ LOAD SORTABLE COLUMNS (MEANING LIST) FROM MIS_Sortable_Columns
    //========================================================================================
    const sortableRows = await queryUTIDatabase(
      `
      SELECT display_label, meaning
      FROM MIS.dbo.MIS_Sortable_Columns
      WHERE section = ?
      ORDER BY display_label ASC
      `,
      [SECTION]
    );

    const allowedSortableMeaning = (sortableRows || [])
      .filter((r) => r.meaning)
      .map((r) => r.meaning);

    //========================================================================================
    // 3️⃣ BUILD BASE QUERY
    //========================================================================================
    let query = `SELECT ${colNames.join(", ")} FROM [dbo].[${SECTION}] WHERE 1=1`;
    let countQuery = `SELECT COUNT(*) AS totalRecords FROM [dbo].[${SECTION}] WHERE 1=1`;

    let params = [];
    let countParams = [];

    const addFilter = (meaning, value, exact = true) => {
  if (!value) return;
  const col = getCol(meaning);
  if (!col) return;

  // ================================
  // ⭐ FIX FOR BRANCH CODE (numeric)
  // ================================
  if (meaning === "branch_code") {
    query += ` AND TRY_CAST(LTRIM(RTRIM([${col}])) AS INT) = TRY_CAST(? AS INT)`;
    countQuery += ` AND TRY_CAST(LTRIM(RTRIM([${col}])) AS INT) = TRY_CAST(? AS INT)`;
    params.push(value);
    countParams.push(value);
    return;
  }

  // ================================
  // Normal String Filters
  // ================================
  if (exact) {
    query += ` AND LTRIM(RTRIM([${col}])) = ?`;
    countQuery += ` AND LTRIM(RTRIM([${col}])) = ?`;
    params.push(value.trim());
    countParams.push(value.trim());
  } else {
    query += ` AND LOWER(LTRIM(RTRIM([${col}]))) LIKE ?`;
    countQuery += ` AND LOWER(LTRIM(RTRIM([${col}]))) LIKE ?`;
    params.push(`%${value.trim().toLowerCase()}%`);
    countParams.push(`%${value.trim().toLowerCase()}%`);
  }
};


    //========================================================================================
    // 4️⃣ LEVEL-BASED ACCESS CONTROL (RLS)
    //========================================================================================
    let effectiveCluster = null;

    if (level === "Level 2") {
      // Extract cluster from designation: "Cluster Head-XYZ" or "Cluster Head: XYZ"
      if (designation?.match(/Cluster\s*Head\s*[-:]\s*/i)) {
        effectiveCluster = designation.replace(
          /Cluster\s*Head\s*[-:]\s*/i,
          ""
        ).trim();
      } else if (clusterName) {
        effectiveCluster = clusterName.trim();
      }

      if (!effectiveCluster) {
        return res.status(400).json({
          success: false,
          message: "Cluster information missing for Level 2 user.",
        });
      }

      const clusterCol = getCol("cluster") || "col4";
      const branchCol = getCol("branch_code") || "col1";
      const nameCol = getCol("branch_name") || "col2";

      // ✅ Validate branchCode belongs to this cluster
      if (branchCode) {
        const chk = await queryUTIDatabase(
          `
          SELECT TOP 1 LTRIM(RTRIM([${clusterCol}])) AS Cluster
          FROM [dbo].[${SECTION}]
          WHERE LTRIM(RTRIM([${branchCol}])) = ?
        `,
          [branchCode]
        );

        if (
          !chk.length ||
          chk[0].Cluster.trim().toLowerCase() !== effectiveCluster.toLowerCase()
        ) {
          return res.status(403).json({
            success: false,
            message: `❌ Branch ${branchCode} does not belong to your cluster (${effectiveCluster}).`,
          });
        }
      }

      // ✅ Validate branchName belongs to this cluster
      if (branchName) {
        const chk = await queryUTIDatabase(
          `
          SELECT TOP 1 LTRIM(RTRIM([${clusterCol}])) AS Cluster
          FROM [dbo].[${SECTION}]
          WHERE LOWER(LTRIM(RTRIM([${nameCol}]))) LIKE ?
        `,
          [`%${branchName.toLowerCase()}%`]
        );

        if (
          !chk.length ||
          chk[0].Cluster.trim().toLowerCase() !== effectiveCluster.toLowerCase()
        ) {
          return res.status(403).json({
            success: false,
            message: `❌ Branch ${branchName} does not belong to your cluster (${effectiveCluster}).`,
          });
        }
      }

      // Lock to this cluster
      addFilter("cluster", effectiveCluster, true);
    }

    // General filters
    if (branchCode) addFilter("branch_code", branchCode);
    if (branchName) addFilter("branch_name", branchName, false);
    if (districtName) addFilter("district", districtName, false);
    if (clusterName && level !== "Level 2")
      addFilter("cluster", clusterName, false);

    //========================================================================================
    // 5️⃣ MEANING-BASED SORTING
    //========================================================================================
    let orderCol = getCol("branch_code") || "col1";

    if (sortBy && allowedSortableMeaning.includes(sortBy)) {
      const col = getCol(sortBy);
      if (col) orderCol = col;
    }

    const orderDir =
      sortOrder && sortOrder.toUpperCase() === "ASC" ? "ASC" : "DESC";

    // Always numeric sort for metrics (and branch code also works)
    query += ` ORDER BY TRY_CAST(LTRIM(RTRIM([${orderCol}])) AS DECIMAL(18,2)) ${orderDir}`;

    //========================================================================================
    // 6️⃣ PAGINATION
    //========================================================================================
    const offset = (page - 1) * pageSize;
    query += " OFFSET ? ROWS FETCH NEXT ? ROWS ONLY";
    params.push(offset, pageSize);
    //========================================================================================
    // 7️⃣ EXECUTE MAIN & COUNT QUERIES
    //========================================================================================
    const dbRows = await queryUTIDatabase(query, params);
    const countRows = await queryUTIDatabase(countQuery, countParams);
    const totalRecords = Number(countRows?.[0]?.totalRecords || 0);
    const totalPages = Math.ceil(totalRecords / pageSize);

    //========================================================================================
    // 8️⃣ FORMAT DATA (colX → displayLabel) + SAFE VALUES
    //========================================================================================
    const skipWords = ["total"]; // used later for grouping, not here

    const data = dbRows.map((row) => {
      const obj = {};
      mapping.forEach((m) => {
        const raw = row[m.colName];

        if (raw && typeof raw === "object") {
          // 🔒 Prevent React Native crash from unexpected objects
          obj[m.displayLabel] = "-";
        } else if (raw === null || raw === "") {
          obj[m.displayLabel] = "-";
        } else {
          obj[m.displayLabel] = raw;
        }
      });
      return obj;
    });

    const headers = mapping.map((m) => m.displayLabel);

    //========================================================================================
    // 9️⃣ GROUPED HEADERS (For Horizontal Grouped Table)
//========================================================================================
    const baseMeanings = new Set([
      "branch_code",
      "branch_name",
      "district",
      "cluster",
    ]);

    const groupMap = new Map();

    mapping.forEach((m) => {
      if (!m.displayLabel) return;

      // ❌ Don’t group base columns
      if (baseMeanings.has(m.meaning)) return;

      // ❌ Don’t group columns that look like TOTAL
      if (skipWords.some((w) => m.displayLabel.toLowerCase().includes(w))) {
        return;
      }

      const label = m.displayLabel.trim();
      const parts = label.split(/\s+/);

      let title, child;

      // Rule: group title = all words except last, child = last word
      if (parts.length === 1) {
        title = label;
        child = label;
      } else {
        title = parts.slice(0, -1).join(" "); // e.g. "OVER ALL FY from April-2025"
        child = parts[parts.length - 1];      // e.g. "A/Cs" / "Amount"
      }

      if (!groupMap.has(title)) {
        groupMap.set(title, { title, cols: [] });
      }

      groupMap.get(title).cols.push({
        key: label,   // full header used in data object
        label: child, // short label for grouped sub-header
      });
    });

    const groupedHeaders = Array.from(groupMap.values());
    // Build meaningMap: displayLabel -> meaning (for frontend graphs)
    const meaningMap = {};
    mapping.forEach((m) => {
      if (m.meaning) {
        meaningMap[m.displayLabel] = m.meaning;
      }
    });

    //========================================================================================
    // 🔟 FINAL RESPONSE
    //========================================================================================
        return res.json({
      success: true,
      headers,          // full headers as shown in table
      groupedHeaders,   // for grouped horizontal view (Type / A/Cs / Amount)
      data,             // row objects with displayLabel keys
      totalRecords,
      totalPages,
      sortableColumns: (sortableRows || [])
        .filter((r) => r.meaning)
        .map((r) => ({
          label: r.display_label,
          meaning: r.meaning,
        })),
      fixedCluster: effectiveCluster || null,
      meaningMap,       // 👈 added
    });

  } catch (err) {
    console.error("❌ Error (meaning-based loans sanctioned):", err);
    return res.status(500).json({
      success: false,
      message: "Server error fetching loans sanctioned",
      error: err.message,
    });
  }
});


//============================================================================================
//                     DEPOSITS CLUSTER SUMMARY (Dynamic Mapping + Intelligent Grouping)
//============================================================================================

app.post("/get-cluster-deposits-summary", async (req, res) => {
  try {
    const { userId, role, level, designation } = req.body;
    const SECTION = "DCS"; // Deposits cluster section (same as ACS/NPA mapping format)

    // 1️⃣ Fetch mapping
    const mappingRows = await queryUTIDatabase(
      `
        SELECT col_number, display_label, meaning 
        FROM MIS.dbo.MIS_Column_Mapping
        WHERE section = ?
        ORDER BY col_number ASC
      `,
      [SECTION]
    );

    if (!mappingRows.length) {
      return res.status(400).json({
        success: false,
        message: "⚠️ Deposits summary mapping missing. Upload DCS MIS first.",
      });
    }

    // Build mapping array
    const mapping = mappingRows.map((m) => ({
      colNumber: Number(m.col_number),
      colName: `col${m.col_number}`,
      displayLabel: m.display_label,
      meaning: m.meaning || null,
    }));

    // Build meaning lookup
    const byMeaning = {};
    mapping.forEach((m) => {
      if (m.meaning) byMeaning[m.meaning] = m.colName;
    });

    // Fallbacks if meaning not found
    const fallback = {
      cluster: "col1",
      branch_code: "col1",
      branch_name: "col2",
      district: "col3",
      as_on_date: "col5",
    };

    const getColByMeaning = (meaning) =>
      byMeaning[meaning] || fallback[meaning] || null;

    const clusterCol = getColByMeaning("cluster") || "col1";

    // Build SELECT list
    const colList = mapping.map((m) => `[${m.colName}]`).join(", ");

    // 2️⃣ Build SQL with Level-2 filtering
    let query = `SELECT ${colList} FROM [dbo].[DCS] WHERE 1=1`;
    const params = [];

    let effectiveLevel = level || "Level 1";
    let effectiveCluster = null;

    // Level 2 → Determine Cluster
    if (effectiveLevel === "Level 2") {
      if (designation && designation.toLowerCase().includes("cluster head")) {
        effectiveCluster = designation.replace(/Cluster Head\s*-\s*/i, "").trim();
      }

      if (effectiveCluster) {
        query += ` AND LOWER(LTRIM(RTRIM([${clusterCol}]))) = ?`;
        params.push(effectiveCluster.toLowerCase());
      } else {
        console.warn("⚠ No cluster found for Level 2 user — blocking data.");
        query += " AND 1=0"; // Block unauthorized access
      }
    }

    // 3️⃣ Custom ordering
    query += `
      ORDER BY CASE 
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'Krishna' THEN 1
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'Guntur' THEN 2
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'West Godavari' THEN 3
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'Visakhapatnam' THEN 4
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'Corporate Office' THEN 5
        WHEN UPPER(LTRIM(RTRIM([${clusterCol}]))) = 'TOTAL' THEN 6
        ELSE 7
      END
    `;

    // 4️⃣ Run Query
    const rows = await queryUTIDatabase(query, params);

    if (!rows.length) {
      return res.json({
        success: true,
        headers: [],
        groupedHeaders: [],
        data: [],
        totalRecords: 0,
      });
    }

    // 5️⃣ Convert SQL rows → human-readable rows
    const data = rows.map((r) => {
      const obj = {};
      mapping.forEach((m) => {
        const v = r[m.colName];
        obj[m.displayLabel] =
          v === null || v === undefined || v === "" ? "-" : v;
      });
      return obj;
    });

    const totalRecords = rows.length;

    // 6️⃣ Static headers
    const staticMeaning = new Set([
      "branch_code",
      "branch_name",
      "district",
      "cluster",
      "as_on_date",
    ]);

    const staticHeaders = mapping
      .filter((m) => m.meaning && staticMeaning.has(m.meaning))
      .map((m) => m.displayLabel);

    // 7️⃣ Intelligent grouping
    const groupBucket = {};

    mapping.forEach((m) => {
      if (m.meaning && staticMeaning.has(m.meaning)) return;

      const label = m.displayLabel.trim();
      const parts = label.split(/\s+/);

      let title, child;

      if (parts.length === 1) {
        title = label;
        child = label;
      } else {
        title = parts.slice(0, -1).join(" ");
        child = parts[parts.length - 1];
      }

      if (!groupBucket[title]) groupBucket[title] = [];
      groupBucket[title].push({ key: m.displayLabel, label: child });
    });

    // 8️⃣ Final grouped format
    const groupedHeaders = [];

    for (const title in groupBucket) {
      const cols = groupBucket[title];
      if (cols.length >= 2) groupedHeaders.push({ title, cols });
      else groupedHeaders.push({ title, cols: [cols[0]], single: true });
    }

    // 9️⃣ Return response
    return res.json({
  success: true,
  headers: staticHeaders.length ? staticHeaders : ["Cluster"],
  groupedHeaders,
  data,
  totalRecords,
});

  } catch (err) {
    console.error("❌ Error in /get-cluster-deposits-summary:", err);
    return res.status(500).json({
      success: false,
      error: "Server error while fetching Deposits Summary",
    });
  }
});



//============================================================================================
//                        DEPOSITS ACCOUNTS OPENED (DAOCS) CLUSTER SUMMARY
//============================================================================================

app.post("/get-daocs", async (req, res) => {
  try {
    const { userId, role, level, clusterName, branchCode, designation } = req.body;

    const SECTION = "DAOCS"; // mapping section

    // 1️⃣ LOAD MAPPING FOR DAOCS
    const mappingRows = await queryUTIDatabase(
      `
      SELECT col_number, display_label, meaning
      FROM MIS.dbo.MIS_Column_Mapping
      WHERE section = ?
      ORDER BY col_number ASC
      `,
      [SECTION]
    );

    if (!mappingRows.length) {
      return res.status(400).json({
        success: false,
        message:
          "⚠️ DAOCS summary mapping not found. Please upload DAOCS MIS first.",
      });
    }

    const mapping = mappingRows.map((m) => ({
      colNumber: Number(m.col_number),
      colName: `col${m.col_number}`,
      displayLabel: m.display_label,
      meaning: m.meaning || null,
    }));

    // Meaning lookup table
    const byMeaning = {};
    mapping.forEach((m) => {
      if (m.meaning) byMeaning[m.meaning] = m.colName;
    });

    const fallback = {
      cluster: "col1",
      branch_code: "col1",
      branch_name: "col2",
      district: "col3",
    };

    const getColByMeaning = (meaning) =>
      byMeaning[meaning] || fallback[meaning] || null;

    const clusterCol = getColByMeaning("cluster");

    // 2️⃣ BUILD BASE QUERY
    const colList = mapping.map((m) => `[${m.colName}]`).join(", ");
    let query = `SELECT ${colList} FROM [dbo].[DAOCS] WHERE 1=1`;
    const params = [];

    // ------ LEVEL-2 RESTRICTION -------
    if (level === "Level 2") {
      let cluster = null;

      if (clusterName) cluster = clusterName.trim().toLowerCase();
      else if (designation?.toLowerCase().includes("cluster head")) {
        cluster = designation.replace(/cluster head\s*[-:]?/i, "").trim().toLowerCase();
      }

      if (cluster) {
        query += ` AND LOWER(LTRIM(RTRIM([${clusterCol}]))) = ?`;
        params.push(cluster);
      }
    }

    // ------ LEVEL-1 RESTRICTION -------
    const branchCodeCol = getColByMeaning("branch_code");

    if (level === "Level 1" && branchCode) {
      query += ` AND LTRIM(RTRIM([${branchCodeCol}])) = ?`;
      params.push(branchCode.trim());
    }

    // 3️⃣ ORDERING
    query += `
      ORDER BY CASE 
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'Krishna' THEN 1
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'Guntur' THEN 2
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'West Godavari' THEN 3
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'Visakhapatnam' THEN 4
        WHEN UPPER(LTRIM(RTRIM([${clusterCol}]))) = 'TOTAL' THEN 5
        ELSE 6
      END
    `;

    const rows = await queryUTIDatabase(query, params);

    // 4️⃣ MAP colX → displayLabel
    const data = rows.map((r) => {
      const obj = {};
      mapping.forEach((m) => {
        const raw = r[m.colName];
        obj[m.displayLabel] = raw == null || raw === "" ? "-" : raw;
      });
      return obj;
    });

    if (!data.length) {
      return res.json({
        success: true,
        headers: [],
        groupedHeaders: [],
        data: [],
        totalRecords: 0,
      });
    }

    // 5️⃣ STATIC HEADERS (cluster, branch info)
    const branchSet = new Set([
      "cluster",
      "branch_code",
      "branch_name",
      "district",
    ]);

    const staticHeaders = mapping
      .filter((m) => m.meaning && branchSet.has(m.meaning))
      .map((m) => m.displayLabel);
// 6️⃣ GROUPED HEADERS (special-case for DAOCS to ensure TDR / RD / Savings grouping)
let groupedHeaders = [];
if (SECTION === "DAOCS") {
  // Build groups using meanings → column names (colX) and display labels.
  // byMeaning maps meaning -> colName (eg. 'tdr_accounts' -> 'col5')
  const makeColObj = (meaning, fallbackLabel) => {
    const colName = getColByMeaning(meaning);
    if (!colName) return null;
    // find mapping row to get displayLabel (safer)
    const mapRow = mapping.find((m) => m.colName === colName);
    const label = (mapRow && mapRow.displayLabel) || fallbackLabel || meaning;
    return { key: label, label }; // frontend expects key = displayLabel
  };

  const groups = [
    {
      title: "TDR",
      cols: [
        makeColObj("tdr_accounts", "TDR A/Cs"),
        makeColObj("tdr_amount", "TDR Amount"),
      ].filter(Boolean),
    },
    {
      title: "RD",
      cols: [
        makeColObj("rd_accounts", "RD A/Cs"),
        makeColObj("rd_amount", "RD Amount"),
      ].filter(Boolean),
    },
    {
      title: "Savings",
      cols: [
        makeColObj("savings_accounts", "Savings A/Cs"),
        makeColObj("savings_amount", "Savings Amount"),
      ].filter(Boolean),
    },
    {
      title: "Others",
      cols: [
        makeColObj("vouchers", "Vouchers"),
      ].filter(Boolean),
    },
  ];

  // remove empty groups (if a column wasn't present)
  groupedHeaders = groups.filter((g) => g.cols && g.cols.length);
} else {
  // existing generic grouping (split last word)
  const groupMap = new Map();
  mapping.forEach((m) => {
    if (m.meaning && branchSet.has(m.meaning)) return;

    const label = m.displayLabel.trim();
    const parts = label.split(/\s+/);

    let title, child;
    if (parts.length === 1) {
      title = label;
      child = label;
    } else {
      title = parts.slice(0, -1).join(" ");
      child = parts[parts.length - 1];
    }

    if (!groupMap.has(title)) {
      groupMap.set(title, { title, cols: [] });
    }

    groupMap.get(title).cols.push({
      key: m.displayLabel,
      label: child,
    });
  });

  groupedHeaders = Array.from(groupMap.values());
}
    // 7️⃣ LOG ACTIVITY
    await logActivity(
      userId,
      role,
      "View Deposits Accounts Opened Summary",
      `Fetched ${data.length} rows | Level:${level}`
    );

    return res.json({
      success: true,
      headers: staticHeaders,
      groupedHeaders,
      data,
      totalRecords: data.length,
    });
  } catch (err) {
    console.error("❌ Error in /get-daocs:", err);
    return res.status(500).json({
      success: false,
      message: "Server error while fetching DAOCS summary",
    });
  }
});

//============================================================================================
//                     ADVANCES SUMMARY (Dynamic Mapping + Intelligent Grouping)
//============================================================================================

app.post("/get-advances-summary", async (req, res) => {
  try {
    const { userId, role, level, clusterName } = req.body;
    const SECTION = "ACS"; // Advances Cluster Summary section

    // 1️⃣ Fetch mapping
    const mappingRows = await queryUTIDatabase(
      `
        SELECT col_number, display_label, meaning 
        FROM MIS.dbo.MIS_Column_Mapping
        WHERE section = ?
        ORDER BY col_number ASC
      `,
      [SECTION]
    );

    if (!mappingRows.length) {
      return res.status(400).json({
        success: false,
        message: "⚠️ Advances summary mapping missing. Upload ACS MIS first.",
      });
    }

    const mapping = mappingRows.map((m) => ({
      colNumber: Number(m.col_number),
      colName: `col${m.col_number}`,
      displayLabel: m.display_label,
      meaning: m.meaning || null,
    }));

    // Meaning → colX lookup table
    const byMeaning = {};
    mapping.forEach((m) => {
      if (m.meaning) byMeaning[m.meaning] = m.colName;
    });

    // Fallbacks (MEANING MUST NOT FALLBACK)
    const fallback = {
      cluster: "col1",
      branch_code: "col1",
      branch_name: "col2",
      district: "col3",
    };

    const getColByMeaning = (meaning) =>
      byMeaning[meaning] || fallback[meaning] || null;

    const clusterCol = getColByMeaning("cluster") || "col1";

    const colList = mapping.map((m) => `[${m.colName}]`).join(", ");

    // 2️⃣ Build SQL with Level-2 filtering
    let query = `SELECT ${colList} FROM [dbo].[ACS] WHERE 1=1`;
    const params = [];

    if (level === "Level 2" && clusterName?.trim()) {
      query += ` AND LOWER(LTRIM(RTRIM([${clusterCol}]))) = ?`;
      params.push(clusterName.trim().toLowerCase());
    }

    // 3️⃣ Sorting
    query += `
      ORDER BY CASE 
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'Krishna' THEN 1
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'Guntur' THEN 2
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'West Godavari' THEN 3
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'Visakhapatnam' THEN 4
        WHEN UPPER(LTRIM(RTRIM([${clusterCol}]))) = 'TOTAL' THEN 5
        ELSE 6
      END
    `;

    const rows = await queryUTIDatabase(query, params);

    if (!rows.length) {
      return res.json({
        success: true,
        headers: [],
        groupedHeaders: [],
        data: [],
        totalRecords: 0,
      });
    }

    // 4️⃣ Convert colX values to displayLabel → value
    const data = rows.map((r) => {
      const obj = {};
      mapping.forEach((m) => {
        const v = r[m.colName];
        obj[m.displayLabel] =
          v === null || v === undefined || v === "" ? "-" : v;
      });
      return obj;
    });

    const totalRecords = rows.length;

    // 5️⃣ Static headers (never grouped)
    const staticMeaning = new Set([
      "branch_code",
      "branch_name",
      "district",
      "cluster",
      "as_on_date", // ACS pie chart meaning
    ]);

    const staticHeaders = mapping
      .filter((m) => m.meaning && staticMeaning.has(m.meaning))
      .map((m) => m.displayLabel);

    // 6️⃣ Intelligent grouping
    const groupBucket = {};

    mapping.forEach((m) => {
      if (m.meaning && staticMeaning.has(m.meaning)) return;

      const label = m.displayLabel.trim();
      const parts = label.split(/\s+/);

      let title, child;

      if (parts.length === 1) {
        title = label;
        child = label;
      } else {
        title = parts.slice(0, -1).join(" ");
        child = parts[parts.length - 1];
      }

      if (!groupBucket[title]) groupBucket[title] = [];
      groupBucket[title].push({ key: m.displayLabel, label: child });
    });

    // 7️⃣ Final groupedHeaders
    const groupedHeaders = [];

    for (const title in groupBucket) {
      const cols = groupBucket[title];

      if (cols.length >= 2) groupedHeaders.push({ title, cols });
      else groupedHeaders.push({ title, cols: [cols[0]], single: true });
    }

    // 8️⃣ Meaning map (required for pie chart selection)
    const meaningMap = {};
    mapping.forEach((m) => {
      if (m.meaning) meaningMap[m.meaning] = m.displayLabel;
    });

    return res.json({
      success: true,
      headers: staticHeaders,
      groupedHeaders,
      data,
      meaningMap,
      totalRecords,
    });
  } catch (err) {
    console.error("❌ Error in /get-advances-summary:", err);
    return res.status(500).json({
      success: false,
      error: "Server error while fetching Advances Summary",
    });
  }
});



//============================================================================================
//                                      NPA SUMMARY (Dynamic Mapping + Grouped Headers)
//============================================================================================

// BACKEND: get-npa-summary (replace existing NPACS route)
app.post("/get-npa-summary", async (req, res) => {
  try {
    const { userId, role, level, clusterName } = req.body;
    const SECTION = "NPACS";

    // 1️⃣ Load mapping
    const mappingRows = await queryUTIDatabase(
      `
      SELECT col_number, display_label, meaning
      FROM MIS.dbo.MIS_Column_Mapping
      WHERE section = ?
      ORDER BY col_number ASC
    `,
      [SECTION]
    );

    if (!mappingRows || !mappingRows.length) {
      return res.status(400).json({
        success: false,
        message: "⚠️ NPA summary mapping not found. Please upload NPACS MIS first.",
      });
    }

    const mapping = mappingRows.map((m) => ({
      colNumber: Number(m.col_number),
      colName: `col${m.col_number}`,
      displayLabel: m.display_label,
      meaning: m.meaning || null,
    }));

    // quick lookup meaning -> colName
    const byMeaning = {};
    mapping.forEach((m) => {
      if (m.meaning) byMeaning[m.meaning] = m.colName;
    });

    const fallback = {
      cluster: "col1",
      branch_code: "col1",
      branch_name: "col2",
      district: "col3",
    };
    const getColByMeaning = (meaning) => byMeaning[meaning] || fallback[meaning] || null;
    const clusterCol = getColByMeaning("cluster") || "col1";

    // Select list based on mapping col names
    const colList = mapping.map((m) => `[${m.colName}]`).join(", ");

    // 2️⃣ Build SQL + level filter
    let query = `SELECT ${colList} FROM [dbo].[NPACS] WHERE 1=1`;
    const params = [];

    if (level === "Level 2" && clusterName && clusterName.trim() !== "") {
      query += ` AND LOWER(LTRIM(RTRIM([${clusterCol}]))) = ?`;
      params.push(clusterName.trim().toLowerCase());
    }

    // 3️⃣ Custom ordering (preserve your current ordering logic)
    query += `
      ORDER BY CASE 
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'Krishna' THEN 1
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'Guntur' THEN 2
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'West Godavari' THEN 3
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'Visakhapatnam' THEN 4
        WHEN UPPER(LTRIM(RTRIM([${clusterCol}]))) = 'TOTAL' THEN 5
        ELSE 6
      END
    `;

    const rows = await queryUTIDatabase(query, params);

    await logActivity(
      userId,
      role,
      "View NPA Summary",
      `Level:${level || "All"}, Cluster:${clusterName || "All"}`
    );

    if (!rows || rows.length === 0) {
      return res.json({
        success: true,
        headers: [],
        groupedHeaders: [],
        data: [],
        meaningMap: {},
        totalRecords: 0,
      });
    }

    // 4️⃣ Map colX -> displayLabel values
    const data = rows.map((r) => {
      const obj = {};
      mapping.forEach((m) => {
        const raw = r[m.colName];
        obj[m.displayLabel] = raw === null || raw === undefined || raw === "" ? "-" : raw;
      });
      return obj;
    });

    // 5️⃣ Static headers (keep cluster/branch/district static)
    const staticSet = new Set(["cluster", "branch_code", "branch_name", "district"]);
    const staticHeaders = mapping
      .filter((m) => m.meaning && staticSet.has(m.meaning))
      .map((m) => m.displayLabel);

    // 6️⃣ Grouped headers — Option B: keep FULL parent titles (including dates / "As on Date")
    // Rule: split last token (A/Cs, Amount, Balance) -> child label; parent = everything before last token
    const groupMap = new Map();

    mapping.forEach((m) => {
      // skip static meanings
      if (m.meaning && staticSet.has(m.meaning)) return;

      const label = (m.displayLabel || "").trim();
      if (!label) return;

      const parts = label.split(/\s+/);
      let parent = label;
      let child = label;
      if (parts.length > 1) {
        child = parts[parts.length - 1]; // A/Cs, Balance, Amount
        parent = parts.slice(0, -1).join(" ");
      }

      if (!groupMap.has(parent)) groupMap.set(parent, { title: parent, cols: [] });

      groupMap.get(parent).cols.push({
        key: m.displayLabel,
        label: child,
      });
    });

    const groupedHeaders = Array.from(groupMap.values());

    // 7️⃣ Build meaningMap: meaning -> displayLabel (frontend will use this to pick pie column)
    const meaningMap = {};
    mapping.forEach((m) => {
      if (m.meaning) meaningMap[m.meaning] = m.displayLabel;
    });

    const totalRecords = rows.length;

    return res.json({
      success: true,
      headers: staticHeaders,
      groupedHeaders,
      data,
      meaningMap,
      totalRecords,
    });
  } catch (err) {
    console.error("❌ Error in /get-npa-summary:", err.stack || err);
    return res.status(500).json({
      success: false,
      error: "Server error while fetching NPA Summary",
    });
  }
});

//============================================================================================
//                             LOANS SANCTIONED SUMMARY (Dynamic Mapping + Grouped Headers)
//============================================================================================

// BACKEND: get-loans-sanctioned-summary (replace existing LSCS route)
app.post("/get-loans-sanctioned-summary", async (req, res) => {
  try {
    const { userId, role, level, clusterName } = req.body;
    const SECTION = "LSCS";

    // 1️⃣ Load mapping
    const mappingRows = await queryUTIDatabase(
      `
      SELECT col_number, display_label, meaning
      FROM MIS.dbo.MIS_Column_Mapping
      WHERE section = ?
      ORDER BY col_number ASC
    `,
      [SECTION]
    );

    if (!mappingRows || !mappingRows.length) {
      return res.status(400).json({
        success: false,
        message: "⚠️ Loans Sanctioned summary mapping not found. Please upload LSCS MIS first.",
      });
    }

    const mapping = mappingRows.map((m) => ({
      colNumber: Number(m.col_number),
      colName: `col${m.col_number}`,
      displayLabel: m.display_label,
      meaning: m.meaning || null,
    }));

    // quick lookup meaning -> colName
    const byMeaning = {};
    mapping.forEach((m) => {
      if (m.meaning) byMeaning[m.meaning] = m.colName;
    });

    const fallback = {
      cluster: "col1",
      branch_code: "col1",
      branch_name: "col2",
      district: "col3",
    };
    const getColByMeaning = (meaning) => byMeaning[meaning] || fallback[meaning] || null;
    const clusterCol = getColByMeaning("cluster") || "col1";

    const colList = mapping.map((m) => `[${m.colName}]`).join(", ");

    // 2️⃣ Build base query + level filter
    let query = `SELECT ${colList} FROM [dbo].[LSCS] WHERE 1=1`;
    const params = [];

    if (level === "Level 2" && clusterName && clusterName.trim() !== "") {
      query += ` AND LOWER(LTRIM(RTRIM([${clusterCol}]))) = ?`;
      params.push(clusterName.trim().toLowerCase());
    }

    // 3️⃣ Ordering
    query += `
      ORDER BY CASE 
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'Krishna' THEN 1
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'Guntur' THEN 2
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'West Godavari' THEN 3
        WHEN LTRIM(RTRIM([${clusterCol}])) = 'Visakhapatnam' THEN 4
        WHEN UPPER(LTRIM(RTRIM([${clusterCol}]))) = 'TOTAL' THEN 5
        ELSE 6
      END
    `;

    const rows = await queryUTIDatabase(query, params);

    await logActivity(
      userId,
      role,
      "View Loans Sanctioned Summary",
      `Level:${level || "All"}, Cluster:${clusterName || "All"}`
    );

    if (!rows || rows.length === 0) {
      return res.json({
        success: true,
        headers: [],
        groupedHeaders: [],
        data: [],
        meaningMap: {},
        totalRecords: 0,
      });
    }

    // 4️⃣ Map colX -> displayLabel values
    const data = rows.map((r) => {
      const obj = {};
      mapping.forEach((m) => {
        const raw = r[m.colName];
        obj[m.displayLabel] = raw === null || raw === undefined || raw === "" ? "-" : raw;
      });
      return obj;
    });

    // 5️⃣ Static headers
    const staticSet = new Set(["cluster", "branch_code", "branch_name", "district"]);
    const staticHeaders = mapping
      .filter((m) => m.meaning && staticSet.has(m.meaning))
      .map((m) => m.displayLabel);

    // 6️⃣ Grouped headers — Option B: keep FULL parent titles (including dates)
    const groupMap = new Map();

    mapping.forEach((m) => {
      if (m.meaning && staticSet.has(m.meaning)) return;

      const label = (m.displayLabel || "").trim();
      if (!label) return;

      const parts = label.split(/\s+/);
      let parent = label;
      let child = label;

      if (parts.length > 1) {
        child = parts[parts.length - 1];
        parent = parts.slice(0, -1).join(" ");
      }

      if (!groupMap.has(parent)) groupMap.set(parent, { title: parent, cols: [] });

      groupMap.get(parent).cols.push({
        key: m.displayLabel,
        label: child,
      });
    });

    const groupedHeaders = Array.from(groupMap.values());

    // meaningMap for frontend pie selection (meaning -> displayLabel)
    const meaningMap = {};
    mapping.forEach((m) => {
      if (m.meaning) meaningMap[m.meaning] = m.displayLabel;
    });

    const totalRecords = rows.length;

    return res.json({
      success: true,
      headers: staticHeaders,
      groupedHeaders,
      data,
      meaningMap,
      totalRecords,
    });
  } catch (err) {
    console.error("❌ Error in /get-loans-sanctioned-summary:", err.stack || err);
    return res.status(500).json({
      success: false,
      error: "Server error while fetching Loans Sanctioned Summary",
    });
  }
});
//=================================================================================
//=================================================================================
//                                  GRAPHS
//=================================================================================
//=================================================================================


// ========================================================================
//   SIMPLE DEPOSITS → BRANCH LIST BY CLUSTER (For Level-2 Bar Chart)
// ========================================================================
app.post("/get-deposits-by-cluster", async (req, res) => {
  try {
    const { clusterName, userId, role } = req.body;

    if (!clusterName) {
      return res.status(400).json({
        success: false,
        message: "Cluster name missing",
      });
    }

    const SECTION = "Deposits";

    // 1️⃣ Load column mapping
    const mappingRows = await queryUTIDatabase(
      `
        SELECT col_number, display_label, meaning
        FROM MIS.dbo.MIS_Column_Mapping
        WHERE section = ?
        ORDER BY col_number ASC
      `,
      [SECTION]
    );

    if (!mappingRows.length) {
      return res.status(400).json({
        success: false,
        message: "Deposits mapping missing",
      });
    }

    const mapping = mappingRows.map((m) => ({
      colName: `col${m.col_number}`,
      displayLabel: m.display_label,
      meaning: m.meaning,
    }));

    // Build meaning lookup
    const meaningToCol = {};
    mapping.forEach((m) => {
      if (m.meaning) meaningToCol[m.meaning] = m.colName;
    });

    const branchNameCol = meaningToCol["branch_name"] || "col2";
    const clusterCol = meaningToCol["cluster"] || "col4";
    const valueCol =
      meaningToCol["as_on_date_deposits"] ||
      meaningToCol["as_on_date"] ||
      "col5";

    // Build final col list
    const colList = mapping.map((m) => `[${m.colName}]`).join(",");

    const sql = `
      SELECT ${colList}
      FROM MIS.dbo.[Deposits]
      WHERE LOWER(LTRIM(RTRIM([${clusterCol}]))) = ?
      ORDER BY TRY_CAST([${branchNameCol}] AS NVARCHAR(200))
    `;

    const rows = await queryUTIDatabase(sql, [clusterName.toLowerCase()]);

    // Convert DB row → frontend friendly
    const data = rows.map((r) => {
      const obj = {};
      mapping.forEach((m) => {
        obj[m.displayLabel] = r[m.colName];
      });
      return obj;
    });

    return res.json({
      success: true,
      data,
      branchNameKey: mapping.find((m) => m.meaning === "branch_name")?.displayLabel,
      valueKey: mapping.find((m) => m.meaning === "as_on_date_deposits")?.displayLabel,
    });
  } catch (err) {
    console.error("❌ Error in /get-deposits-by-cluster:", err);
    return res.status(500).json({
      success: false,
      message: "Error fetching deposits by cluster",
    });
  }
});


//============================================================================================
//      SMA — Meaning-Based Dynamic Sorting (FINAL + ADVANCES JOIN + ROW %)
//============================================================================================

app.post("/get-sma-denormalized", async (req, res) => {
  try {
    let {
      userId,
      role,
      level,
      designation,
      filters = {},
      page = 1,
      pageSize = 10,
      sortBy,
      sortOrder = "ASC",
    } = req.body;

    console.log("🔹 /get-sma-denormalized called");

    if (userId) {
      userId = parseInt(userId, 10);
    }

    const SECTION = "SMA";

    // ------------------------------------------------------
    // Normalize Filters
    // ------------------------------------------------------
    let branchCode = (filters.branchCode ?? filters.branch_code ?? "").trim();
    let branchName = (filters.branchName ?? filters.branch_name ?? "").trim();
    let districtName = (filters.districtName ?? filters.district ?? "").trim();
    let clusterName = (filters.clusterName ?? filters.cluster ?? "").trim();

    // ⭐⭐⭐ LEVEL-1 MUST GET branch_code FROM USERS TABLE
    if (level === "Level 1") {
      const userBranch = await queryUTIDatabase(
        `SELECT [Br Code] FROM dbo.employees WHERE [Emp No.] = ?`,
        [userId]
      );

      if (!userBranch.length) {
        return res.status(400).json({
          success: false,
          message: "User branch not found for Level 1",
        });
      }

      branchCode = (userBranch[0]["Br Code"] || "").toString().trim();

      // Normalize branchCode "2.00" → "2"
      const num = parseFloat(branchCode);
      if (!isNaN(num)) {
        branchCode = parseInt(num, 10).toString();
      }

      console.log("✔ Level-1 Branch Code Applied:", branchCode);
    }

    level = level || "Level 1";
    page = Number(page);
    pageSize = Number(pageSize);
    const offset = (page - 1) * pageSize;

    // ⭐⭐⭐ LEVEL-1 FIX: Normalise branchCode "2.00" → "2"
    if (level === "Level 1" && branchCode) {
      const num = parseFloat(branchCode);
      if (!isNaN(num)) branchCode = parseInt(num, 10).toString();
    }

    //========================================================================================
    // 1️⃣ LOAD SMA MAPPING (meaning → colX)
    //========================================================================================
    const mappingRows = await queryUTIDatabase(
      `
      SELECT col_number, display_label, meaning, detected_type
      FROM MIS.dbo.MIS_Column_Mapping
      WHERE section = ?
      ORDER BY col_number ASC
      `,
      [SECTION]
    );

    if (!mappingRows.length) {
      return res.status(400).json({
        success: false,
        message: "SMA mapping not found. Upload SMA MIS first.",
      });
    }

    const mapping = mappingRows.map((m) => ({
      colNumber: Number(m.col_number),
      colName: `col${m.col_number}`,
      displayLabel: m.display_label,
      meaning: m.meaning,
      type: m.detected_type,
    }));

    const meaningToCol = {};
    mapping.forEach((m) => {
      if (m.meaning) meaningToCol[m.meaning] = m.colName;
    });

    const fallback = {
      branch_code: "col1",
      branch_name: "col2",
      district: "col3",
      cluster: "col4",
    };

    const getCol = (m) => meaningToCol[m] || fallback[m];

    //========================================================================================
    // 2️⃣ LOAD ADVANCES MAPPING (for as_on_date_advances)
    //========================================================================================
    const advancesMapRows = await queryUTIDatabase(
      `
      SELECT col_number, meaning
      FROM MIS.dbo.MIS_Column_Mapping
      WHERE section = 'Advances'
      `
    );

    const advancesMeaningToCol = {};
    advancesMapRows.forEach((m) => {
      if (m.meaning) advancesMeaningToCol[m.meaning] = `col${m.col_number}`;
    });

    const advBranchCol = advancesMeaningToCol["branch_code"] || "col1";
    const advAsOnCol =
      advancesMeaningToCol["as_on_date_advances"] ||
      advancesMeaningToCol["as_on_date"] ||
      null;

    if (!advAsOnCol) {
      console.warn(
        "⚠️ as_on_date_advances meaning not found in Advances mapping. SMA% will not work."
      );
    }

    //========================================================================================
    // 3️⃣ LOAD SORTABLE COLUMNS
    //========================================================================================
    const sortableRows = await queryUTIDatabase(
      `SELECT meaning FROM MIS.dbo.MIS_Sortable_Columns WHERE section = ?`,
      [SECTION]
    );

    const allowedSortableMeaning = sortableRows
      .map((r) => r.meaning)
      .filter(Boolean);

    //========================================================================================
    // 4️⃣ BUILD BASE QUERY (SMA + LEFT JOIN Advances)
    //========================================================================================
    const colList = mapping.map((m) => `s.[${m.colName}]`).join(", ");

    const extraAdvSelect = advAsOnCol
      ? `, a.[${advAsOnCol}] AS as_on_date_advances`
      : "";

    let query = `
      SELECT ${colList}${extraAdvSelect}
      FROM SMA s
      ${
        advAsOnCol
          ? `
      LEFT JOIN Advances a
        ON TRY_CAST(LTRIM(RTRIM(s.[${getCol("branch_code")}])) AS DECIMAL(18,4)) =
           TRY_CAST(LTRIM(RTRIM(a.[${advBranchCol}])) AS DECIMAL(18,4))
      `
          : ""
      }
      WHERE 1=1
    `;

    let countQuery = `
      SELECT COUNT(*) AS totalRecords
      FROM SMA s
      WHERE 1=1
    `;

    const params = [];
    const countParams = [];

    const addFilter = (meaning, value, exact = false) => {
      if (!value) return;

      const col = getCol(meaning);
      if (!col) return;

      const qualifiedCol = `s.[${col}]`;

      if (meaning === "branch_code") {
        query += ` AND TRY_CAST(LTRIM(RTRIM(${qualifiedCol})) AS DECIMAL(18,4)) = TRY_CAST(? AS DECIMAL(18,4))`;
        countQuery += ` AND TRY_CAST(LTRIM(RTRIM(${qualifiedCol})) AS DECIMAL(18,4)) = TRY_CAST(? AS DECIMAL(18,4))`;
        params.push(value);
        countParams.push(value);
        return;
      }

      if (exact) {
        query += ` AND LTRIM(RTRIM(${qualifiedCol})) = ?`;
        countQuery += ` AND LTRIM(RTRIM(${qualifiedCol})) = ?`;
        params.push(value.trim());
        countParams.push(value.trim());
      } else {
        query += ` AND LOWER(${qualifiedCol}) LIKE ?`;
        countQuery += ` AND LOWER(${qualifiedCol}) LIKE ?`;
        params.push(`%${value.toLowerCase()}%`);
        countParams.push(`%${value.toLowerCase()}%`);
      }
    };

    //========================================================================================
    // 5️⃣ LEVEL-BASED SECURITY
    //========================================================================================
    let effectiveCluster = clusterName;

    if (level === "Level 2" && designation) {
      const match = designation.match(/Cluster\s*Head\s*[-:]\s*(.*)/i);
      if (match?.[1]) effectiveCluster = match[1].trim();
    }

    if (level === "Level 2") {
      if (!effectiveCluster) {
        return res.status(400).json({
          success: false,
          message: "Cluster missing for Level 2 user",
        });
      }
      addFilter("cluster", effectiveCluster, true);
    }

    if (level === "Level 1") {
      addFilter("branch_code", branchCode, true);
    }

    if (branchCode && level !== "Level 1") addFilter("branch_code", branchCode);
    if (branchName && level !== "Level 1") addFilter("branch_name", branchName);
    if (districtName) addFilter("district", districtName);
    if (clusterName && level !== "Level 2") addFilter("cluster", clusterName);

    //========================================================================================
    // 6️⃣ SORTING
    //========================================================================================
    let orderCol = getCol("branch_code");

    if (sortBy && allowedSortableMeaning.includes(sortBy)) {
      orderCol = getCol(sortBy);
    }

    const orderDir =
      sortOrder.toUpperCase() === "DESC" ? "DESC" : "ASC";

    query += `
      ORDER BY TRY_CAST(LTRIM(RTRIM(s.[${orderCol}])) AS DECIMAL(18,4)) ${orderDir}
    `;

    //========================================================================================
    // 7️⃣ PAGINATION
    //========================================================================================
    query += ` OFFSET ? ROWS FETCH NEXT ? ROWS ONLY`;
    params.push(offset, pageSize);

    //========================================================================================
    // 8️⃣ EXECUTE
    //========================================================================================
    const rows = await queryUTIDatabase(query, params);
    const totalRecords =
      (await queryUTIDatabase(countQuery, countParams))[0]?.totalRecords || 0;

    const totalPages = Math.ceil(totalRecords / pageSize);

    // Helper for numeric safe
    const safeNum = (v) => {
  if (v === null || v === undefined || v === "" || v === "-" || v === 0) return 0;

  // Clean formatting junk
  let cleaned = String(v)
    .replace(/[^\d.\-]/g, "")  // remove commas, ₹, spaces, unicode minus
    .replace(/\.{2,}/g, ".")   // fix double dots
    .trim();

  let num = Number(cleaned);
  return isNaN(num) ? 0 : num;
};


    //========================================================================================
    // 9️⃣ CLEAN TOTAL ROWS + BUILD DATA OBJECT + ROW %
//========================================================================================
    const skipWords = ["total"];

    const data = rows
      .map((row) => {
        const bc = (row[getCol("branch_code")] + "").toUpperCase();
        const bn = (row[getCol("branch_name")] + "").toUpperCase();

        if (bc.includes("TOTAL") || bn.includes("TOTAL")) return null;

        const obj = {};

        mapping.forEach((m) => {
          if (skipWords.some((w) => m.displayLabel.toLowerCase().includes(w))) return;

          const raw = row[m.colName];
          obj[m.displayLabel] =
            raw && typeof raw === "object" ? "-" : raw ?? "-";
        });

        // 👉 Extra field from Advances
        if (typeof row.as_on_date_advances !== "undefined") {
          obj.as_on_date_advances = row.as_on_date_advances;
        }

        // 👉 Row-wise SMA% of Advances (backend only)
        const advVal = safeNum(row.as_on_date_advances);

        const sma0 = safeNum(row[meaningToCol["sma0_amount"]]);
        const sma1 = safeNum(row[meaningToCol["sma1_amount"]]);
        const sma2 = safeNum(row[meaningToCol["sma2_amount"]]);

        const calcPct = (sma, adv) => {
  sma = safeNum(sma);
  adv = safeNum(adv);

  if (adv <= 0) return 0;

  return Number(((sma / adv) * 100).toFixed(2));
};


        obj.sma0_percent = calcPct(sma0, advVal);
        obj.sma1_percent = calcPct(sma1, advVal);
        obj.sma2_percent = calcPct(sma2, advVal);

        return obj;
      })
      .filter(Boolean);

    //========================================================================================
    // 🔟 GROUP HEADERS (unchanged)
//========================================================================================
    const skipMeaning = new Set([
      "branch_code",
      "branch_name",
      "district",
      "cluster",
    ]);

    const groupMap = new Map();

    mapping.forEach((m) => {
      if (skipMeaning.has(m.meaning)) return;
      if (skipWords.some((w) => m.displayLabel.toLowerCase().includes(w))) return;

      const parts = m.displayLabel.split(" ");
      const title =
        parts.length === 1 ? m.displayLabel : parts.slice(0, -1).join(" ");
      const child =
        parts.length === 1 ? m.displayLabel : parts.slice(-1)[0];

      if (!groupMap.has(title)) groupMap.set(title, { title, cols: [] });

      groupMap.get(title).cols.push({ key: m.displayLabel, label: child });
    });

    const groupedHeaders = Array.from(groupMap.values());

    //========================================================================================
    // 1️⃣1️⃣ RESPONSE
    //========================================================================================
    return res.json({
      success: true,
      data,
      groupedHeaders,
      totalRecords,
      totalPages,
      fixedCluster: effectiveCluster || null,
      unit: "lakhs",
    });
  } catch (err) {
    console.error("❌ SMA Error:", err);
    return res.status(500).json({
      success: false,
      message: "Server error fetching SMA data",
      error: err.message,
    });
  }
});


//============================================================================================
//                            SMA TOTAL (MEANING-BASED + FINAL)
//============================================================================================

app.post("/sma-total", async (req, res) => {
  try {
    const { level, clusterName } = req.body;

    const cleanLevel = (level || "").trim().toLowerCase();
    const cleanCluster = (clusterName || "").trim();

    // Safe numeric conversion (no formatting here)
    const safe = (v) => {
      if (v === null || v === undefined) return 0;
      const n = parseFloat(String(v).replace(/,/g, ""));
      return isNaN(n) ? 0 : n;
    };

    //========================================================================================
    // 1️⃣ LOAD SMA COLUMN MAPPING
    //========================================================================================
    const mappingSMA = await queryUTIDatabase(`
      SELECT col_number, meaning
      FROM MIS.dbo.MIS_Column_Mapping
      WHERE section = 'SMA'
    `);

    if (!mappingSMA.length) {
      return res.json({
        success: false,
        message: "SMA mapping missing. Upload SMA MIS first.",
      });
    }

    const meaningToCol = {};
    mappingSMA.forEach((m) => {
      if (m.meaning) meaningToCol[m.meaning] = `col${m.col_number}`;
    });

    // Required meanings
    const SMA_BAL = ["sma0_amount", "sma1_amount", "sma2_amount", "total_balance"];
    const SMA_AC = ["sma0_accounts", "sma1_accounts", "sma2_accounts", "total_accounts"];

    // Validate
    for (let m of [...SMA_BAL, ...SMA_AC]) {
      if (!meaningToCol[m]) {
        return res.json({
          success: false,
          message: `Meaning '${m}' missing in SMA mapping.`,
        });
      }
    }

    const c = meaningToCol;

    //========================================================================================
    // 2️⃣ LOAD ACS MAPPING (to fetch total advances)
    //========================================================================================
    const mappingACS = await queryUTIDatabase(`
      SELECT col_number, meaning
      FROM MIS.dbo.MIS_Column_Mapping
      WHERE section = 'ACS'
    `);

    const meaningACS = {};
    mappingACS.forEach((m) => {
      if (m.meaning) meaningACS[m.meaning] = `col${m.col_number}`;
    });

    // ⚠️ This column in ACS is ALREADY in Lakhs
    const advancesCol =
      meaningACS["as_on_date_advances"] ||
      meaningACS["as_on_date"] ||
      meaningACS["total_advances"] ||
      "col5";

    let totalAdvances = 0;

    //========================================================================================
    // 3️⃣ FETCH TOTAL ADVANCES (DYNAMIC)  — already in LAKHS
    //========================================================================================
    if (cleanLevel === "grandtotal") {
      const q = `
        SELECT TOP 1 [${advancesCol}] AS adv
        FROM ACS
        WHERE UPPER(LTRIM(RTRIM(col1))) LIKE '%TOTAL%'
      `;
      const r = await queryUTIDatabase(q);
      totalAdvances = safe(r?.[0]?.adv); // lakhs
    } else {
      const q = `
        SELECT TOP 1 [${advancesCol}] AS adv
        FROM ACS
        WHERE UPPER(LTRIM(RTRIM(col1))) = UPPER(?)
      `;
      const r = await queryUTIDatabase(q, [cleanCluster]);
      totalAdvances = safe(r?.[0]?.adv); // lakhs
    }

    if (!totalAdvances) {
      return res.json({
        success: false,
        message: "Total Advances not found in ACS.",
      });
    }

    //========================================================================================
    // 4️⃣ SUM QUERY FOR SMA  (SMA balances are in RUPEES)
    //========================================================================================
    const allSums = [...SMA_BAL, ...SMA_AC]
      .map((m) => `SUM(TRY_CAST([${c[m]}] AS DECIMAL(18,2))) AS [${m}]`)
      .join(",");

    let smaRows = [];

    if (cleanLevel === "grandtotal") {
      // include all numeric branches
      const q = `
        SELECT ${allSums}
        FROM SMA
        WHERE ISNUMERIC([${c.branch_code}]) = 1
      `;
      smaRows = await queryUTIDatabase(q);
    } else if (["level 2", "level 3"].includes(cleanLevel)) {
      const q = `
        SELECT ${allSums}
        FROM SMA
        WHERE UPPER(LTRIM(RTRIM([${c.cluster}]))) = UPPER(?)
          AND ISNUMERIC([${c.branch_code}]) = 1
      `;
      smaRows = await queryUTIDatabase(q, [cleanCluster]);
    }

    if (!smaRows.length) {
      return res.json({
        success: false,
        message: "No SMA data found.",
      });
    }

    const row = smaRows[0];

    //========================================================================================
    // 5️⃣ FORM FINAL OUTPUT
    //     - SMA amounts → convert RUPEES → LAKHS
    //     - ACS total_advances stays as-is (already LAKHS)
    //     - % = (SMA_LAKHS / ACS_LAKHS) * 100
    //========================================================================================
    const result = { total_advances: totalAdvances }; // lakhs

    // Balances + % (ALL IN LAKHS)
    SMA_BAL.forEach((key) => {
      const amountRupees = safe(row[key]);        // from SMA table
      const amountLakhs = amountRupees / 100000;  // convert to lakhs

      result[key] = amountLakhs;

      // percentage in positive value as per your earlier requirement
      const pct =
        totalAdvances > 0 ? Math.abs((amountLakhs / totalAdvances) * 100) : 0;

      result[`${key}_percent`] = pct; // keep as NUMBER
    });

    // Accounts (int only)
    SMA_AC.forEach((key) => {
      result[key] = Math.abs(Math.round(safe(row[key]))); // positive integer
    });

    //========================================================================================
    // 6️⃣ RESPONSE
    //========================================================================================
    return res.json({
      success: true,
      level,
      cluster: cleanCluster || "ALL",
      data: result,
    });
  } catch (err) {
    console.error("❌ SMA TOTAL ERROR:", err);
    return res.status(500).json({
      success: false,
      message: "Server error in SMA Total",
      error: err.message,
    });
  }
});




//============================================================================================
//                             UPDATE DAILY SUMMARY  (MEANING-AWARE)
//============================================================================================
app.post("/update-daily-summary", async (req, res) => {
  try {
    console.log("📩 Incoming request: /update-daily-summary");

    // Step 1️⃣: Run dynamic procedure (uses MIS_Column_Mapping meanings internally)
    await queryUTIDatabase("EXEC update_daily_summary_from_history");
    console.log("✅ Stored procedure executed successfully.");

    // Step 2️⃣: Get the SECOND-LATEST DISTINCT upload date from daily_summary_history
    const prevWorkingDateQuery = `
      ;WITH dates AS (
        SELECT DISTINCT CAST(upload_timestamp AS DATE) AS dt
        FROM daily_summary_history
      ),
      ranked AS (
        SELECT dt, ROW_NUMBER() OVER (ORDER BY dt DESC) AS rn
        FROM dates
      )
      SELECT dt AS prev_working_date FROM ranked WHERE rn = 2;
    `;

    const prevRes = await queryUTIDatabase(prevWorkingDateQuery);
    const prevWorkingDate = prevRes?.[0]?.prev_working_date || null;

    if (!prevWorkingDate) {
      console.warn("⚠️ No previous working day found.");
      return res.json({
        success: true,
        message:
          "Only one upload batch exists. Cannot determine previous working day.",
        lastUploadDate: null,
      });
    }

    // Step 3️⃣: Ensure tracking table exists
    const createTableQuery = `
      IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='last_summary_update' AND xtype='U')
      CREATE TABLE last_summary_update (
        id INT IDENTITY(1,1) PRIMARY KEY,
        last_upload_date DATE NOT NULL,
        updated_at DATETIME DEFAULT GETDATE()
      );
    `;
    await queryUTIDatabase(createTableQuery);

    // Step 4️⃣: Insert/Update last business date (using ? placeholder like rest of code)
    const upsertQuery = `
      IF EXISTS (SELECT 1 FROM last_summary_update)
        UPDATE last_summary_update
        SET last_upload_date = ?, updated_at = GETDATE();
      ELSE
        INSERT INTO last_summary_update (last_upload_date)
        VALUES (?);
    `;

    await queryUTIDatabase(upsertQuery, [prevWorkingDate, prevWorkingDate]);

    console.log(
      `📅 Stored business date (2nd latest distinct): ${prevWorkingDate}`
    );

    return res.json({
      success: true,
      message: "Daily summary updated successfully.",
      lastUploadDate: prevWorkingDate,
    });
  } catch (error) {
    console.error("❌ Error updating daily summary:", error);
    return res.status(500).json({
      success: false,
      message: "Failed to update daily summary",
      error: error.message,
    });
  }
});

// ==============================================
//   DAILY SUMMARY → Meaning → Column Name Map
//   (Used by formatWithMeanings formatter)
// ==============================================
const DAILY_SUMMARY_MEANINGS = {
  branch_code: "branch_code",
  branch_name: "branch_name",
  district: "district",
  cluster: "cluster",

  total_deposits: "total_deposits",
  total_advances: "total_advances",

  tdr_accounts: "tdr_accounts",
  rd_accounts: "rd_accounts",
  savings_accounts: "savings_accounts",

  stamped_npa_accounts: "stamped_npa_accounts",
  stamped_npa_balance: "stamped_npa_balance",
  upgraded_npa_accounts: "upgraded_npa_accounts",
  downgraded_npa_accounts: "downgraded_npa_accounts",

  sma_total_accounts: "sma_total_accounts",
  sma_total_balance: "sma_total_balance",
};

// ==================== Column Mapping (with UTI connection + meanings) ======================
app.post("/api/filters", async (req, res) => {
  const { reportType } = req.body;

  try {
    let table = "";
    let section = "";

    if (reportType === "daily_summary") {
      table = "daily_summary"; // real columns
    } else if (reportType === "deposits") {
      table = "Deposits";
      section = "Deposits";
    } else if (reportType === "deposit_accounts") {
      table = "deposits_accounts_opened";
      section = "deposits_accounts_opened";
    } else {
      return res
        .status(400)
        .json({ success: false, message: "Invalid report type" });
    }

    // ▶ Case 1: DAILY SUMMARY (already real columns, not col1..col30)
    if (reportType === "daily_summary") {
      const query = `
        SELECT DISTINCT
          branch_code,
          branch_name,
          cluster AS cluster_name,
          district AS district_name
        FROM [dbo].[daily_summary]
        WHERE ISNUMERIC(branch_code) = 1
      `;
      const rows = await queryUTIDatabase(query);
      return res.json({ success: true, filters: rows });
    }

    // ▶ Case 2: Deposits / Deposits Accounts → use MIS_Column_Mapping (meaning → colX)
    // We need colX for branch_code, branch_name, cluster, district
    const mappingRows = await queryUTIDatabase(
      `
        SELECT col_number, meaning
        FROM MIS.dbo.MIS_Column_Mapping
        WHERE section = ?
      `,
      [section]
    );

    const meaningToCol = {};
    (mappingRows || []).forEach((m) => {
      if (m.meaning) {
        meaningToCol[m.meaning] = `col${m.col_number}`;
      }
    });

    // Fallback for safety if mapping incomplete
    const fallback = {
      branch_code: "col1",
      branch_name: "col2",
      district: "col3",
      cluster: "col4",
    };

    const getCol = (meaning) => meaningToCol[meaning] || fallback[meaning];

    const colBranchCode = getCol("branch_code");
    const colBranchName = getCol("branch_name");
    const colDistrict = getCol("district");
    const colCluster = getCol("cluster");

    const query = `
      SELECT DISTINCT
        LTRIM(RTRIM([${colBranchCode}])) AS branch_code,
        LTRIM(RTRIM([${colBranchName}])) AS branch_name,
        LTRIM(RTRIM([${colCluster}])) AS cluster_name,
        LTRIM(RTRIM([${colDistrict}])) AS district_name
      FROM [dbo].[${table}]
      WHERE ISNUMERIC([${colBranchCode}]) = 1
    `;

    const rows = await queryUTIDatabase(query);

    return res.json({ success: true, filters: rows });
  } catch (error) {
    console.error("Filter fetch error:", error);
    res
      .status(500)
      .json({ success: false, message: "Error fetching filters" });
  }
});
// ============================================================
//                   DAILY SUMMARY (MEANING-DRIVEN)
// ============================================================
app.post("/get-daily-summary", async (req, res) => {
  try {
    const { userId, level, filters = {}, designation, sortBy, sortOrder } =
      req.body;

    /* ---------------------------------------------------
       CLEAN NUMBER UTILITY
    --------------------------------------------------- */
    function cleanNumber(value) {
      if (value == null) return null;
      return String(value)
        .replace(/,/g, "")
        .replace(/\.00$/, "")
        .replace(/^0+/, "")
        .trim();
    }

    const branchCode = cleanNumber(filters.branchCode || "");
    const cleanUserId = cleanNumber(userId);
    const branchName = (filters.branchName || "").trim();
    const clusterName = (filters.clusterName || "").trim();
    const districtName = (filters.districtName || "").trim();

    const effectiveLevel = level || "Level 1";
    let effectiveCluster = clusterName || "";

    /* ---------------------------------------------------
       LOAD SORTABLE MEANINGS
    --------------------------------------------------- */
    const sortableRows = await queryUTIDatabase(
      `
        SELECT display_label, meaning
        FROM MIS.dbo.MIS_Sortable_Columns
        WHERE section = ?
        ORDER BY display_label ASC
      `,
      ["DailySummary"]
    );

    const allowedSortableMeaning = (sortableRows || [])
      .map((r) => r.meaning)
      .filter((m) => m);

    /* ---------------------------------------------------
       BUILD WHERE CLAUSES
    --------------------------------------------------- */
    const whereClauses = ["1=1"];
    const params = [];

    /* ---------------------------------------------------
       LEVEL 1 → Only own branch
    --------------------------------------------------- */
    if (effectiveLevel === "Level 1") {
      whereClauses.push(`
        branch_code = (
          SELECT TOP 1 [Br Code]
          FROM employees
          WHERE [Emp No.] = CAST(? AS INT)
        )
      `);
      params.push(cleanUserId);

      // LEVEL 1 SHOULD NEVER SEE totals
      whereClauses.push("is_total = 0");
    }

    /* ---------------------------------------------------
       LEVEL 2 → Cluster Head
    --------------------------------------------------- */
    if (effectiveLevel === "Level 2") {
      // Extract cluster from designation
      if (designation?.includes("Cluster Head-")) {
        effectiveCluster = designation.replace("Cluster Head-", "").trim();
      }

      if (!effectiveCluster) {
        return res.status(400).json({
          success: false,
          message: "Cluster not found in designation",
        });
      }

      whereClauses.push("LOWER(cluster) = LOWER(?)");
      params.push(effectiveCluster);

      // VALIDATE BRANCH BELONGS TO CLUSTER
      if (branchCode) {
        const chk = await queryUTIDatabase(
          "SELECT TOP 1 cluster FROM daily_summary WHERE branch_code = CAST(? AS INT)",
          [branchCode]
        );

        if (
          chk.length === 0 ||
          chk[0].cluster.trim().toLowerCase() !==
            effectiveCluster.toLowerCase()
        ) {
          return res.status(403).json({
            success: false,
            message: `Branch ${branchCode} does not belong to ${effectiveCluster} cluster`,
          });
        }
      }

      // LEVEL 2 SHOULD SEE BOTH:
      //  • normal rows (is_total = 0)
      //  • ONE total row for that cluster (is_total = 1)
      // So we DO NOT filter is_total here.
    }

    /* ---------------------------------------------------
       LEVEL 3 → All branches, but NO totals
    --------------------------------------------------- */
    if (effectiveLevel === "Level 3") {
      whereClauses.push("is_total = 0");
    }

    /* ---------------------------------------------------
       ADDITIONAL FILTERS (except Level 1 restriction)
    --------------------------------------------------- */
    if (branchCode && effectiveLevel !== "Level 1") {
      whereClauses.push("branch_code = ?");
      params.push(branchCode);
    }

    if (branchName && effectiveLevel !== "Level 1") {
      whereClauses.push("LOWER(branch_name) LIKE ?");
      params.push(`%${branchName.toLowerCase()}%`);
    }

    if (districtName) {
      whereClauses.push("LOWER(district) LIKE ?");
      params.push(`%${districtName.toLowerCase()}%`);
    }

    if (clusterName && effectiveLevel !== "Level 2") {
      whereClauses.push("LOWER(cluster) LIKE ?");
      params.push(`%${clusterName.toLowerCase()}%`);
    }

    /* ---------------------------------------------------
       SORTING
    --------------------------------------------------- */
    const sortDir =
      sortOrder && sortOrder.toUpperCase() === "DESC" ? "DESC" : "ASC";

    let orderBy = `
      -- always push totals to bottom
      is_total ASC,
      TRY_CAST(FLOOR(TRY_CAST(branch_code AS FLOAT)) AS INT) ASC
    `;

    if (sortBy && allowedSortableMeaning.includes(sortBy)) {
      orderBy = `
        is_total ASC,
        ISNULL(TRY_CAST(REPLACE(LTRIM(RTRIM(${sortBy})), ',', '') AS FLOAT), 0) ${sortDir}
      `;
    }

    /* ---------------------------------------------------
       FINAL QUERY
    --------------------------------------------------- */
    const query = `
      SELECT
        branch_code,
        branch_name,
        district,
        cluster,
        total_deposits,
        total_advances,
        tdr_accounts,
        rd_accounts,
        savings_accounts,
        stamped_npa_accounts,
        stamped_npa_balance,
        upgraded_npa_accounts,
        downgraded_npa_accounts,
        sma_total_accounts,
        sma_total_balance,
        is_total
      FROM daily_summary
      WHERE ${whereClauses.join(" AND ")}
      ORDER BY ${orderBy}
    `;

    const rows = await queryUTIDatabase(query, params);

    /* ---------------------------------------------------
       HEADERS
    --------------------------------------------------- */
    const headers = [
      { key: "branch_code", label: "Branch Code" },
      { key: "branch_name", label: "Branch Name" },
      { key: "district", label: "District" },
      { key: "cluster", label: "Cluster" },

      { key: "total_deposits", label: "Total Deposits" },

      { key: "tdr_accounts", label: "TDR Accounts" },
      { key: "rd_accounts", label: "RD Accounts" },
      { key: "savings_accounts", label: "Savings Accounts" },

      { key: "total_advances", label: "Total Advances" },

      { key: "stamped_npa_accounts", label: "Stamped NPA A/Cs" },
      { key: "stamped_npa_balance", label: "Stamped NPA Balance" },
      { key: "upgraded_npa_accounts", label: "Upgraded NPA A/Cs" },
      { key: "downgraded_npa_accounts", label: "Downgraded NPA A/Cs" },

      { key: "sma_total_accounts", label: "SMA Total Accounts" },
      { key: "sma_total_balance", label: "SMA Total Balance" },
    ];

    /* ---------------------------------------------------
       RETURN JSON
    --------------------------------------------------- */
    return res.json({
      success: true,
      headers,
      data: rows,
      fixedCluster: effectiveCluster || null,
      sortableColumns: sortableRows.map((r) => ({
        key: r.meaning,
        label: r.display_label,
      })),
    });
  } catch (err) {
    console.error("❌ Error (get-daily-summary):", err);
    return res.status(500).json({
      success: false,
      message: "Server error fetching daily summary",
      error: err.message,
    });
  }
});

// ==================== ✅ Dynamic Sortable Columns Endpoint (DailySummary) ====================
// ==================== ✅ Dynamic Sortable Columns Endpoint (DailySummary) ====================
app.get("/get-daily-summary-columns", async (req, res) => {
  try {
    const rows = await queryUTIDatabase(
      `
        SELECT display_label, meaning
        FROM MIS.dbo.MIS_Sortable_Columns
        WHERE section = ?
        ORDER BY display_label ASC
      `,
      ["DailySummary"]
    );

    const columns = (rows || [])
      .filter((r) => r.meaning)
      .map((r) => ({
        key: r.meaning,         // meaning used by frontend
        label: r.display_label, // UI label
      }));

    return res.json({ success: true, columns });
  } catch (err) {
    console.error("❌ Error fetching daily summary columns:", err);
    return res.status(500).json({
      success: false,
      message: "Failed to fetch column list",
      error: err.message,
    });
  }
});


// ===================================================================================
//                 ⭐ DAILY SUMMARY TREND (MEANING-DRIVEN & UI MAPPED)
// ===================================================================================
function cleanNumber(value) {
  if (!value) return "";
  return String(value)
    .replace(/,/g, "")
    .replace(/\.00$/, "")
    .replace(/^0+/, "")
    .trim();
}

app.post("/get-daily-summary-trend", async (req, res) => {
  try {
    const { userId, level, filters = {}, designation } = req.body;

    const cleanUserId = cleanNumber(userId);
    const cleanBranchCode = cleanNumber(filters.branchCode);
    const cleanCluster = (filters.clusterName || "").trim();

    // Load all trends from SP
    const rows = await queryUTIDatabase("EXEC sp_get_daily_summary_trend");

    if (!rows || rows.length === 0) {
      return res.json([]);
    }

    let filteredRows = [...rows]; // default full data

    // LEVEL 1 → ONLY user's exact branch
    if (level === "Level 1") {
      const userBranch = await queryUTIDatabase(`
        SELECT TOP 1 CAST([Br Code] AS INT) AS branch_code
        FROM employees
        WHERE [Emp No.] = CAST(? AS INT)
      `, [cleanUserId]);

      if (userBranch.length > 0) {
        const branch = userBranch[0].branch_code;
        filteredRows = rows.filter(r => parseInt(r.branch_code) === branch);
      }
    }

    // LEVEL 2 → filter by cluster
    if (level === "Level 2") {
      let cluster = cleanCluster;

      if (designation?.includes("Cluster Head-")) {
        cluster = designation.replace("Cluster Head-", "").trim();
      }

      filteredRows = rows.filter(
        (r) => (r.cluster || "").trim().toLowerCase() === cluster.toLowerCase()
      );
    }

    // Trend key mapping
    const trendMap = {
      total_deposits: "total_deposits_trend",
      total_advances: "total_advances_trend",
      tdr_accounts: "tdr_trend",
      rd_accounts: "rd_trend",
      savings_accounts: "savings_trend",
      stamped_npa_accounts: "stamped_npa_accounts_trend",
      stamped_npa_balance: "stamped_npa_balance_trend",
      upgraded_npa_accounts: "upgraded_npa_accounts_trend",
      downgraded_npa_accounts: "downgraded_npa_accounts_trend",
      sma_total_accounts: "sma_total_accounts_trend",
      sma_total_balance: "sma_total_balance_trend",
    };

    const processed = filteredRows.map((row) => {
      const converted = { branch_code: row.branch_code };

      Object.keys(trendMap).forEach((key) => {
        converted[trendMap[key]] = row[trendMap[key]] ?? "SAME";
      });

      return converted;
    });

    res.json(processed);
  } catch (err) {
    console.error("Trend error:", err);
    res.status(500).json({
      message: "Failed to fetch trend",
      error: err.message,
    });
  }
});



// ===================================================================================
//                        ✅ GENERATE LOANS OPENED
// ===================================================================================

app.post("/generate-loans-opened-summary", async (req, res) => {
  try {

    const {
      branchCode,
      branchName,
      clusterName,
      districtName,
      page = 1,
      pageSize = 10,
      sortBy = null,
    } = req.body;

    // -----------------------------------------------------
    // NORMALIZE BRANCH CODE
    // -----------------------------------------------------
    const normalizeBranchCode = (code) => {
      if (!code) return "";
      return String(code)
        .replace(/,/g, "")
        .replace(/\.00$/, "")
        .replace(/^0+/, "")
        .trim();
    };

    // -----------------------------------------------------
    // 1) MASTER BRANCH LIST
    // -----------------------------------------------------
    let masterQuery = `
      SELECT 
        [col1] AS branch_code,
        [col2] AS branch_name,
        [col4] AS cluster_name,
        [col3] AS district_name
      FROM [dbo].[loanssanctioned]
      WHERE 1 = 1
    `;

    if (branchCode) {
      const bc = normalizeBranchCode(branchCode);
      masterQuery += `
        AND TRY_CAST(
              TRY_CAST(LTRIM(RTRIM([col1])) AS DECIMAL(18,2))
            AS INT)
            = TRY_CAST('${bc}' AS INT)
      `;
    }

    if (branchName) {
      masterQuery += ` AND LOWER(LTRIM(RTRIM([col2]))) LIKE '%' + LOWER('${branchName.trim()}') + '%'`;
    }

    if (clusterName) {
      masterQuery += ` AND LOWER(LTRIM(RTRIM([col4]))) LIKE '%' + LOWER('${clusterName.trim()}') + '%'`;
    }

    if (districtName) {
      masterQuery += ` AND LOWER(LTRIM(RTRIM([col3]))) LIKE '%' + LOWER('${districtName.trim()}') + '%'`;
    }

    const masterData = await queryUTIDatabase(masterQuery);

    if (!masterData || masterData.length === 0) {
      return res.json({
        success: true,
        message: "No branches found matching filters.",
        totalRecords: 0,
        totalPages: 0,
        data: [],
        loan_categories: [],
        column_structure: [],
      });
    }

    // -----------------------------------------------------
    // 2) FETCH RAW LOAN DATA
    // -----------------------------------------------------
    const loanQuery = `
      SELECT
        TRY_CAST(
          TRY_CAST([Branch Code] AS DECIMAL(18,2))
        AS INT) AS branch_code,
        [Loan Type] AS loan_type,
        TRY_CAST([Sanctioned Amount] AS decimal(18,2)) AS amount
      FROM [dbo].[LoansAccountsOpened]
      WHERE [Sanctioned Amount] IS NOT NULL
    `;

    const loanData = await queryUTIDatabase(loanQuery);

    // -----------------------------------------------------
    // 3) NORMALIZE LOAN TYPE CONSISTENTLY
    // -----------------------------------------------------
    const normalizeLoanType = (loanType = "") => {
      const t = String(loanType).toLowerCase();

      if (t.includes("gold") || t.includes("gl")) return "Gold Loans";
      if (t.includes("crop")) return "Crop Loans";
      if (t.includes("deposit")) return "Deposit Loans";
      if (t.includes("vehicle") || t.includes("vh")) return "Vehicle Loans";
      if (t.includes("staff")) return "Staff Loans";
      if (t.includes("housing") || t.includes("hl")) return "Housing Loans";
      if (t.includes("personal") || t.includes("pl")) return "Personal Loans";
      if (t.includes("cash credit") || t.includes("cc")) return "Cash Credit";
      if (t.includes("term loan") || t.includes("tl")) return "Term Loans";
      if (t.includes("od") || t.includes("overdraft")) return "Overdraft";

      return String(loanType).replace(/\s+/g, " ").trim();
    };

    const normalized = (loanData || []).map((r) => ({
      branch_code: normalizeBranchCode(r.branch_code),
      loan_type: normalizeLoanType(r.loan_type),
      amount: Number(r.amount || 0),
    }));

    // -----------------------------------------------------
    // 4) GET CATEGORY LIST
    // -----------------------------------------------------
    const loanCategories = Array.from(
      new Set(
        normalized.map((r) =>
          r.loan_type.replace(/\s+/g, " ").trim()
        )
      )
    ).sort();

    // -----------------------------------------------------
    // 5) GROUP DATA BY BRANCH + LOAN TYPE
    // -----------------------------------------------------
    const grouped = {};
    for (const row of normalized) {
      const cleanType = row.loan_type.replace(/\s+/g, " ").trim();
      const key = `${row.branch_code}||${cleanType}`;

      if (!grouped[key]) {
        grouped[key] = {
          branch_code: row.branch_code,
          loan_type: cleanType,
          accounts: 0,
          total_amount: 0,
        };
      }

      grouped[key].accounts += 1;
      grouped[key].total_amount += row.amount;
    }

    const summarized = Object.values(grouped);

    // -----------------------------------------------------
    // 6) MERGE MASTER DATA WITH SUMMARY
    // -----------------------------------------------------
    const mergedData = masterData.map((b) => {
      const branch_code = normalizeBranchCode(b.branch_code);

      const row = {
        "Branch Code": branch_code || "-",
        "Branch Name": b.branch_name || "-",
        Cluster: b.cluster_name || "-",
        District: b.district_name || "-",
      };

      let totalAc = 0;
      let totalAmt = 0;

      for (const cat of loanCategories) {
        const cleanCat = cat.replace(/\s+/g, " ").trim();

        const match = summarized.find(
          (s) => s.branch_code === branch_code && s.loan_type === cleanCat
        );

        const ac = match ? match.accounts : 0;
        const amt = match ? match.total_amount : 0;

        row[`${cleanCat}_Accounts`] = ac > 0 ? ac : "-";
        row[`${cleanCat}_Amount`] = amt > 0 ? amt.toFixed(2) : "-";

        totalAc += ac;
        totalAmt += amt;
      }

      row["Total_Accounts"] = totalAc > 0 ? totalAc : "-";
      row["Total_Amount"] = totalAmt > 0 ? totalAmt.toFixed(2) : "-";

      return row;
    });

    // -----------------------------------------------------
    // 7) SORTING
    // -----------------------------------------------------
    let sorted = [...mergedData];

    if (sortBy === "Total_Accounts") {
      sorted.sort((a, b) => Number(b.Total_Accounts || 0) - Number(a.Total_Accounts || 0));
    } else if (sortBy === "Total_Amount") {
      sorted.sort((a, b) => Number(b.Total_Amount || 0) - Number(a.Total_Amount || 0));
    } else {
      sorted.sort((a, b) => Number(a["Branch Code"]) - Number(b["Branch Code"]));
    }

    // -----------------------------------------------------
    // 8) PAGINATION
    // -----------------------------------------------------
    const totalRecords = sorted.length;
    const totalPages = Math.max(1, Math.ceil(totalRecords / pageSize));
    const pagedData = sorted.slice((page - 1) * pageSize, page * pageSize);

    // -----------------------------------------------------
    // 9) COLUMN STRUCTURE
    // -----------------------------------------------------
    const column_structure = [
      ...loanCategories.map((cat) => ({
        parent: cat.replace(/\s+/g, " ").trim(),
        children: ["A/Cs", "Amount"],
      })),
      {
        parent: "Total",
        children: ["A/Cs", "Amount"],
      },
    ];

    return res.json({
      success: true,
      totalRecords,
      totalPages,
      loan_categories: [...loanCategories, "Total"],
      column_structure,
      data: pagedData,
    });

  } catch (err) {
    console.error("❌ Error in /generate-loans-opened-summary:", err);
    return res.status(500).json({
      success: false,
      message: "Server error generating summary",
      error: err.message,
    });
  }
});

//============================================================================================
//                                      BRANCH CONTACTS
//=============================================================================================
app.post("/get-branch-contacts", async (req, res) => {
  try {
    let { branchCodes } = req.body;

    // --- Ensure branchCodes exists and is an array
    if (!branchCodes) {
      console.warn("⚠️ branchCodes missing in request body");
      return res.json({ data: [] });
    }

    if (!Array.isArray(branchCodes)) branchCodes = [branchCodes];

    // --- Convert to numbers and ignore invalid codes
    branchCodes = branchCodes
      .map((c) => {
        if (!c && c !== 0) return null;

        // If object, try to extract possible keys
        if (typeof c === "object") {
          const keys = ["Branch Code", "BR Code", "Br Code", "branch_code", "Brcode"];
          for (const key of keys) {
            if (c[key] !== undefined && c[key] !== null && c[key] !== "") {
              const n = parseInt(c[key], 10);
              return isNaN(n) ? null : n;
            }
          }
          return null;
        }

        // string or number
        const n = parseInt(c, 10);
        return isNaN(n) ? null : n;
      })
      .filter(Boolean); // remove invalid / NaN codes

    if (!branchCodes.length) {
      console.warn("⚠️ No valid branch codes after normalization");
      return res.json({ data: [] });
    }

    // --- Prepare query placeholders
    const placeholders = branchCodes.map(() => "?").join(",");

    const query = `
      SELECT
        [Br Code] AS branch_code,
        [Cluster],
        [Branch Name] AS branch_name,
        [Branch Manager] AS manager,
        [Contact Number] AS contact_number
      FROM [dbo].[branchContacts]
      WHERE TRY_CAST([Br Code] AS INT) IN (${placeholders})
      ORDER BY [Cluster], [Branch Name]
    `;

    let results = [];
    try {
      results = await queryUTIDatabase(query, branchCodes);
    } catch (dbErr) {
      console.error("❌ Database query error:", dbErr.stack || dbErr);
      return res.status(500).json({ error: "Database query failed" });
    }

    // 🚫 DO NOT format these results with formatNumbersDeep()
    // because it will add commas/decimals to phone numbers.
    // Just send raw DB output as-is.
    return res.json({ data: results });
  } catch (err) {
    console.error("❌ Unexpected error in /get-branch-contacts:", err.stack || err);
    return res.status(500).json({ error: "Internal server error" });
  }
});


// ============================================================================================
//                                 UPLOAD MIS (DYNAMIC + SMART SORT CHECK)
// ============================================================================================

// ============================================================================================
//                                PART A — CONSTANTS + MEANINGS
// ============================================================================================

// SMART NORMALIZER
function normalizeSortableHeader(text) {
  if (!text) return "";
  return text
    .toLowerCase()
    .replace(/\b(19|20)\d{2}\b/g, "")
    .replace(/\b\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}\b/g, "")
    .replace(
      /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december)\b/g,
      ""
    )
    .replace(/\bfy\b/g, "")
    .replace(/\bfrom\b/g, "")
    .replace(/\bto\b/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

// SECTION → TABLE
const sectionTableMap = {
  deposits: "Deposits",
  "accounts opened": "deposits_accounts_opened",
  advances: "Advances",
  npa: "NPA",
  loans: "loanssanctioned",
  sma: "SMA",
  "loans accounts opened": "LoansAccountsOpened",
  "daily summary": "DailySummary",
};

// CLUSTER SECTION → TABLE
const clusterTableMap = {
  deposits: "DCS",
  advances: "ACS",
  "accounts opened": "DAOCS",
  npa: "NPACS",
  loans: "LSCS",
};

// BASIC AUTO-MEANING
function inferMeaning(label) {
  if (!label || typeof label !== "string") return null;

  const normalized = label.replace(/\u00A0/g, " ").toLowerCase().trim();

  if (normalized.includes("branch code")) return "branch_code";
  if (normalized.includes("branch name")) return "branch_name";
  if (normalized.includes("district")) return "district";
  if (normalized.includes("cluster")) return "cluster";

  if (normalized === "as on date") return "as_on_date";
  return null;
}

// ============================================================================================
//        STRICT MEANINGS — MATCH EXACTLY WHAT THE STORED PROCEDURE EXPECTS
// ============================================================================================

// -------- SMA ---------
const SMA_MEANING_OVERRIDES = {
  "SMA 0 A/Cs": "sma0_accounts",
  "SMA 0 Amount": "sma0_amount",
  "SMA 1 A/Cs": "sma1_accounts",
  "SMA 1 Amount": "sma1_amount",
  "SMA 2 A/Cs": "sma2_accounts",
  "SMA 2 Amount": "sma2_amount",
  "Total_ac": "total_accounts",     // required by SP
  "Total_balance": "total_balance", // required by SP
};

// -------- DEPOSITS ACCOUNTS OPENED --------
const DAO_MEANING_OVERRIDES = {
  "TDR (>=1LAKH ) A/Cs": "tdr_accounts",
  "RD A/Cs": "rd_accounts",
  "Savings A/Cs": "savings_accounts",
  "TDR (>=1LAKH ) Amount": "tdr_amount",
  "RD Amount": "rd_amount",
  "Savings Amount": "savings_amount",
};
const DAOCS_MEANING_OVERRIDES = {
  "TDR(>=1 LAKH) A/Cs": "tdr_accounts",
  "TDR(>=1 LAKH) Amount": "tdr_amount",
  "RD A/Cs": "rd_accounts",
  "RD Amount": "rd_amount",
  "Savings A/Cs": "savings_accounts",
  "Savings Amount": "savings_amount",
  Vouchers: "vouchers"
};

// -------- NPACS (Cluster NPA Summary) --------
const NPACS_MEANING_OVERRIDES = {
  // --- As On Date (today) ---
  "stamped npa as on date a/cs": "stamped_npa_accounts",
  "stamped npa as on date balance": "stamped_npa_balance",

  "unstamped pnpa as on date a/cs": "unstamped_npa_accounts",
  "unstamped pnpa as on date balance": "unstamped_npa_balance",

  "upgraded as on date a/cs": "upgraded_npa_accounts",
  "upgraded as on date amount": "upgraded_npa_amount",

  "downgraded as on date a/cs": "downgraded_npa_accounts",
  "downgraded as on date amount": "downgraded_npa_amount",
};

// -------- FINAL MEANINGS REQUIRED BY update_daily_summary_from_history --------
const DAILY_SUMMARY_REQUIRED = {
  Deposits: { "As on Date": "as_on_date_deposits" },
  Advances: { "As on Date": "as_on_date_advances" },

  deposits_accounts_opened: {
  "TDR (>=1LAKH ) A/Cs": "tdr_accounts",
  "RD A/Cs": "rd_accounts",
  "Savings A/Cs": "savings_accounts"},
 
  NPA: {
    "STAMPED NPA As on Date A/Cs": "stamped_npa_accounts",
    "STAMPED NPA As on Date Balance": "stamped_npa_balance",
    "UPGRADED As on Date A/Cs": "upgraded_npa_accounts",
    "DOWNGRADED As on Date A/Cs": "downgraded_npa_accounts",
  },
};



// ============================================================================================
//       SORTABLE OVERRIDES — NO CHANGE NEEDED HERE, BUT CLEANED AND STRICT
// ============================================================================================

const SORTABLE_MEANING_OVERRIDES = {
  Deposits: {
    "Change over PD": "change_over_pd_deposits",
    GDM: "gdm_deposits",
    GUM: "gum_deposits",
    "Budget Achieved%": "budget_achieved_deposits",
  },

  Advances: {
    "Change over PD": "change_over_pd_advances",
    GDM: "gdm_advances",
    GUM: "gum_advances",
    "Budget Achieved %": "budget_achieved_advances",
  },

  deposits_accounts_opened: {
  "TDR (>=1LAKH ) A/Cs": "tdr_accounts",
  "RD A/Cs": "rd_accounts",
  "Savings A/Cs": "savings_accounts"
  },

  NPA: {
    "STAMPED NPA As on Date Balance": "stamped_npa_balance",
    "STAMPED NPA As on Date A/Cs": "stamped_npa_accounts",
  },

SMA: {
    "SMA 0 A/Cs": "sma0_accounts",
    "SMA 1 A/Cs": "sma1_accounts",
    "SMA 2 A/Cs": "sma2_accounts",
  },
  loanssanctioned: {
    "OVER ALL FY from April-2025 A/Cs": "over_all_fy_accounts",
    "OVER ALL FY from April-2025 Amount": "over_all_fy_amount",
  },

  LoansAccountsOpened: {
    Total_Accounts: "total_accounts",
    Total_Amount: "total_amount",
  },

  DailySummary: {
    total_deposits: "total_deposits",
    tdr_accounts: "tdr_accounts",
    rd_accounts: "rd_accounts",
    savings_accounts: "savings_accounts",
    total_advances: "total_advances",
    stamped_npa_accounts: "stamped_npa_accounts",
    stamped_npa_balance: "stamped_npa_balance",
    upgraded_npa_accounts: "upgraded_npa_accounts",
    downgraded_npa_accounts: "downgraded_npa_accounts",
    sma_total_accounts: "sma_total_accounts",
    sma_total_balance: "sma_total_balance",
  },
};

// ============================================================================================
// TYPE DETECTOR
// ============================================================================================
function detectDataType(values) {
  const sample = values.find((v) => v !== null && v !== undefined && v !== "");
  if (!sample) return "string";

  const num = Number(sample);
  if (!isNaN(num)) return sample.toString().includes(".") ? "decimal" : "int";
  if (!isNaN(Date.parse(sample))) return "date";

  return "string";
}
// ============================================================================================
//                  PART B — FINAL MEANING RESOLUTION ENGINE (CRITICAL)
// ============================================================================================

function resolveFinalMeaning(label, tableName) {
  const trimmed = label.trim();

  // -------------------------------------------------------------------------
  // 1️⃣ STRICT — meanings required by update_daily_summary_from_history
  // -------------------------------------------------------------------------
  const sectionOverride = DAILY_SUMMARY_REQUIRED[tableName];
  if (sectionOverride && sectionOverride[trimmed]) {
    return sectionOverride[trimmed];
  }

  // -------------------------------------------------------------------------
  // 2️⃣ STRICT SMA overrides (SP depends on these)
  // -------------------------------------------------------------------------
  if (tableName === "SMA" && SMA_MEANING_OVERRIDES[trimmed]) {
    return SMA_MEANING_OVERRIDES[trimmed];
  }

 // -------------------------------------------------------------------------
// 3️⃣ STRICT Deposits Accounts Opened & DAOCS mappings
// -------------------------------------------------------------------------
if (
  (tableName === "deposits_accounts_opened" || tableName === "DAOCS") &&
  (DAO_MEANING_OVERRIDES[trimmed] || DAOCS_MEANING_OVERRIDES[trimmed])
) {
  // prefer DAOCS overrides when tableName === 'DAOCS'
  if (tableName === "DAOCS" && DAOCS_MEANING_OVERRIDES[trimmed]) {
    return DAOCS_MEANING_OVERRIDES[trimmed];
  }
  // otherwise fallback to DAO mapping
  if (DAO_MEANING_OVERRIDES[trimmed]) return DAO_MEANING_OVERRIDES[trimmed];
}
// -------------------------------------------------------------------------
// ⭐ LSCS — Loans Sanctioned Cluster Summary (dynamic date headers)
// -------------------------------------------------------------------------
if (tableName === "LSCS") {
  const lower = trimmed.toLowerCase().replace(/\s+/g, " ").trim();

  // fix: detect dates even when prefixed with underscore
  const hasDate = /_?\d{1,2}[-\/\.]\d{1,2}[-\/\.]\d{2,4}/.test(lower);

  const hasMonthYear =
    /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*[-\s]?\d{2,4}\b/
      .test(lower);

  // 1️⃣ On <DATE> A/Cs or Amount
  if ((lower.startsWith("on") || lower.startsWith("on _")) &&
      (hasDate || hasMonthYear)) {

    if (lower.includes("a/c")) return "on_date_accounts";
    if (lower.includes("amount")) return "on_date_amount";
  }

  // 2️⃣ During <Month-Year> A/Cs or Amount
  if (lower.startsWith("during") && hasMonthYear) {
    if (lower.includes("a/c")) return "during_month_accounts";
    if (lower.includes("amount")) return "during_month_amount";
  }

  // 3️⃣ OVER ALL FY from APRIL-yyyy
  if (lower.includes("over all fy")) {
    if (lower.includes("a/c")) return "fy_overall_accounts";
    if (lower.includes("amount")) return "fy_overall_amount";
  }
}



// -------------------------------------------------------------------------
// ⭐ NPACS — DATE-BASED COLUMNS + AS ON DATE COLUMNS
// -------------------------------------------------------------------------
if (tableName === "NPACS") {
  const lower = trimmed.toLowerCase().replace(/\s+/g, " ").trim();

  // detect any date like 31/03/2025 or 20-11-2025 or 31 03 2025
  const hasDate = /\b\d{1,2}[-\/\s\.]\d{1,2}[-\/\s\.]\d{2,4}\b/.test(lower);

  // -------------------------------
  // ⭐ 1. PREVIOUS DATE COLUMNS
  // -------------------------------
  if (hasDate) {
    if (lower.includes("stamped npa") && lower.includes("a/c")) {
      return "stamped_npa_prev_accounts";
    }
    if (lower.includes("stamped npa") && lower.includes("balance")) {
      return "stamped_npa_prev_balance";
    }
    if (lower.includes("unstamped") && lower.includes("a/c")) {
      return "unstamped_npa_prev_accounts";
    }
    if (lower.includes("unstamped") && lower.includes("balance")) {
      return "unstamped_npa_prev_balance";
    }
  }

  // -------------------------------
  // ⭐ 2. AS ON DATE COLUMNS
  // -------------------------------
  if (lower.includes("as on date")) {
    if (lower.includes("stamped npa") && lower.includes("a/c")) {
      return "stamped_npa_accounts";
    }
    if (lower.includes("stamped npa") && lower.includes("balance")) {
      return "stamped_npa_balance";
    }
    if (lower.includes("unstamped pnpa") && lower.includes("a/c")) {
      return "unstamped_npa_accounts";
    }
    if (lower.includes("unstamped pnpa") && lower.includes("balance")) {
      return "unstamped_npa_balance";
    }
    if (lower.includes("upgraded") && lower.includes("a/c")) {
      return "upgraded_npa_accounts";
    }
    if (lower.includes("upgraded") && lower.includes("amount")) {
      return "upgraded_npa_amount";
    }
    if (lower.includes("downgraded") && lower.includes("a/c")) {
      return "downgraded_npa_accounts";
    }
    if (lower.includes("downgraded") && lower.includes("amount")) {
      return "downgraded_npa_amount";
    }
  }
}


  // -------------------------------------------------------------------------
  // 4️⃣ Sortable overrides (numeric columns used in sorting)
  // -------------------------------------------------------------------------
  const sortableOverride = SORTABLE_MEANING_OVERRIDES[tableName] || {};
  if (sortableOverride[trimmed]) {
    return sortableOverride[trimmed];
  }

  // -------------------------------------------------------------------------
  // 5️⃣ Auto meaning detection (branch_code, branch_name, etc.)
  // -------------------------------------------------------------------------
  const inferred = inferMeaning(trimmed);
  if (inferred) return inferred;

  // -------------------------------------------------------------------------
  // 6️⃣ Otherwise → no meaning
  // -------------------------------------------------------------------------
  return null;
}
// ============================================================================================
// PART C — MAIN UPLOAD ROUTE (ALIGNED WITH PROCEDURE + MEANING ENGINE)
// ============================================================================================

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, "uploads/"),
  filename: (req, file, cb) =>
    cb(null, `${Date.now()}${path.extname(file.originalname)}`),
});

const upload = multer({ storage });

// ============================================================================================
// MAIN UPLOAD ROUTE
// ============================================================================================
app.post("/upload-mis", upload.single("file"), async (req, res) => {
  const { section, clusterSection, userId, role } = req.body;

  if (!req.file)
    return res.status(400).json({ message: "No file uploaded" });

  let tableName;

if (section) {
  const normalized = section.toLowerCase().trim();
  tableName = sectionTableMap[normalized];
}

else if (clusterSection) {
  const normalizedCluster = clusterSection.toLowerCase().trim();
  tableName = clusterTableMap[normalizedCluster];
}


  if (!tableName)
    return res.status(400).json({ message: "Invalid section" });

  const historyTable = `${tableName}_history`;

  // ==========================================================================================
  // SPECIAL CASE: LoansAccountsOpened (real columns, NOT col1..col30)
  // ==========================================================================================
  if (tableName === "LoansAccountsOpened") {
    try {
      const ext = req.file.originalname.toLowerCase();
      const fileContent = fs.readFileSync(req.file.path);
      const workbook = XLSX.read(
        ext.endsWith(".csv") ? fileContent.toString("utf8") : fileContent,
        { type: ext.endsWith(".csv") ? "string" : "buffer" }
      );

      let rows = XLSX.utils.sheet_to_json(
        workbook.Sheets[workbook.SheetNames[0]]
      );

      if (!rows.length)
        return res.status(400).json({ message: "File empty" });

      const fileCols = Object.keys(rows[0]);

      // Clear table
      await queryUTIDatabase(`DELETE FROM [${tableName}]`);

      // Insert raw columns
      for (const row of rows) {
        const cols = fileCols.map((c) => `[${c}]`).join(", ");
        const placeholders = fileCols.map(() => "?").join(", ");
        const values = fileCols.map((c) => row[c] ?? null);

        await queryUTIDatabase(
          `INSERT INTO [${tableName}] (${cols}) VALUES (${placeholders})`,
          values
        );
      }

      // Clear old sortable for this section
      await queryUTIDatabase(
        `DELETE FROM MIS.dbo.MIS_Sortable_Columns WHERE section = ?`,
        [tableName]
      );

      // Rebuild MIS_Sortable_Columns based on overrides
      const sortableOverrides = SORTABLE_MEANING_OVERRIDES[tableName] || {};
      const sortablePromises = [];

      fileCols.forEach((c, idx) => {
        const label = c.trim();
        const meaning = sortableOverrides[label];
        if (meaning) {
          sortablePromises.push(
            queryUTIDatabase(
              `
              INSERT INTO MIS.dbo.MIS_Sortable_Columns
              (section, display_label, meaning)
              VALUES (?, ?, ?)
            `,
              [tableName, label, meaning]
            )
          );
        }
      });

      await Promise.all(sortablePromises);

      await logActivity(
        userId,
        role,
        "Upload MIS Report",
        `Uploaded ${rows.length} rows → LoansAccountsOpened`
      );

      return res.json({
        success: true,
        message:
          "LoansAccountsOpened uploaded with real columns + sortable meanings",
      });
    } catch (err) {
      console.error("❌ Upload LoansAccountsOpened failed:", err);
      return res.status(500).json({
        message: "Error uploading LoansAccountsOpened",
        error: err.message,
      });
    }
  }

  // ==========================================================================================
  // DEFAULT MODE: col1..col30 (Deposits, Advances, NPA, SMA, deposits_accounts_opened, etc.)
  // ==========================================================================================
  try {
    const ext = req.file.originalname.toLowerCase();
    const fileContent = fs.readFileSync(req.file.path);
    const workbook = XLSX.read(
      ext.endsWith(".csv") ? fileContent.toString("utf8") : fileContent,
      { type: ext.endsWith(".csv") ? "string" : "buffer" }
    );

    let rows = XLSX.utils.sheet_to_json(
      workbook.Sheets[workbook.SheetNames[0]]
    );
    if (!rows.length)
      return res.status(400).json({ message: "File empty" });

    const fileCols = Object.keys(rows[0]);

    if (fileCols.length > 30)
      return res.status(400).json({ message: "Max 30 columns allowed" });

    // Remove SMA GRAND TOTAL if present
    if (tableName === "SMA") {
      rows = rows.filter((r) => {
        const firstCol = Object.keys(r)[0];
        const v = (r[firstCol] || "").toString().trim().toUpperCase();
        return v !== "GRAND TOTAL";
      });
    }

    // Clear old mappings and sortable
    await queryUTIDatabase(
      `DELETE FROM MIS_Column_Mapping WHERE section = ?`,
      [tableName]
    );
    await queryUTIDatabase(
      `DELETE FROM MIS_Column_Mapping WHERE section = ?`,
      [historyTable]
    );
    await queryUTIDatabase(
      `DELETE FROM MIS.dbo.MIS_Sortable_Columns WHERE section = ?`,
      [tableName]
    );

    let colIndex = 1;
    const mappingPromises = [];
    const sortablePromises = [];

    // Build MIS_Column_Mapping + MIS_Sortable_Columns
    for (const col of fileCols) {
      const label = col.trim();
      const detectedType = detectDataType(rows.map((r) => r[col]));

      // 🧠 FINAL MEANING (100% aligned with stored procedure)
      const finalMeaning = resolveFinalMeaning(label, tableName);

      // 1️⃣ Insert into MIS_Column_Mapping for main table
      mappingPromises.push(
        queryUTIDatabase(
          `
          INSERT INTO MIS_Column_Mapping
          (section, col_number, display_label, detected_type, meaning)
          VALUES (?, ?, ?, ?, ?)
        `,
          [tableName, colIndex, label, detectedType, finalMeaning]
        )
      );

      // 2️⃣ Insert into MIS_Column_Mapping for *_history table
      mappingPromises.push(
        queryUTIDatabase(
          `
          INSERT INTO MIS_Column_Mapping
          (section, col_number, display_label, detected_type, meaning)
          VALUES (?, ?, ?, ?, ?)
        `,
          [historyTable, colIndex, label, detectedType, finalMeaning]
        )
      );

      // 3️⃣ Insert into MIS_Sortable_Columns **ONLY IF** this label
      //    is in SORTABLE_MEANING_OVERRIDES (so grid sort works)
      const sortableOverride = SORTABLE_MEANING_OVERRIDES[tableName] || {};
      if (sortableOverride[label]) {
        sortablePromises.push(
          queryUTIDatabase(
            `
            INSERT INTO MIS.dbo.MIS_Sortable_Columns
            (section, display_label, meaning)
            VALUES (?, ?, ?)
          `,
            [tableName, label, sortableOverride[label]]
          )
        );
      }

      colIndex++;
    }

    await Promise.all([...mappingPromises, ...sortablePromises]);

    // Prepare rows in col1..col30 order
    const dbRows = rows.map((r) => {
      const ordered = fileCols.map((c) => {
        const val = r[c];
        if (val === "" || val === "-") return null;
        return val;
      });
      while (ordered.length < 30) ordered.push(null);
      return ordered;
    });

    // Clear base tables
    await queryUTIDatabase(`DELETE FROM [${tableName}]`);

    const colNames = [...Array(30).keys()].map((i) => `col${i + 1}`);
    const placeholders = colNames.map(() => "?").join(", ");

    // Insert rows into base + history
    for (const row of dbRows) {
      await queryUTIDatabase(
        `INSERT INTO [${tableName}] (${colNames.join(",")}) VALUES (${placeholders})`,
        row
      );
      await queryUTIDatabase(
        `INSERT INTO [${historyTable}] (${colNames.join(",")}) VALUES (${placeholders})`,
        row
      );
    }

    await logActivity(
      userId,
      role,
      "Upload MIS Report",
      `Uploaded ${dbRows.length} rows into ${tableName}`
    );

    return res.json({
      success: true,
      message: "MIS uploaded successfully",
      uploaded: dbRows.length,
    });
  } catch (err) {
    console.error("❌ Upload MIS failed:", err);
    return res.status(500).json({
      message: "Server error",
      error: err.message,
    });
  }
});

// ============================================================================================
// PART D — UNIVERSAL SORTABLE COLUMNS ENDPOINT
// ============================================================================================
app.post("/get-sortable-columns", async (req, res) => {
  try {
    const { section } = req.body;

    if (!section) {
      return res.status(400).json({
        success: false,
        message: "Section required",
      });
    }

    // Example: "deposits" → "Deposits"
    const normalized = section.toLowerCase().trim();
const tableName = sectionTableMap[normalized] || section;

    // Fetch allowed sortable columns from mapping table
    const rows = await queryUTIDatabase(
      `
      SELECT display_label, meaning
      FROM MIS.dbo.MIS_Sortable_Columns
      WHERE section = ?
      ORDER BY display_label ASC
      `,
      [tableName]
    );

    // Format into UI-friendly objects
    const columns = (rows || [])
      .filter((r) => r.meaning && r.display_label)
      .map((r) => ({
        label: r.display_label, // what user sees in dropdown
        meaning: r.meaning,     // used internally as sortBy value
      }));

    return res.json({
      success: true,
      section: tableName,
      columns,
    });

  } catch (err) {
    console.error("❌ get-sortable-columns failed:", err);
    return res.status(500).json({
      success: false,
      message: "Server error",
      error: err.message,
    });
  }
});



 //=================================================
 //           Business date
 //=================================================

app.post("/set-as-on-date", async (req, res) => {
  try {
    const { as_on_date } = req.body;

    if (!as_on_date) {
      return res.status(400).json({
        success: false,
        message: "as_on_date is required",
      });
    }

    // ⭐ Update the FIRST row (identity-safe)
    await queryUTIDatabase(
      `
        UPDATE MIS.dbo.MIS_date
        SET as_on_date = ?, updated_at = GETDATE()
        WHERE id = (SELECT TOP 1 id FROM MIS.dbo.MIS_date ORDER BY id ASC)
      `,
      [as_on_date]
    );

    return res.json({
      success: true,
      message: "As-On Date updated successfully",
      as_on_date,
    });

  } catch (err) {
    console.error("❌ Error (set-as-on-date):", err);
    return res.status(500).json({
      success: false,
      message: "Server error updating As-On Date",
    });
  }
});


app.post("/get-as-on-date", async (req, res) => {
  try {
    const rows = await queryUTIDatabase(
      `
        SELECT as_on_date
        FROM MIS.dbo.MIS_date
        WHERE id = 1
      `
    );

    if (!rows || rows.length === 0) {
      return res.status(404).json({
        success: false,
        message: "As-On Date not found",
      });
    }

    return res.json({
      success: true,
      as_on_date: rows[0].as_on_date,
    });

  } catch (err) {
    console.error("❌ Error (get-as-on-date):", err);
    return res.status(500).json({
      success: false,
      message: "Server error fetching As-On Date",
    });
  }
});

// =====================================================================================
//                ✅ Fetch Activity Logs (with filters & pagination)
// ======================================================================================
app.post("/get-activity-logs", async (req, res) => {
  try {
    const { userId, fromDate, toDate, role, action, page = 1, limit = 10 } = req.body;

    let sql = `
      SELECT a.*, e.[Emp No.] AS emp_no, e.[Employee Name] AS emp_name
      FROM [dbo].[activity_logs] a
      LEFT JOIN [dbo].[employees] e 
        ON a.user_id = e.[Emp No.]  -- 🔥 FIX HERE
      WHERE 1=1
    `;
    const params = [];

    if (userId) {
      sql += " AND a.user_id = ?";
      params.push(userId);
    }
    if (role) {
      sql += " AND a.role = ?";
      params.push(role);
    }
    if (action) {
      sql += " AND a.action LIKE ?";
      params.push(`%${action}%`);
    }
    if (fromDate && toDate) {
      sql += " AND CAST(a.created_at AS DATE) BETWEEN ? AND ?";
      params.push(fromDate, toDate);
    }

    const countSql = `SELECT COUNT(*) AS total FROM (${sql}) AS subquery`;

    const offset = (page - 1) * limit;
    sql += " ORDER BY a.created_at DESC OFFSET ? ROWS FETCH NEXT ? ROWS ONLY";
    params.push(offset, parseInt(limit));

    const countResult = await queryUTIDatabase(countSql, params.slice(0, -2));
    const total = countResult[0]?.total || 0;
    const totalPages = Math.ceil(total / limit);

    const logs = await queryUTIDatabase(sql, params);

    return res.json({ logs, page: parseInt(page), totalPages, total });
  } catch (err) {
    console.error("❌ Error in /get-activity-logs:", err);
    return res.status(500).json({ message: "Server error while fetching activity logs" });
  }
});

// ======================================================================================
//                          ✅ Activity Stats (Date Filter)
// ======================================================================================
app.post("/get-activity-stats", async (req, res) => {
  try {
    const { fromDate, toDate } = req.body;

    let whereClause = "";
    const params = [];

    if (fromDate && toDate) {
      whereClause = "WHERE CAST(created_at AS DATE) BETWEEN ? AND ?";
      params.push(fromDate, toDate);
    } else if (fromDate) {
      whereClause = "WHERE CAST(created_at AS DATE) >= ?";
      params.push(fromDate);
    } else if (toDate) {
      whereClause = "WHERE CAST(created_at AS DATE) <= ?";
      params.push(toDate);
    } else {
      whereClause = "WHERE CAST(created_at AS DATE) = CAST(GETDATE() AS DATE)";
    }

    const sql = `
      SELECT
        SUM(CASE WHEN action = 'Login' THEN 1 ELSE 0 END) AS totalLogins,
        SUM(CASE WHEN action = 'Login Failed' THEN 1 ELSE 0 END) AS failedLogins,
        SUM(CASE WHEN action = 'Upload MIS Report' THEN 1 ELSE 0 END) AS misUploads,
        SUM(CASE WHEN action = 'Password Reset' THEN 1 ELSE 0 END) AS passwordResets,
        SUM(CASE WHEN action = 'Account Locked' THEN 1 ELSE 0 END) AS lockedAccounts
      FROM [dbo].[activity_logs]
      ${whereClause};
    `;

    const rows = await queryUTIDatabase(sql, params);
    const stats = rows[0] || {};

    res.json({
      totalLogins: stats.totalLogins || 0,
      failedLogins: stats.failedLogins || 0,
      misUploads: stats.misUploads || 0,
      passwordResets: stats.passwordResets || 0,
      lockedAccounts: stats.lockedAccounts || 0,
    });
  } catch (err) {
    console.error("❌ Error fetching activity stats:", err);
    res.status(500).json({ message: "Server error while fetching activity stats" });
  }
});

// ====================================================================================
//                         ✅ Fetch Last Activities per User
// ====================================================================================
app.post("/get-last-activities", async (req, res) => {
  try {
    const { role } = req.body;

    let sql = `
      SELECT a1.*, e.[Emp No.] AS emp_no, e.[Employee Name] AS emp_name
      FROM [dbo].[activity_logs] a1
      INNER JOIN (
        SELECT user_id, MAX(created_at) AS max_date
        FROM [dbo].[activity_logs]
        GROUP BY user_id
      ) a2 ON a1.user_id = a2.user_id AND a1.created_at = a2.max_date
      JOIN [dbo].[employees] e ON a1.user_id = e.[Emp No.]
      WHERE 1=1
    `;

    const params = [];
    if (role) {
      sql += " AND a1.role = ?";
      params.push(role);
    }

    sql += " ORDER BY a1.created_at DESC";

    const results = await queryUTIDatabase(sql, params);
    res.json({ lastActivities: results });
  } catch (err) {
    console.error("❌ Error fetching last activities:", err);
    res.status(500).json({ message: "Server error while fetching last activities" });
  }
});


//============================================================================================
//                                          SERVER CONNECTION
//=============================================================================================

 //app.listen(PORT, () => {
  //console.log(`🚀 Backend server running on ${BASE_URL}`);
 //});
https.createServer(httpsOptions, app).listen(5000, () => {
  console.log("HTTPS Server running on https://mobile.coastal.bank.in:5000");
});


