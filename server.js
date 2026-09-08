require('dotenv').config(); 
const jwt = require('jsonwebtoken');
const authMiddleware = require('./authMiddleware');
const dashboardAuthMiddleware =
  require("./dashboardAuthMiddleware");
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
const { fileTypeFromBuffer } = require("file-type");
const rateLimit = require("express-rate-limit");


const app = express();
const PORT = process.env.PORT ;

const httpsOptions = {
  key: fs.readFileSync(process.env.SSL_KEY),
  cert: fs.readFileSync(process.env.SSL_CERT),
  ca: fs.readFileSync(process.env.SSL_CA)
};

// LAN IP
const LAN_IP = "40.80.79.26";

// BASE URL (now HTTPS)
const BASE_URL = process.env.BASE_URL || `https://mobile.coastal.bank.in:5000`;
app.set("BASE_URL", BASE_URL);

// Middleware
app.use(cors());
app.use(bodyParser.json());
;
app.use((req, res, next) => {

  if (
    req.url !== "/dashboard-session-check"
  ) {
    console.log(
      "🔥 API HIT:",
      req.url
    );
  }

  next();

});

const writeDailyLog = require("./logger");

function formatMessage(args) {
  return args.map(a =>
    typeof a === "object"
      ? JSON.stringify(a)
      : a
  ).join(" ");
}

const originalLog = console.log;
const originalWarn = console.warn;
const originalError = console.error;

console.log = (...args) => {

  writeDailyLog(
    "backend",
    formatMessage(args)
  );

  originalLog(...args);
};

console.warn = (...args) => {

  writeDailyLog(
    "backend",
    formatMessage(args)
  );

  originalWarn(...args);
};

console.error = (...args) => {

  writeDailyLog(
    "backend",
    formatMessage(args)
  );

  originalError(...args);
};
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
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  server: process.env.DB_SERVER,
  database: process.env.DB_NAME,
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

// 🔥 UNIVERSAL SAFE STRING NORMALIZER
const norm = (val) => {
  if (val === null || val === undefined) return "";
  return String(val).trim();
};

// ⭐ UNIVERSAL FILTER NORMALIZER (USE IN ALL APIs)
const safeTrim = (val) => {
  if (val === undefined || val === null) return "";
  return String(val).trim();
};

const parseFilters = (filters = {}) => {
  return {
    branchCode: safeTrim(filters.branchCode ?? filters.branch_code),
    branchName: safeTrim(filters.branchName ?? filters.branch_name),
    districtName: safeTrim(filters.districtName ?? filters.district),
    clusterName: safeTrim(filters.clusterName ?? filters.cluster),
  };
};

app.get("/backend", (req, res) => {
  res.send("Backend is working");
});

// =====================================================
// RATE LIMITERS
// =====================================================

// Login protection
const loginLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 10,
  standardHeaders: true,
  legacyHeaders: false,

  message: {
    success: false,
    message: "Too many login attempts. Try again after 15 minutes.",
  },
});

// Upload protection
const uploadLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 20,
  standardHeaders: true,
  legacyHeaders: false,

  message: {
    success: false,
    message: "Too many uploads. Try again later.",
  },
});

// Password reset protection
const passwordLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 10,
  standardHeaders: true,
  legacyHeaders: false,

  message: {
    success: false,
    message: "Too many password reset attempts.",
  },
});

//========================= MSSQL MISUAT Connection =========================//
const mssqlConfigMISUAT = {
  user: process.env.MISUAT_USER,
  password: process.env.MISUAT_PASSWORD,
  server: process.env.MISUAT_SERVER,
  database: process.env.MISUAT_DB,
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
//==========================================================
//    level wise rstriction for At A Glance 
//==========================================================
function applyLevel1Restriction({
  userLevel,
  requestedLevel,
  branch_code,
  branch_name,
  cluster_name,
  designation,
}) {

  // LEVEL 1
  if (userLevel === "Level 1") {
    return {
      level: "BRANCH",
      branch_code,
      branch_name: null,
      cluster_name: null,
    };
  }

// LEVEL 2
if (userLevel === "Level 2") {

  let fixedCluster = cluster_name;

  const match =
    designation?.match(
      /Cluster\s*Head\s*[-:]\s*(.*)/i
    );

  if (match?.[1]) {
    fixedCluster = match[1].trim();
  }

  // Branch selected
  if (branch_code || branch_name) {
    return {
      level: "BRANCH",
      branch_code,
      branch_name,
      cluster_name: null,
    };
  }

  // Cluster view
  return {
    level: "CLUSTER",
    branch_code: null,
    branch_name: null,
    cluster_name: fixedCluster,
  };
}
  // LEVEL 3
  return {
    level: requestedLevel,
    branch_code,
    branch_name,
    cluster_name:
      requestedLevel === "CLUSTER"
        ? cluster_name
        : null,
  };
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

//========================Enumeration Monitoring====================
async function logSecurityEvent(userId, action, details) {
  try {
    await queryUTIDatabase(
      `
      INSERT INTO activity_logs
      (user_id, role, action, details, created_at)
      VALUES (?, ?, ?, ?, GETDATE())
      `,
      [
        userId || "UNKNOWN",
        "Security",
        action,
        details,
      ]
    );
  } catch (err) {
    console.error("Security log failed:", err);
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
app.post("/register", loginLimiter, async (req, res) => {
  const { employeeId, phone, password, securityQuestions, deviceId } = req.body;
console.log(
  "REGISTER_API_HIT",
  {
    employeeId,
    deviceId
  }
);
  if (!employeeId || !phone || !password || !deviceId) {
console.warn(
  "REGISTER_MISSING_FIELDS",
  {
    employeeId
  }
);
    return res.status(400).json({ message: "All fields are required." });
  }

  if (!Array.isArray(securityQuestions) || securityQuestions.length !== 3) {
console.warn(
  "REGISTER_INVALID_SECURITY_QUESTIONS",
  {
    employeeId
  }
);
    return res.status(400).json({
      message: "Exactly 3 security questions with answers are required.",
    });
  }

  const normalizedPhone = normalizePhoneOrBranch(phone);

  try {
    // 1️⃣ Check employee in MASTER
    const empRows = await queryUTIDatabase(
      `
      SELECT [Emp No.], [Designation], [Br Code]
      FROM [dbo].[employees_master]
      WHERE [Emp No.] = ? AND [Mobile number] = ?
      `,
      [employeeId, normalizedPhone]
    );

    if (empRows.length === 0) {
await logSecurityEvent(
  employeeId,
  "Registration Enumeration",
  JSON.stringify({
    employeeId,
    phone: normalizedPhone,
    reason: "Employee ID or Phone not found"
  })
);
console.warn(
  "REGISTER_EMPLOYEE_NOT_FOUND",
  {
    employeeId,
    phone: normalizedPhone
  }
);
      return res.status(404).json({
        message: "Employee ID or Phone not found.",
      });
    }

    const emp = empRows[0];

    // 2️⃣ Already registered?
    const authRows = await queryUTIDatabase(
  `
  SELECT
    [Password],
    [GA_Secret],
    [device_id],
    [security_q1],
    [security_q2],
    [security_q3]
  FROM [dbo].[employees_auth]
  WHERE [Emp No.] = ?
  `,
  [employeeId]
);

if (authRows.length === 0) {
  // This should never happen if trigger is working
console.warn(
  "REGISTER_AUTH_RECORD_MISSING",
  {
    employeeId
  }
);
  return res.status(403).json({
    message: "User not enabled for application access. Contact admin."
  });
}

const auth = authRows[0];

// 🚫 If ANY key is already filled → block registration
const alreadyRegistered =
  auth.Password ||
  auth.GA_Secret ||
  auth.device_id ||
  auth.security_q1 ||
  auth.security_q2 ||
  auth.security_q3;

if (alreadyRegistered) {
console.warn(
  "REGISTER_ALREADY_COMPLETED",
  {
    employeeId
  }
);
  return res.status(409).json({
    message: "Registration already completed. Contact admin if reset is required."
  });
}


    // 3️⃣ Device check
    const deviceRows = await queryUTIDatabase(
  `
  SELECT [Emp No.]
  FROM [dbo].[employees_auth]
  WHERE [device_id] = ?
    AND [Emp No.] <> ?
  `,
  [deviceId, employeeId]
);

if (deviceRows.length > 0) {
console.warn(
  "REGISTER_DEVICE_ALREADY_USED",
  {
    employeeId,
    deviceId
  }
);
  return res.status(409).json({
    message: "This device is already registered to another employee.",
  });
}


    // 4️⃣ Decide Level
    let level = "Level 1";

    if (emp.Designation?.includes("Cluster Head")) {
      level = "Level 2";
    } else if (emp["Br Code"] === 9999) {
      level = "Level 3";
    }
console.log(
  "REGISTER_LEVEL_ASSIGNED",
  {
    employeeId,
    level
  }
);
    // 5️⃣ Generate GA
    const secret = speakeasy.generateSecret({ length: 32 });
console.log(
  "REGISTER_GA_SECRET_GENERATED",
  {
    employeeId
  }
);
    // 6️⃣ Insert AUTH record
   // 6️⃣ UPDATE AUTH record (CORRECT)
const result = await queryUTIDatabase(
  `
  UPDATE [dbo].[employees_auth]
  SET
    [Password] = ?,
    [Approval status] = 'pending',
    [role] = 'user',
    [GA_Secret] = ?,
    [security_q1] = ?, [security_a1] = ?,
    [security_q2] = ?, [security_a2] = ?,
    [security_q3] = ?, [security_a3] = ?,
    [device_id] = ?,
    [Level] = ?
  WHERE [Emp No.] = ?
  `,
  [
    password,
    secret.base32,
    securityQuestions[0].question,
    securityQuestions[0].answer,
    securityQuestions[1].question,
    securityQuestions[1].answer,
    securityQuestions[2].question,
    securityQuestions[2].answer,
    deviceId,
    level,
    employeeId,
  ]
);
console.log(
  "REGISTER_AUTH_RECORD_UPDATED",
  {
    employeeId,
    level
  }
);
    logActivity(employeeId, "User", "Register", `Device:${deviceId}`);
console.log(
  "REGISTER_SUCCESS",
  {
    employeeId,
    level
  }
);
    return res.status(200).json({
      message: "Registration successful. Awaiting admin approval.",
      googleAuthKey: secret.base32,
    });

  } catch (err) {
 console.error(
  "REGISTER_FAILED",
  {
    employeeId: req.body?.employeeId,
    error: err.message,
    stack: err.stack
  }
);
    return res.status(500).json({
      message: "Server error during registration.",
    });
  }
});

//============================================================================================
//                                           LOGIN
//=============================================================================================
//========================= /login Route =========================//
app.post("/login", loginLimiter, async (req, res) => {
  try {
    const { employeeId, password, deviceId, token } = req.body;
console.log(
  "LOGIN_API_HIT",
  {
    employeeId,
    deviceId
  }
);
    if (!employeeId || !password || !deviceId || !token) {

      return res.status(400).json({ message: "All fields are required." });
    }

    /* =================================================
       1️⃣ FETCH MASTER + AUTH DATA
    ================================================= */
    const rows = await queryUTIDatabase(
      `
      SELECT
        m.[Emp No.],
        m.[Employee Name],
        m.[Br Code],
        m.[Branch Name],
        m.[Designation],

        a.[Password],
        a.[Approval status],
        a.[role],
        a.[Level],
        a.[failed_attempts],
        a.[reset_required],
        a.[account_locked],
        a.[device_id],
        a.[GA_Secret]
      FROM [dbo].[employees_master] m
      JOIN [dbo].[employees_auth] a
        ON m.[Emp No.] = a.[Emp No.]
      WHERE m.[Emp No.] = ?
      `,
      [employeeId]
    );

    if (rows.length === 0) {
      logActivity(
        String(employeeId || "UNKNOWN"),
        "Unknown",
        "Login Failed",
        "Employee not found"
      );
await logSecurityEvent(
  employeeId,
  "Login Enumeration",
  JSON.stringify({
    employeeId,
    deviceId,
    reason: "Employee not found"
  })
);
console.warn(
  "LOGIN_EMPLOYEE_NOT_FOUND",
  {
    employeeId
  }
);
      return res.status(404).json({ message: "Employee not found." });
    }

    const user = rows[0];
    const userId = user["Emp No."].toString();
    const role = user.role || "user";

    /* =================================================
       2️⃣ ACCOUNT STATUS CHECKS
    ================================================= */
    if (user.account_locked) {
      logActivity(userId, role, "Login Failed", "Account locked");
console.warn(
  "LOGIN_ACCOUNT_LOCKED",
  {
    employeeId
  }
);
      return res.status(403).json({ message: "Account locked. Contact admin." });
    }

    if (user["Approval status"] !== "approved") {
      logActivity(userId, role, "Login Failed", "Account not approved");
console.warn(
  "LOGIN_ACCOUNT_NOT_APPROVED",
  {
    employeeId
  }
);
      return res.status(403).json({
        message: "Account not approved by admin yet.",
      });
    }

    if (user.reset_required) {
      logActivity(userId, role, "Login Failed", "Password reset required");
console.warn(
  "LOGIN_PASSWORD_RESET_REQUIRED",
  {
    employeeId
  }
);
      return res.status(403).json({
        message: "Password reset required before login.",
        forceForgotPassword: true,
      });
    }

    /* =================================================
       3️⃣ PASSWORD CHECK
    ================================================= */
// ❌ Wrong password
if (user.Password !== password) {
 const newAttempts = (user.failed_attempts || 0) + 1;
console.warn(
  "LOGIN_INVALID_PASSWORD",
  {
    employeeId,
    attempts: newAttempts
  }
);

  await queryUTIDatabase(
    `
    UPDATE [dbo].[employees_auth]
    SET failed_attempts = ?
    WHERE [Emp No.] = ?
    `,
    [newAttempts, employeeId]
  );

  logActivity(userId, role, "Login Failed", "Invalid password");

  // ✅ ONLY ON 3RD ATTEMPT
  if (newAttempts >= 3) {
    await queryUTIDatabase(
      `
      UPDATE [dbo].[employees_auth]
      SET reset_required = 1
      WHERE [Emp No.] = ?
      `,
      [employeeId]
    );

    return res.status(403).json({
      message: "Password reset required before login.",
      forceForgotPassword: true,   // 🔥 THIS FLAG IS IMPORTANT
    });
  }

  // ✅ 1st & 2nd attempt → NORMAL ERROR ONLY
  return res.status(401).json({
    message: "Invalid password.",
  });
}


    /* =================================================
       4️⃣ GOOGLE AUTHENTICATOR CHECK
    ================================================= */
    const verified = speakeasy.totp.verify({
      secret: user.GA_Secret,
      encoding: "base32",
      token,
      window: 1,
    });

    if (!verified) {
      logActivity(userId, role, "Login Failed", "Invalid Google Auth code");
console.warn(
  "LOGIN_INVALID_GA_TOKEN",
  {
    employeeId
  }
);
      return res.status(401).json({
        message: "Invalid Google Authenticator code.",
      });
    }

    /* =================================================
   5️⃣ DEVICE BINDING
================================================= */

// ✅ Allow testing accounts to login from multiple devices
const multiDeviceUsers = ["1", "3" , "65"];

if (multiDeviceUsers.includes(String(employeeId))) {
  console.log(`🧪 Multi-device login allowed for ${employeeId}`);
}
else if (!user.device_id) {
  await queryUTIDatabase(
    `
    UPDATE [dbo].[employees_auth]
    SET device_id = ?
    WHERE [Emp No.] = ?
    `,
    [deviceId, employeeId]
  );
}
else if (user.device_id !== deviceId) {
  logActivity(userId, role, "Login Failed", "Device mismatch");
console.warn(
  "LOGIN_DEVICE_MISMATCH",
  {
    employeeId
  }
);
  return res.status(403).json({
    message: "Login denied: Account bound to another device.",
  });
}

    /* =================================================
       6️⃣ SUCCESS LOGIN
    ================================================= */
    await queryUTIDatabase(
      `
      UPDATE [dbo].[employees_auth]
      SET failed_attempts = 0
      WHERE [Emp No.] = ?
      `,
      [employeeId]
    );

    logActivity(
      userId,
      role,
      "Login",
      `User ${user["Employee Name"]} logged in`
    );
const jwtToken = jwt.sign(
  {
    employeeId: userId,
    role: role,
    level: user.Level,
    branchCode: user["Br Code"],
    designation: user.Designation
  },
  process.env.JWT_SECRET,
  {
    expiresIn: '1h'
  }
);
console.log(
  "LOGIN_SUCCESS",
  {
    employeeId,
    role,
    level: user.Level
  }
);
    return res.status(200).json({
  success: true,
  token: jwtToken,
      message: `Welcome, ${user["Employee Name"]}!`,
      role,
      userId,
      level: user.Level,
      branchCode: user["Br Code"],
      designation: user.Designation,
      employeeName: user["Employee Name"],
      clusterName:
        user.Level === "Level 2" &&
        user.Designation?.includes("Cluster Head-")
          ? user.Designation.replace("Cluster Head-", "").trim()
          : null,
    });

  } catch (err) {
   console.error(
  "LOGIN_FAILED",
  {
    employeeId: req.body?.employeeId,
    error: err.message,
    stack: err.stack
  }
);

    logActivity(
      String(req.body?.employeeId || "UNKNOWN"),
      "Unknown",
      "Login Failed",
      `Server error: ${err.message}`
    );

    return res.status(500).json({
      message: "Server error during login.",
    });
  }
});


//============================================================================================
//                                          LOGOUT ENDPOINT
//=============================================================================================

app.post("/logout", authMiddleware, async (req, res) => {
 const { lastActivity } = req.body;

const userId = req.user.employeeId;
const role = req.user.role;
console.log(
  "LOGOUT_API_HIT",
  {
    employeeId: userId,
    role,
    lastActivity
  }
);
  try {
    await logActivity(

      userId,
      role,
      "Logout",
      `Last activity: ${lastActivity || "N/A"}`
    );
console.log(
  "LOGOUT_SUCCESS",
  {
    employeeId: userId,
    role
  }
);
    res.json({ success: true, message: "Logged out successfully" });
  } catch (err) {
   console.error(
  "LOGOUT_FAILED",
  {
    employeeId: userId,
    role,
    error: err.message
  }
);
    res.status(500).json({ success: false, message: "Logout failed" });
  }
});

//============================================================================================
//                                 FORGOT PASSWORD
//=============================================================================================

// ✅ Forgot Password - Step 1: Send back user's stored security questions
// ✅ Step 1: Forgot Password Start — Fetch Security Questions
app.post("/forgot-password/start", passwordLimiter, async (req, res) => {
  try {
    const { employeeId, phone } = req.body;
console.log(
  "FORGOT_PASSWORD_START_API_HIT",
  {
    employeeId,
    phone
  }
);
    if (!employeeId || !phone) {
console.warn(
  "FORGOT_PASSWORD_START_MISSING_INPUT",
  {
    employeeId,
    phone
  }
);
      return res
        .status(400)
        .json({ message: "Employee ID and phone are required." });
    }

    const query = `
      SELECT
        a.security_q1,
        a.security_q2,
        a.security_q3,
        a.role
      FROM [dbo].[employees_auth] a
      JOIN [dbo].[employees_master] m
        ON a.[Emp No.] = m.[Emp No.]
      WHERE a.[Emp No.] = ?
        AND m.[Mobile number] = ?
    `;

    const results = await queryUTIDatabase(query, [employeeId, phone]);

    if (results.length === 0) {
console.warn(
  "FORGOT_PASSWORD_START_EMPLOYEE_NOT_FOUND",
  {
    employeeId,
    phone
  }
);
await logSecurityEvent(
  employeeId,
  "Security Question Enumeration",
  JSON.stringify({
    employeeId,
    phone,
    reason: "Employee not found"
  })
);
      return res.status(404).json({ message: "Employee not found" });
    }

    const user = results[0];
    const role = user.role || "user";

    // Question map
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

    const securityQuestions = [
      { id: "q1", text: questionBank[user.security_q1] },
      { id: "q2", text: questionBank[user.security_q2] },
      { id: "q3", text: questionBank[user.security_q3] },
    ];

  await logActivity(
  employeeId,
  role,
  "Forgot Password Start",
  "Requested security questions"
);
console.log(
  "FORGOT_PASSWORD_START_SUCCESS",
  {
    employeeId,
    questionsReturned: securityQuestions.length
  }
);
    return res.json({ questions: securityQuestions });

} catch (err) {
console.error(
  "FORGOT_PASSWORD_START_FAILED",
  {
    employeeId: req.body?.employeeId,
    error: err.message,
    stack: err.stack
  }
);
    console.error("❌ /forgot-password/start error:", err);
    return res.status(500).json({ message: "Server error" });
  }
});



app.post("/verify-security-answers", passwordLimiter, async (req, res) => {
  try {
    const { employeeId, phone, answers } = req.body;
console.log(
  "VERIFY_SECURITY_ANSWERS_API_HIT",
  {
    employeeId,
    phone
  }
);
    // -----------------------------
    // 1️⃣ Fetch user
    // -----------------------------
    const rows = await queryUTIDatabase(
      `
      SELECT
        a.security_a1,
        a.security_a2,
        a.security_a3,
        a.reset_attempts,
        a.lock_until,
        a.account_locked,
        a.role
      FROM employees_auth a
      JOIN employees_master m
        ON a.[Emp No.] = m.[Emp No.]
      WHERE a.[Emp No.] = ?
        AND m.[Mobile number] = ?
      `,
      [employeeId, phone]
    );

    if (!rows.length) {
console.warn(
  "VERIFY_SECURITY_ANSWERS_EMPLOYEE_NOT_FOUND",
  {
    employeeId,
    phone
  }
);
await logSecurityEvent(
  employeeId,
  "Security Question Enumeration",
  `Unknown employee verification attempt`
);
      return res.status(404).json({ message: "Employee not found." });
    }

    const user = rows[0];
    const role = user.role || "user";

    // -----------------------------
    // 2️⃣ PERMANENT LOCK
    // -----------------------------
    if (user.account_locked) {
console.warn(
  "VERIFY_SECURITY_ANSWERS_ACCOUNT_LOCKED",
  {
    employeeId
  }
);
      return res.status(403).json({
        message: "Account locked. Contact admin.",
        permanentLock: true,
      });
    }

    // -----------------------------
    // 3️⃣ TEMP LOCK CHECK (SQL TIME)
    // -----------------------------
    const lockCheck = await queryUTIDatabase(
      `
      SELECT
        CASE
          WHEN lock_until IS NOT NULL AND lock_until > GETDATE()
          THEN DATEDIFF(MINUTE, GETDATE(), lock_until)
          ELSE 0
        END AS minutesLeft
      FROM employees_auth
      WHERE [Emp No.] = ?
      `,
      [employeeId]
    );

    const minutesLeft = lockCheck[0].minutesLeft;

    if (minutesLeft > 0) {
console.warn(
  "VERIFY_SECURITY_ANSWERS_TEMP_LOCK_ACTIVE",
  {
    employeeId,
    minutesLeft
  }
);
      return res.status(403).json({
        message: `Too many attempts. Try again after ${minutesLeft} minutes.`,
        lockActive: true,
        minutesLeft,
      });
    }

    // -----------------------------
    // 4️⃣ CLEAR EXPIRED LOCK
    // -----------------------------
    await queryUTIDatabase(
      `
      UPDATE employees_auth
      SET lock_until = NULL,
          reset_attempts = 0
      WHERE [Emp No.] = ?
        AND lock_until IS NOT NULL
        AND lock_until <= GETDATE()
      `,
      [employeeId]
    );

    // -----------------------------
    // 5️⃣ VERIFY ANSWERS
    // -----------------------------
    const norm = (v) => v?.trim().toLowerCase();
    let correct = 0;

    if (norm(answers.q1) === norm(user.security_a1)) correct++;
    if (norm(answers.q2) === norm(user.security_a2)) correct++;
    if (norm(answers.q3) === norm(user.security_a3)) correct++;

    if (correct >= 2) {
      await queryUTIDatabase(
        `
        UPDATE employees_auth
        SET reset_required = 1,
            reset_attempts = 0,
            lock_until = NULL
        WHERE [Emp No.] = ?
        `,
        [employeeId]
      );

   await logActivity(
  employeeId,
  role,
  "Forgot Password Verified",
  "Success"
);
console.log(
  "VERIFY_SECURITY_ANSWERS_SUCCESS",
  {
    employeeId,
    correctAnswers: correct
  }
);
      return res.json({
        message: "Verification successful. You may reset your password.",
        verified: true,
      });
    }

    // -----------------------------
    // 6️⃣ WRONG ANSWER → INCREMENT
    // -----------------------------
    const newAttempts = (user.reset_attempts || 0) + 1;

    await queryUTIDatabase(
      `
      UPDATE employees_auth
      SET reset_attempts = ?
      WHERE [Emp No.] = ?
      `,
      [newAttempts, employeeId]
    );

    // -----------------------------
    // 7️⃣ TEMP LOCK AT 3 ATTEMPTS
    // -----------------------------
    if (newAttempts === 3) {
      await queryUTIDatabase(
        `
        UPDATE employees_auth
        SET lock_until = DATEADD(MINUTE, 2, GETDATE())
        WHERE [Emp No.] = ?
        `,
        [employeeId]
      );
console.warn(
  "VERIFY_SECURITY_ANSWERS_TEMP_LOCK_CREATED",
  {
    employeeId,
    attempts: newAttempts
  }
);
      return res.status(403).json({
        message: "Too many wrong answers. Try again after 2 minutes.",
        lockActive: true,
        minutesLeft: 2,
      });
    }

    // -----------------------------
    // 8️⃣ PERMANENT LOCK AT 6
    // -----------------------------
    if (newAttempts >= 6) {
      await queryUTIDatabase(
        `
        UPDATE employees_auth
        SET account_locked = 1
        WHERE [Emp No.] = ?
        `,
        [employeeId]
      );
console.error(
  "VERIFY_SECURITY_ANSWERS_PERMANENT_LOCK",
  {
    employeeId,
    attempts: newAttempts
  }
);
      return res.status(403).json({
        message: "Account locked. Contact admin.",
        permanentLock: true,
      });
    }
console.warn(
  "VERIFY_SECURITY_ANSWERS_FAILED",
  {
    employeeId,
    attempts: newAttempts
  }
);
    // -----------------------------
    // 9️⃣ ATTEMPTS LEFT
    // -----------------------------
    return res.status(401).json({
      message: `Wrong answers. Attempts left: ${3 - newAttempts}`,
    });

} catch (err) {
console.error(
  "VERIFY_SECURITY_ANSWERS_ERROR",
  {
    employeeId: req.body?.employeeId,
    error: err.message,
    stack: err.stack
  }
);
    return res.status(500).json({ message: "Server error" });
  }
});




// ✅ Step 3: Reset Password
app.post("/forgot-password/reset", passwordLimiter, async (req, res) => {
  try {
    const { employeeId, phone, newPassword } = req.body;
console.log(
  "PASSWORD_RESET_API_HIT",
  {
    employeeId,
    phone
  }
);
    if (!employeeId || !phone || !newPassword) {
console.warn(
  "PASSWORD_RESET_MISSING_INPUT",
  {
    employeeId,
    phone
  }
);
      return res.status(400).json({
        message: "All fields required.",
      });
    }

    const rows = await queryUTIDatabase(
      `
      SELECT
        a.role,
        a.account_locked,
        a.reset_required
      FROM [dbo].[employees_auth] a
      JOIN [dbo].[employees_master] m
        ON a.[Emp No.] = m.[Emp No.]
      WHERE a.[Emp No.] = ?
        AND m.[Mobile number] = ?
      `,
      [employeeId, phone]
    );

    if (rows.length === 0) {
console.warn(
  "PASSWORD_RESET_EMPLOYEE_NOT_FOUND",
  {
    employeeId,
    phone
  }
);
await logSecurityEvent(
  employeeId,
  "Password Reset Enumeration",
  JSON.stringify({
    employeeId,
    phone,
    reason: "Employee not found"
  })
);
      return res.status(404).json({
        message: "Employee not found.",
      });
    }

    const user = rows[0];
    const role = user.role || "user";

    // 🔒 Permanent lock
    if (user.account_locked) {
console.warn(
  "PASSWORD_RESET_ACCOUNT_LOCKED",
  {
    employeeId
  }
);
      return res.status(403).json({
        message: "Account locked. Please contact admin.",
      });
    }

    // 🔐 Ensure security verification happened
    if (!user.reset_required) {
console.warn(
  "PASSWORD_RESET_NOT_AUTHORIZED",
  {
    employeeId
  }
);
      return res.status(403).json({
        message: "Password reset not authorized. Please verify security answers first.",
      });
    }

    // ✅ Reset password (DEVICE REMAINS SAME)
    await queryUTIDatabase(
      `
      UPDATE [dbo].[employees_auth]
      SET
        [Password] = ?,
        reset_required = 0,
        reset_attempts = 0,
        lock_until = NULL
      WHERE [Emp No.] = ?
      `,
      [newPassword, employeeId]
    );

await logActivity(
  employeeId,
  role,
  "Password Reset",
  "Password reset successful (device retained)"
);
console.log(
  "PASSWORD_RESET_SUCCESS",
  {
    employeeId
  }
);
    return res.json({
      message: "Password reset successful. You can login from your registered device.",
    });

} catch (err) {
console.error(
  "PASSWORD_RESET_FAILED",
  {
    employeeId: req.body?.employeeId,
    error: err.message,
    stack: err.stack
  }
);
    return res.status(500).json({
      message: "Server error",
    });
  }
});

//============================================================================================
//                                       CIRCULAR MANAGEMENT
//=============================================================================================

/**
 * POST /circulars
 * Body: { empNo, title, message }
 */
app.post("/circulars", loginLimiter, authMiddleware, async (req, res) => {

  console.log("📩 Incoming request: /circulars");

  try {
// 🔒 ADMIN ONLY
if (req.user.role !== "admin") {
console.warn(
  "CIRCULAR_CREATE_ACCESS_DENIED",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role
  }
);
  return res.status(403).json({
    success: false,
    message: "Access denied"
  });
}
   const empNo = parseInt(req.user.employeeId);
    let title = req.body.title;
    let message = req.body.message;
console.log(
  "CIRCULAR_CREATE_API_HIT",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    title
  }
);

    if (!empNo || !title || !message) {
console.warn(
  "CIRCULAR_CREATE_VALIDATION_FAILED",
  {
    employeeId: req.user?.employeeId
  }
);
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
console.log(
  "CIRCULAR_CREATE_SUCCESS",
  {
    employeeId: empNo,
    title
  }
);
    res.json({ ok: true, message: "Circular created successfully" });
  } catch (err) {
  console.error(
  "CIRCULAR_CREATE_FAILED",
  {
    employeeId: req.user?.employeeId,
    error: err.message
  }
);
    res.status(500).json({ message: "Error creating circular", error: err.message });
  }
});


/**
 * POST /circulars/delete
 * Body: { circularId, empNo }
 */
app.post("/circulars/delete", loginLimiter, authMiddleware, async (req, res) => {
console.log(
  "CIRCULAR_DELETE_API_HIT",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    circularId: req.body.circularId
  }
);
  try {
// 🔒 ADMIN ONLY
if (req.user.role !== "admin") {
console.warn(
  "CIRCULAR_DELETE_ACCESS_DENIED",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role
  }
);
  return res.status(403).json({
    success: false,
    message: "Access denied"
  });
}
    let circularId = parseInt(req.body.circularId);
 const empNo = req.user.employeeId;

    if (!circularId || !empNo) {
console.warn(
  "CIRCULAR_DELETE_VALIDATION_FAILED",
  {
    employeeId: req.user?.employeeId,
    circularId
  }
);
      return res.status(400).json({ message: "circularId and empNo required" });
    }

    // Check if exists
    const checkQuery = `SELECT id FROM circulars WHERE id = @param1`;
    const rows = await queryUTIDatabase(checkQuery, [circularId]);

    if (!rows.length) {
console.warn(
  "CIRCULAR_DELETE_NOT_FOUND",
  {
    employeeId: req.user?.employeeId,
    circularId
  }
);
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
console.log(
  "CIRCULAR_DELETE_SUCCESS",
  {
    employeeId: empNo,
    circularId
  }
);
    return res.json({ success: true, message: "Circular deleted successfully" });
  } catch (err) {
    console.error(
  "CIRCULAR_DELETE_FAILED",
  {
    employeeId: req.user?.employeeId,
    error: err.message
  }
);
    return res.status(500).json({ message: "Error deleting circular", error: err.message });
  }
});



/**
 * POST /circulars/fetch
 * Body: { empNo }
 * Returns only pending (unacknowledged) circulars for an employee
 */
app.post("/circulars/fetch", authMiddleware, async (req, res) => {
console.log(
  "CIRCULAR_FETCH_API_HIT",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role
  }
);
  try {
    const empNo = req.user.employeeId;

    if (!empNo) {
console.warn(
  "CIRCULAR_FETCH_VALIDATION_FAILED",
  {
    employeeId: req.user?.employeeId
  }
);
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
console.log(
  "CIRCULAR_FETCH_SUCCESS",
  {
    employeeId: empNo,
    circulars: rows.length
  }
);
await logActivity(
  empNo,
  req.user.role,
  "View Circulars",
  `Fetched ${rows.length} circulars`
);
    return res.json(rows);
  } catch (err) {
    console.error(
  "CIRCULAR_FETCH_FAILED",
  {
    employeeId: req.user?.employeeId,
    error: err.message
  }
);
    return res.status(500).json({ message: "Error fetching circulars" });
  }
});

/**
 * POST /circulars/acknowledge
 * Body: { circularId, empNo }
 */
app.post("/circulars/acknowledge", authMiddleware, async (req, res) => {
  try {
    const empNo = req.user.employeeId;
	const { circularId } = req.body;
console.log(
  "CIRCULAR_ACKNOWLEDGE_API_HIT",
  {
    employeeId: req.user?.employeeId,
    circularId
  }
);
    if (!circularId || !empNo) {
console.warn(
  "CIRCULAR_ACKNOWLEDGE_VALIDATION_FAILED",
  {
    employeeId: req.user?.employeeId,
    circularId
  }
);
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
console.log(
  "CIRCULAR_ACKNOWLEDGE_SUCCESS",
  {
    employeeId: empNo,
    circularId
  }
);
await logActivity(
  empNo,
  req.user.role,
  "Acknowledge Circular",
  `Circular ID ${circularId}`
);
    return res.json({ success: true });

  } catch (err) {
    console.error(
  "CIRCULAR_ACKNOWLEDGE_FAILED",
  {
    employeeId: req.user?.employeeId,
    circularId,
    error: err.message
  }
);
    return res.status(500).json({ message: "Error acknowledging circular" });
  }
});


/**
 * POST /circulars/history
 * Admin endpoint: Returns all circulars with acknowledgement count
 * Body: { adminId }
 */
app.post("/circulars/history", authMiddleware, async (req, res) => {
console.log(
  "CIRCULAR_HISTORY_API_HIT",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role
  }
);
  try {
// 🔒 ADMIN ONLY
if (req.user.role !== "admin") {
console.warn(
  "CIRCULAR_HISTORY_ACCESS_DENIED",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role
  }
);
  return res.status(403).json({
    success: false,
    message: "Access denied"
  });
}
  const empNo = req.user.employeeId;

	const cleanEmpNo = parseInt(empNo || 0);


 const sql = `
  SELECT
    c.id,
    c.title,
    c.message,
    FORMAT(c.created_at,'dd-MM-yyyy hh:mm:ss tt') AS created_at,
    c.emp_no,
    e.[Employee Name] AS uploader_name,
    (
      SELECT COUNT(*)
      FROM circular_acknowledgements ca
      WHERE ca.circular_id = c.id
    ) AS acknowledged_count
  FROM circulars c
  LEFT JOIN employees_master e ON e.[Emp No.] = c.emp_no
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
console.log(
  "CIRCULAR_HISTORY_SUCCESS",
  {
    employeeId: cleanEmpNo,
    records: result.length
  }
);
    return res.json(result);
  } catch (err) {
    console.error(
  "CIRCULAR_HISTORY_FAILED",
  {
    employeeId: req.user?.employeeId,
    error: err.message
  }
);
    return res.status(500).json({
      message: "Error fetching circular history",
      error: err.message,
    });
  }
});



// ===========================================================================================================
//                                            USER STATS
// ===========================================================================================================
app.post("/user-stats", authMiddleware, async (req, res) => {
  try {
console.log(
  "USER_STATS_API_HIT",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role
  }
);
	  // 🔒 ADMIN ONLY
if (req.user.role !== "admin") {

  console.warn(
    "USER_STATS_ACCESS_DENIED",
    {
      employeeId: req.user?.employeeId,
      role: req.user?.role
    }
  );
  return res.status(403).json({
    success: false,
    message: "Access denied"
  });
}
    const query = `
      SELECT
        CAST(SUM(CASE WHEN a.[Approval status] = 'approved' THEN 1 ELSE 0 END) AS INT) AS approved,
        CAST(SUM(CASE WHEN a.[Approval status] = 'pending' THEN 1 ELSE 0 END) AS INT) AS pending,
        CAST(SUM(CASE WHEN a.[Approval status] = 'rejected' THEN 1 ELSE 0 END) AS INT) AS rejected,
        CAST(SUM(CASE WHEN a.account_locked = 1 THEN 1 ELSE 0 END) AS INT) AS locked,
        CAST(COUNT(a.[Emp No.]) AS INT) AS total_registered,
        CAST(
          (
            SELECT COUNT(*)
            FROM [dbo].[employees_master] m
            WHERE NOT EXISTS (
              SELECT 1
              FROM [dbo].[employees_auth] a2
              WHERE a2.[Emp No.] = m.[Emp No.]
                AND a2.[Password] IS NOT NULL
            )
          ) AS INT
        ) AS not_registered
      FROM [dbo].[employees_auth] a;
    `;

    const rows = await queryUTIDatabase(query, []);
    const row = rows[0] || {};
console.log(
  "USER_STATS_SUCCESS",
  {
    employeeId: req.user?.employeeId,
    approved: Number(row.approved || 0),
    pending: Number(row.pending || 0),
    rejected: Number(row.rejected || 0),
    locked: Number(row.locked || 0),
    totalRegistered: Number(row.total_registered || 0),
    notRegistered: Number(row.not_registered || 0)
  }
);
    return res.status(200).json({
      approved: Number(row.approved || 0),
      pending: Number(row.pending || 0),
      rejected: Number(row.rejected || 0),
      locked: Number(row.locked || 0),
      total_registered: Number(row.total_registered || 0),
      not_registered: Number(row.not_registered || 0),
    });

  } catch (err) {
   console.error(
  "USER_STATS_FAILED",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    error: err.message
  }
);
    return res.status(500).json({
      message: "Server error while fetching user stats",
    });
  }
});



// ======================================================================================================
//                                ✅ Get Users (with pagination & status)
// ======================================================================================================
app.post("/get-users", authMiddleware, async (req, res) => {
  try {
console.log("GET_USERS_API_HIT", {
  employeeId: req.user.employeeId,
  role: req.user.role,
  body: req.body,
});
	  // 🔒 ADMIN ONLY
if (req.user.role !== "admin") {
  return res.status(403).json({
    success: false,
    message: "Access denied"
  });
}
    const { status = "all", page = 1, limit = 10 } = req.body;
console.log("GET_USERS_FILTERS", {
  status,
  page,
  limit,
});
    const offset = (page - 1) * limit;

    let whereClause = "";

    if (status === "pending") {
      whereClause = "WHERE a.[Approval status] = 'pending'";
    } else if (status === "approved") {
      whereClause = "WHERE a.[Approval status] = 'approved'";
    } else if (status === "rejected") {
      whereClause = "WHERE a.[Approval status] = 'rejected'";
    } else if (status === "locked") {
      whereClause = "WHERE a.account_locked = 1";
    } else if (status === "not_registered") {
      whereClause = "WHERE a.[Password] IS NULL";
    }

console.log("GET_USERS_WHERE_CLAUSE", whereClause || "NO_FILTER");
    // status === "all" → no WHERE clause

    // 🔢 COUNT QUERY
    const countQuery = `
      SELECT COUNT(*) AS total
      FROM [dbo].[employees_master] m
      LEFT JOIN [dbo].[employees_auth] a
        ON m.[Emp No.] = a.[Emp No.]
      ${whereClause};
    `;

  const totalResult = await queryUTIDatabase(countQuery, []);

console.log("GET_USERS_COUNT_RESULT", totalResult);
    const total = totalResult[0]?.total || 0;

    // 📄 DATA QUERY
    const usersQuery = `
      SELECT
        m.[Emp No.],
        m.[Employee Name],
        m.[Mobile number],
        m.[Designation],
        m.[Branch Name],
        a.[Approval status],
        a.account_locked,
        a.[Level]
      FROM [dbo].[employees_master] m
      LEFT JOIN [dbo].[employees_auth] a
        ON m.[Emp No.] = a.[Emp No.]
      ${whereClause}
      ORDER BY m.[Emp No.]
      OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY;
    `;

    const results = await queryUTIDatabase(usersQuery, []);
console.log("GET_USERS_RESULTS_COUNT", results.length);

console.log(
  "GET_USERS_SAMPLE",
  results.slice(0, 3)
);
console.log("GET_USERS_RESPONSE", {
  total,
  page,
  limit,
  totalPages: Math.ceil(total / limit),
  usersReturned: results.length,
});
    return res.json({
      total,
      page: Number(page),
      limit: Number(limit),
      totalPages: Math.ceil(total / limit),
      users: results,
    });

  } catch (err) {
    console.error("❌ Error fetching users:", err);
    return res.status(500).json({
      message: "Server error while fetching users",
    });
  }
});


// =============================================================================================
//                                     ✅ Update User Status
// =============================================================================================
app.post("/update-user-status", loginLimiter, authMiddleware, async (req, res) => {
  try {
console.log("UPDATE_USER_STATUS_API_HIT", {
  employeeId: req.user.employeeId,
  role: req.user.role,
  body: req.body,
});
	  // 🔒 ADMIN ONLY
if (req.user.role !== "admin") {
  return res.status(403).json({
    success: false,
    message: "Access denied"
  });
}
   let { empNo, status } = req.body;

const adminId = req.user.employeeId;
const adminRole = req.user.role;

    if (!empNo || !status) {
      return res.status(400).json({
        message: "Employee number and status are required",
      });
    }

    // Clean employee number (handles commas like 1,011)
    const cleanEmpNo = parseInt(String(empNo).replace(/[, ]/g, ""), 10);
console.log("UPDATE_USER_STATUS_CLEAN_EMP", {
  original: empNo,
  cleanEmpNo,
  status,
});
    if (isNaN(cleanEmpNo)) {
      return res.status(400).json({
        message: "Invalid Employee Number",
      });
    }

    // Allow only valid statuses
    const allowedStatuses = ["approved", "rejected", "pending"];
    if (!allowedStatuses.includes(status)) {
      return res.status(400).json({
        message: "Invalid status value",
      });
    }
console.log("UPDATE_USER_STATUS_DB_UPDATE", {
  empNo: cleanEmpNo,
  status,
});
    // 1️⃣ Update approval status in AUTH table
    await queryUTIDatabase(
      `
      UPDATE [dbo].[employees_auth]
      SET [Approval status] = ?
      WHERE [Emp No.] = ?
      `,
      [status, cleanEmpNo]
    );

    // 2️⃣ Optional: If rejected, clear auth-sensitive fields
    if (status === "rejected") {
console.log("UPDATE_USER_STATUS_REJECTED_RESET", {
  empNo: cleanEmpNo,
});
      await queryUTIDatabase(
        `
        UPDATE [dbo].[employees_auth]
        SET
          [Password] = NULL,
          [GA_Secret] = NULL,
          [device_id] = NULL,
          [security_q1] = NULL,
          [security_a1] = NULL,
          [security_q2] = NULL,
          [security_a2] = NULL,
          [security_q3] = NULL,
          [security_a3] = NULL,
          failed_attempts = 0,
          reset_required = 0,
          account_locked = 0
        WHERE [Emp No.] = ?
        `,
        [cleanEmpNo]
      );
    }

    // 3️⃣ Log admin action
    logActivity(
      adminId || "system",
      adminRole || "admin",
      "Update User Status",
      `Set user ${cleanEmpNo} to ${status}`
    );
console.log("UPDATE_USER_STATUS_SUCCESS", {
  empNo: cleanEmpNo,
  status,
  adminId,
});
    return res.status(200).json({
      success: true,
      message: `User ${cleanEmpNo} ${status} successfully`,
    });

  } catch (err) {
    console.error("❌ Error updating user status:", err);
    return res.status(500).json({
      message: "Server error while updating user status",
    });
  }
});

// ===================================================================================================
//                                                ✅ Unlock User
// ===================================================================================================
// Utility to clean employee/admin numbers
function cleanNumber(value) {
  if (!value) return null;
  return value.toString().replace(/,/g, "").replace(/\.00$/, "").trim();
}

app.post("/unlock-user", loginLimiter, authMiddleware, async (req, res) => {
  try {
	  console.log("UNLOCK_USER_API_HIT", {
  employeeId: req.user.employeeId,
  role: req.user.role,
  body: req.body,
});
	      // 🔒 ADMIN ONLY
    if (req.user.role !== "admin") {
      return res.status(403).json({
        success: false,
        message: "Access denied"
      });
    }
	 let { empNo } = req.body;
 const adminId = req.user.employeeId;
const adminRole = req.user.role;

    if (!empNo) {
      return res.status(400).json({
        success: false,
        message: "Employee number is required",
      });
    }

    const cleanEmpNo = cleanNumber(empNo);
    const cleanAdminId = cleanNumber(adminId);
console.log("UNLOCK_USER_CLEAN_VALUES", {
  originalEmpNo: empNo,
  cleanEmpNo,
  cleanAdminId,
});

console.log("UNLOCK_USER_DB_UPDATE", {
  empNo: cleanEmpNo,
});
    // 1️⃣ Unlock user in AUTH table
    await queryUTIDatabase(
      `
      UPDATE [dbo].[employees_auth]
      SET
        failed_attempts = 0,
        account_locked = 0,
        attempts = 0,
        reset_attempts = 0,
        lock_until = NULL,
        reset_required = 0
      WHERE [Emp No.] = ?
      `,
      [cleanEmpNo]
    );
console.log("UNLOCK_USER_DB_UPDATE_SUCCESS", {
  empNo: cleanEmpNo,
});
    // 2️⃣ Log admin action (never block main flow)
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
console.log("UNLOCK_USER_SUCCESS_RESPONSE", {
  empNo: cleanEmpNo,
  adminId: cleanAdminId,
});
    return res.json({
      success: true,
      message: "User unlocked successfully",
    });

  } catch (err) {
    console.error("❌ Unlock user error:", err);
 return res.status(500).json({
  success: false,
  message: "Server error while unlocking user",
  error: err.message
});
  }
});

// --------------------------------------------------------------------------------------------
//          FETCH DISTINCT DISTRICTS & CLUSTERS (Mapping-based, NOT Hardcoded)
//---------------------------------------------------------------------------------------------

app.post("/get-districts-clusters", authMiddleware , async (req, res) => {
  try {
const level = req.user.level;
const designation = req.user.designation;
const userId = req.user.employeeId;
const role = req.user.role;
console.log(
  "GET_DISTRICTS_CLUSTERS_API_HIT",
  {
    employeeId: userId,
    role,
    level
  }
);
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
console.log(
  "GET_DISTRICTS_CLUSTERS_LEVEL2_RESTRICTION",
  {
    employeeId: userId,
    cluster: effectiveCluster
  }
);

      baseDistrictQuery += ` AND LTRIM(RTRIM([${clusterCol}])) = ?`;
      baseClusterQuery  += ` AND LTRIM(RTRIM([${clusterCol}])) = ?`;

      params.push(effectiveCluster);
    }

    // 4️⃣ Execute Queries
    const districtsRaw = await queryUTIDatabase(baseDistrictQuery, params);
    const clustersRaw  = await queryUTIDatabase(baseClusterQuery, params);
console.log(
  "GET_DISTRICTS_CLUSTERS_QUERY_SUCCESS",
  {
    employeeId: req.user.employeeId,
    districts: districtsRaw.length,
    clusters: clustersRaw.length
  }
);
    const districts = districtsRaw
      .map(d => d.District?.trim())
      .filter(Boolean);

    const clusters = clustersRaw
      .map(c => c.Cluster?.trim())
      .filter(c => c && c.toUpperCase() !== "CO");
console.log(
  "GET_DISTRICTS_CLUSTERS_SUCCESS",
  {
    employeeId: req.user.employeeId,
    districts: districts.length,
    clusters: clusters.length,
    fixedCluster: effectiveCluster || null
  }
);
    // 5️⃣ Return
    return res.status(200).json({
      success: true,
      districts,
      clusters,
      fixedCluster: effectiveCluster
    });

  } catch (err) {
console.error(
  "GET_DISTRICTS_CLUSTERS_FAILED",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    level: req.user?.level,
    error: err.message
  }
);
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

app.post(
  "/get-deposits",
  authMiddleware,
  async (req, res) => {
  try {
    let {
      branchCode,
      branchName,
      clusterName,
      districtName,
      page = 1,
      pageSize = 10,
      sortBy,                 // meaning (e.g., "gdm_deposits")
      sortOrder = "ASC",
    } = req.body;
	
const userId = req.user.employeeId;
const role = req.user.role;
let level = req.user.level;
const branchCodeFromToken = req.user.branchCode;
const designation = req.user.designation;

console.log(
  "GET_DEPOSITS_API_HIT",
  {
    employeeId: userId,
    role,
    level,
    branchCode,
    branchName,
    districtName,
    clusterName,
    page,
    pageSize
  }
);

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
    if (level === "Level 1") {
  branchCode = branchCodeFromToken;

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
console.warn(
  "GET_DEPOSITS_MAPPING_FALLBACK_USED",
  {
    employeeId: userId,
    section: SECTION
  }
);
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
  if (!effectiveCluster) {
    console.warn(
      "GET_DEPOSITS_CLUSTER_MISSING",
      {
        employeeId: userId,
        role
      }
    );

    return res.status(400).json({
      success: false,
      message: "Cluster missing for Level 2 user",
    });
  }

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
console.log(
  "GET_DEPOSITS_QUERY_STARTED",
  {
    employeeId: userId,
    level,
    paramsCount: params.length,
    page,
    pageSize
  }
);
    const dbRows = await queryUTIDatabase(query, params);
console.log(
  "GET_DEPOSITS_QUERY_SUCCESS",
  {
    employeeId: userId,
    records: dbRows.length
  }
);
    // OUT OF SCOPE CHECK — BEFORE formatting and before sending success:true
    // ❗ Level 1 users should NOT trigger out-of-scope check
    if (
      level !== "Level 1" &&
      dbRows.length === 0 &&
      branchCode &&
      (clusterName || districtName)
    ) {
console.warn(
  "GET_DEPOSITS_OUT_OF_SCOPE",
  {
    employeeId: userId,
    branchCode,
    clusterName,
    districtName
  }
);

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
await logActivity(
  userId,
  role,
  "View Deposits",
  `Level=${level}, Records=${totalRecords}`
);

console.log(
  "VIEW_DEPOSITS_ACTIVITY_LOGGED",
  {
    employeeId: userId
  }
);

console.log(
  "GET_DEPOSITS_SUCCESS",
  {
    employeeId: userId,
    totalRecords,
    page
  }
);

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
console.error(
  "GET_DEPOSITS_FAILED",
  {
    employeeId: userId,
    role,
    level,
    error: err.message,
    stack: err.stack
  }
);
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

app.post("/get-deposits-accounts", authMiddleware ,
async (req, res) => {
  console.log("🔹 /get-deposits-accounts called");

  try {
let {
  branchCode,
  branchName,
  districtName,
  clusterName,
  sortBy,
  sortOrder = "ASC",
  fetchAll = false,
  page = 1,
  pageSize = 10,
} = req.body;

const userId = req.user.employeeId;
const role = req.user.role;
let level = req.user.level;
const designation = req.user.designation;

console.log(
  "GET_DEPOSITS_ACCOUNTS_API_HIT",
  {
    employeeId: userId,
    role,
    level,
    branchCode,
    branchName,
    districtName,
    clusterName,
    page,
    pageSize
  }
);

const branchCodeFromToken = req.user.branchCode;
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
if (level === "Level 1") {
  branchCode = branchCodeFromToken;

  const numeric = parseFloat(branchCode);

  if (!Number.isNaN(numeric)) {
    branchCode = parseInt(numeric, 10).toString();
  }
}
    const SECTION = "deposits_accounts_opened";
const HIDDEN_DEPOSITS_MEANINGS = new Set([
  "vouchers",
  "vouchers_average",   
]);


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

let mapping = mappingRows
  .map((m) => ({
    colNumber: Number(m.col_number),
    colName: `col${m.col_number}`,
    displayLabel: m.display_label,
    meaning: m.meaning,
    type: m.detected_type || "string",
  }))
  .filter(
    (m) =>
      !m.meaning || !HIDDEN_DEPOSITS_MEANINGS.has(m.meaning) // 🔥 hide vouchers
  );

    if (!mapping.length) {
console.warn(
  "GET_DEPOSITS_ACCOUNTS_MAPPING_FALLBACK_USED",
  {
    employeeId: userId,
    section: SECTION
  }
);
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
  if (!effectiveCluster) {
    console.warn(
      "GET_DEPOSITS_ACCOUNTS_CLUSTER_MISSING",
      {
        employeeId: userId,
        role
      }
    );

    return res.status(400).json({
      success: false,
      message: "Cluster missing for Level 2 user.",
    });
  }

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
console.log(
  "GET_DEPOSITS_ACCOUNTS_QUERY_STARTED",
  {
    employeeId: userId,
    level,
    paramsCount: params.length,
    page,
    pageSize
  }
);
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
console.log(
  "GET_DEPOSITS_ACCOUNTS_QUERY_SUCCESS",
  {
    employeeId: userId,
    records: dbRows.length
  }
);
await logActivity(
  userId,
  role,
  "View Deposits Accounts",
  `Level=${level}, Records=${totalRecords}`
);

console.log(
  "VIEW_DEPOSITS_ACCOUNTS_ACTIVITY_LOGGED",
  {
    employeeId: userId
  }
);

console.log(
  "GET_DEPOSITS_ACCOUNTS_SUCCESS",
  {
    employeeId: userId,
    totalRecords,
    page
  }
);
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
    console.error(
  "GET_DEPOSITS_ACCOUNTS_FAILED",
  {
    employeeId: userId,
    role,
    level,
    error: err.message,
    stack: err.stack
  }
);
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

app.post(
  "/get-advances",
  authMiddleware,
  async (req, res) => {
  try {
let {
  branchCode,
  branchName,
  clusterName,
  districtName,
  page = 1,
  pageSize = 10,
  sortBy,
  sortOrder = "DESC",
} = req.body;

const userId = req.user.employeeId;
const role = req.user.role;
let level = req.user.level;

console.log(
  "GET_ADVANCES_API_HIT",
  {
    employeeId: userId,
    role,
    level,
    filters: {
      branchCode,
      branchName,
      clusterName,
      districtName
    },
    page,
    pageSize,
    sortBy,
    sortOrder
  }
);

const designation = req.user.designation;
const branchCodeFromToken = req.user.branchCode;

    // ---------- Normalize ----------
branchCode   = norm(branchCode);
branchName   = norm(branchName);
clusterName  = norm(clusterName);
districtName = norm(districtName);

    level = level || "Level 1";
	
	// ===============================
// FIX: Level 1 branchCode "2.00" → "2"
// ===============================
if (level === "Level 1") {
  branchCode = branchCodeFromToken;

  const num = parseFloat(branchCode);

  if (!isNaN(num)) {
    branchCode = parseInt(num, 10).toString();
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

  console.warn(
    "GET_ADVANCES_MAPPING_FALLBACK_USED",
    {
      employeeId: userId,
      section: SECTION
    }
  );
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

  // ✅ branch_code numeric compare
  if (meaning === "branch_code") {
    query += `
      AND TRY_CAST(LTRIM(RTRIM([${col}])) AS INT)
          = TRY_CAST(? AS INT)
    `;

    params.push(value);
    return;
  }

  // ✅ exact compare
  if (exact) {
    query += `
      AND LTRIM(RTRIM([${col}])) = ?
    `;

    params.push(value.trim());
  }

  // ✅ LIKE compare
  else {
    query += `
      AND LOWER(LTRIM(RTRIM([${col}]))) LIKE ?
    `;

    params.push(`%${value.trim().toLowerCase()}%`);
  }
};

    //========================================================================================
    // 4️⃣ LEVEL-BASED ACCESS CONTROL
    //========================================================================================
let effectiveCluster = clusterName;

if (level === "Level 2" && designation) {
  const m = designation.match(/Cluster\s*Head\s*[-:]\s*(.*)/i);

  if (m && m[1]) {
    effectiveCluster = m[1].trim();
  }
}

    if (level === "Level 1" && branchCode) {
      addFilter("branch_code", branchCode, true);
    }

if (level === "Level 2") {
  if (!effectiveCluster) {
console.warn(
  "GET_ADVANCES_CLUSTER_MISSING",
  {
    employeeId: userId,
    role,
    level
  }
);
    return res.status(400).json({
      success: false,
      message: "Cluster missing.",
    });
  }

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
console.log(
  "GET_ADVANCES_DATA_FETCHED",
  {
    employeeId: userId,
    records: dbRows.length,
    page
  }
);

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
if (meaning === "branch_code") {
  countQuery += `
    AND TRY_CAST(LTRIM(RTRIM([${col}])) AS INT)
        = TRY_CAST(? AS INT)
  `;

  countParams.push(value);
  return;
}
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

console.log(
  "VIEW_ADVANCES_ACTIVITY_LOGGED",
  {
    employeeId: userId,
    role
  }
);

    // Build meaningMap: displayLabel -> meaning (for frontend graphs)
    const meaningMap = {};
    mapping.forEach((m) => {
      if (m.meaning) {
        meaningMap[m.displayLabel] = m.meaning;
      }
    });

console.log(
  "GET_ADVANCES_SUCCESS",
  {
    employeeId: userId,
    role,
    totalRecords,
    totalPages: Math.ceil(totalRecords / pageSize),
    returnedRecords: data.length
  }
);

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
   console.error(
  "GET_ADVANCES_FAILED",
  {
    employeeId: userId,
    role,
    level,
    error: err.message
  }
);
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

app.post("/get-npa", authMiddleware, async (req, res) => {
  try {
let {
  branchCode,
  branchName,
  clusterName,
  districtName,
  page = 1,
  pageSize = 10,
  sortBy,
  sortOrder = "DESC",
} = req.body;

// ✅ Always take identity from JWT
const userId = req.user.employeeId;
const role = req.user.role;
let level = req.user.level;

console.log(
  "GET_NPA_API_HIT",
  {
    employeeId: userId,
    role,
    level,
    filters: {
      branchCode,
      branchName,
      clusterName,
      districtName
    },
    page,
    pageSize,
    sortBy,
    sortOrder
  }
);

const designation = req.user.designation;
const tokenBranchCode = req.user.branchCode;

// Normalize
branchCode   = norm(branchCode);
branchName   = norm(branchName);
clusterName  = norm(clusterName);
districtName = norm(districtName);

// ✅ Force Level 1 to use branch from JWT
if (level === "Level 1") {
  branchCode = tokenBranchCode;
}

level = level || "Level 1";

const HIDDEN_NPA_MEANINGS = new Set([
  "unstamped_npa_prev_accounts",
  "unstamped_npa_prev_balance",
  "npa_monthly_accounts",
  "npa_monthly_balance",
]);

// ✅ Level 1 branchCode "2.00" → "2"
if (level === "Level 1" && branchCode) {
  const num = parseFloat(branchCode);
  if (!isNaN(num)) {
    branchCode = parseInt(num, 10).toString();
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

    let mapping = (mappingRows || [])
  .map((m) => ({
    colNumber: Number(m.col_number),
    colName: `col${m.col_number}`,
    displayLabel: m.display_label,
    meaning: m.meaning || null,
    detectedType: m.detected_type || "string",
  }))
  .filter(
    (m) =>
      !m.meaning ||                    // keep columns without meaning
      !HIDDEN_NPA_MEANINGS.has(m.meaning) // 🔥 hide unwanted NPA columns
  );


    // Fallback (rare)
    if (!mapping.length) {
console.warn(
  "GET_NPA_MAPPING_FALLBACK_USED",
  {
    employeeId: userId,
    section: SECTION
  }
);
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
if (level === "Level 2" && !effectiveCluster) {

  console.warn(
    "GET_NPA_CLUSTER_MISSING",
    {
      employeeId: userId,
      role,
      level
    }
  );

  return res.status(400).json({
    success: false,
    message: "Cluster missing."
  });
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
console.log(
  "GET_NPA_DATA_FETCHED",
  {
    employeeId: userId,
    records: dbRows.length,
    page
  }
);
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
console.log(
  "VIEW_NPA_ACTIVITY_LOGGED",
  {
    employeeId: userId,
    role
  }
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

console.log(
  "GET_NPA_SUCCESS",
  {
    employeeId: userId,
    role,
    totalRecords,
    totalPages: Math.ceil(totalRecords / pageSize),
    returnedRecords: data.length
  }
);
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
   console.error(
  "GET_NPA_FAILED",
  {
    employeeId: userId,
    role,
    level,
    error: err.message
  }
);
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

app.post("/get-loans-sanctioned", authMiddleware, async (req, res) => {
  try {
let {
  branchCode,
  branchName,
  clusterName,
  districtName,
  page = 1,
  pageSize = 10,
  sortBy,
  sortOrder = "DESC",
} = req.body;

const userId = req.user.employeeId;
const role = req.user.role;
let level = req.user.level;

console.log(
  "GET_LOANS_SANCTIONED_API_HIT",
  {
    employeeId: userId,
    role,
    level,
    filters: {
      branchCode,
      branchName,
      clusterName,
      districtName
    },
    page,
    pageSize,
    sortBy,
    sortOrder
  }
);

const designation = req.user.designation;
const branchCodeFromToken = req.user.branchCode;

    // ------------------------------------------------------
    // Normalize values
    // ------------------------------------------------------
branchCode   = norm(branchCode);
branchName   = norm(branchName);
clusterName  = norm(clusterName);
districtName = norm(districtName);
	
level = level || "Level 1";

// ✅ Force Level 1 branch from JWT
if (level === "Level 1") {
  branchCode = branchCodeFromToken;

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
console.warn(
  "GET_LOANS_SANCTIONED_MAPPING_FALLBACK_USED",
  {
    employeeId: userId,
    section: SECTION
  }
);
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
 let effectiveCluster = clusterName;

if (level === "Level 2") {
      // Extract cluster from designation: "Cluster Head-XYZ" or "Cluster Head: XYZ"
      if (designation?.match(/Cluster\s*Head\s*[-:]\s*/i)) {
        effectiveCluster = designation.replace(
          /Cluster\s*Head\s*[-:]\s*/i,
          ""
        ).trim();
      } 
	  clusterName = effectiveCluster;

      if (!effectiveCluster) {

console.warn(
  "GET_LOANS_SANCTIONED_CLUSTER_MISSING",
  {
    employeeId: userId,
    role,
    level
  }
);

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
console.warn(
  "GET_LOANS_SANCTIONED_BRANCH_ACCESS_DENIED",
  {
    employeeId: userId,
    branchCode,
    cluster: effectiveCluster
  }
);
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
console.warn(
  "GET_LOANS_SANCTIONED_BRANCHNAME_ACCESS_DENIED",
  {
    employeeId: userId,
    branchName,
    cluster: effectiveCluster
  }
);
          return res.status(403).json({
            success: false,
            message: `❌ Branch ${branchName} does not belong to your cluster (${effectiveCluster}).`,
          });
        }
      }

      // Lock to this cluster
      addFilter("cluster", effectiveCluster, true);
    }
if (level === "Level 1" && branchCode) {
  addFilter("branch_code", branchCode, true);
}
    // General filters
if (level !== "Level 1" && branchCode)
  addFilter("branch_code", branchCode);

if (branchName)
  addFilter("branch_name", branchName, false);

if (districtName)
  addFilter("district", districtName, false);

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
console.log(
  "GET_LOANS_SANCTIONED_DATA_FETCHED",
  {
    employeeId: userId,
    records: dbRows.length,
    page
  }
);
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

try {
  await logActivity(
    userId,
    role,
    "View Loans Sanctioned",
    `Filters: Branch ${branchCode}, Cluster ${effectiveCluster}`
  );

  console.log(
    "VIEW_LOANS_SANCTIONED_ACTIVITY_LOGGED",
    {
      employeeId: userId,
      role
    }
  );
} catch (logErr) {
  console.error(
    "VIEW_LOANS_SANCTIONED_ACTIVITY_LOG_FAILED",
    {
      employeeId: userId,
      error: logErr.message
    }
  );
}

    //========================================================================================
    // 🔟 FINAL RESPONSE
    //========================================================================================
console.log(
  "GET_LOANS_SANCTIONED_SUCCESS",
  {
    employeeId: userId,
    role,
    totalRecords,
    totalPages,
    returnedRecords: data.length
  }
);
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
    console.error(
  "GET_LOANS_SANCTIONED_FAILED",
  {
    employeeId: userId,
    role,
    level,
    error: err.message
  }
);
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

app.post("/get-cluster-deposits-summary", authMiddleware, async (req, res) => {
  try {
    const userId = req.user.employeeId;
const role = req.user.role;
const level = req.user.level;
const designation = req.user.designation;

console.log(
  "CLUSTER_DEPOSITS_SUMMARY_API_HIT",
  {
    employeeId: userId,
    role,
    level
  }
);

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
console.warn(
  "CLUSTER_DEPOSITS_SUMMARY_MAPPING_MISSING",
  {
    employeeId: userId,
    section: SECTION
  }
);
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
 const colList = mapping
  .map(
    (m) =>
      `CASE 
         WHEN ISNUMERIC(REPLACE(REPLACE([${m.colName}], ',', ''), '-', '0')) = 1 
         THEN REPLACE(REPLACE([${m.colName}], ',', ''), '-', '0')
         ELSE [${m.colName}]
       END AS [${m.colName}]`
  )
  .join(", ");

    // 2️⃣ Build SQL with Level-2 filtering
    let query = `SELECT ${colList} FROM [dbo].[DCS] WHERE 1=1`;
    const params = [];

    let effectiveLevel = level || "Level 1";
    let effectiveCluster = null;

if (effectiveLevel === "Level 1") {
console.warn(
  "CLUSTER_DEPOSITS_SUMMARY_LEVEL1_ACCESS_DENIED",
  {
    employeeId: userId,
    role
  }
);
  return res.status(403).json({
    success: false,
    message: "Access denied. Cluster summary is not available for Level 1 users.",
  });
}

    // Level 2 → Determine Cluster
    if (effectiveLevel === "Level 2") {
      const match = designation?.match(/Cluster\s*Head\s*[-:]\s*(.*)/i);

if (match && match[1]) {
  effectiveCluster = match[1].trim();
}

      if (effectiveCluster) {
        query += ` AND LOWER(LTRIM(RTRIM([${clusterCol}]))) = ?`;
        params.push(effectiveCluster.toLowerCase());
      } else {
      console.warn(
  "CLUSTER_DEPOSITS_SUMMARY_CLUSTER_MISSING",
  {
    employeeId: userId,
    role
  }
);
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
console.log(
  "CLUSTER_DEPOSITS_SUMMARY_QUERY_STARTED",
  {
    employeeId: userId,
    level,
    paramsCount: params.length
  }
);
    // 4️⃣ Run Query
    const rows = await queryUTIDatabase(query, params);
console.log(
  "CLUSTER_DEPOSITS_SUMMARY_QUERY_SUCCESS",
  {
    employeeId: userId,
    records: rows.length
  }
);
    if (!rows.length) {
console.warn(
  "CLUSTER_DEPOSITS_SUMMARY_NO_RECORDS",
  {
    employeeId: userId,
    level
  }
);
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
await logActivity(
  userId,
  role,
  "View Cluster Deposits Summary",
  `Level=${level}`
);
console.log(
  "VIEW_CLUSTER_DEPOSITS_SUMMARY_ACTIVITY_LOGGED",
  {
    employeeId: userId
  }
);
console.log(
  "CLUSTER_DEPOSITS_SUMMARY_SUCCESS",
  {
    employeeId: userId,
    totalRecords
  }
);
    // 9️⃣ Return response
    return res.json({
  success: true,
  headers: staticHeaders.length ? staticHeaders : ["Cluster"],
  groupedHeaders,
  data,
  totalRecords,
});

 } catch (err) {
  console.error(
    "CLUSTER_DEPOSITS_SUMMARY_FAILED",
    {
      employeeId: req.user?.employeeId,
      role: req.user?.role,
      level: req.user?.level,
      error: err.message
    }
  );

  return res.status(500).json({
    success: false,
    error: "Server error while fetching Deposits Summary",
  });
}
});

//============================================================================================
//                        DEPOSITS ACCOUNTS OPENED (DAOCS) CLUSTER SUMMARY
//============================================================================================
app.post("/get-daocs", authMiddleware, async (req, res) => {
  try {
    const userId = req.user.employeeId;
const role = req.user.role;
const level = req.user.level;
const designation = req.user.designation;
const clusterName = req.user.clusterName || null;
const branchCode = req.user.branchCode;
console.log(
  "GET_DAOCS_API_HIT",
  {
    employeeId: userId,
    role,
    level,
    clusterName,
    branchCode
  }
);
    const SECTION = "DAOCS"; // mapping section
const HIDDEN_DAOCS_MEANINGS = new Set([
  "vouchers",
  "vouchers_average",
]);

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
console.warn(
  "GET_DAOCS_MAPPING_MISSING",
  {
    employeeId: userId,
    section: SECTION
  }
);
      return res.status(400).json({
        success: false,
        message:
          "⚠️ DAOCS summary mapping not found. Please upload DAOCS MIS first.",
      });
    }
const mapping = mappingRows
  .map((m) => ({
    colNumber: Number(m.col_number),
    colName: `col${m.col_number}`,
    displayLabel: m.display_label,
    meaning: m.meaning || null,
  }))
  .filter(
    (m) =>
      !m.meaning || !HIDDEN_DAOCS_MEANINGS.has(m.meaning) // 🔥 hide vouchers
  );

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
	
	// 🔒 Level 1 users cannot access cluster summaries
if (level === "Level 1") {
console.warn(
  "GET_DAOCS_LEVEL1_ACCESS_DENIED",
  {
    employeeId: userId,
    role
  }
);
  return res.status(403).json({
    success: false,
    message: "Access denied. Cluster summary is not available for Level 1 users.",
  });
}

    // ------ LEVEL-2 RESTRICTION -------
if (level === "Level 2") {
  let cluster = null;

  const match = designation?.match(/Cluster\s*Head\s*[-:]\s*(.*)/i);

  if (match && match[1]) {
    cluster = match[1].trim().toLowerCase();
  }

      if (cluster) {
        query += ` AND LOWER(LTRIM(RTRIM([${clusterCol}]))) = ?`;
        params.push(cluster);
      }
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
console.log(
  "GET_DAOCS_QUERY_STARTED",
  {
    employeeId: userId,
    level,
    paramsCount: params.length
  }
);
    const rows = await queryUTIDatabase(query, params);
console.log(
  "GET_DAOCS_QUERY_SUCCESS",
  {
    employeeId: userId,
    records: rows.length
  }
);
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
console.warn(
  "GET_DAOCS_NO_RECORDS",
  {
    employeeId: userId,
    level
  }
);
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
    title: "Current",
    cols: [
      makeColObj("current_accounts", "Current A/Cs"),
      makeColObj("current_amount", "Current Amount"),
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
console.log(
  "VIEW_DAOCS_ACTIVITY_LOGGED",
  {
    employeeId: userId
  }
);
console.log(
  "GET_DAOCS_SUCCESS",
  {
    employeeId: userId,
    totalRecords: data.length
  }
);
    return res.json({
      success: true,
      headers: staticHeaders,
      groupedHeaders,
      data,
      totalRecords: data.length,
    });
  } catch (err) {
  console.error(
    "GET_DAOCS_FAILED",
    {
      employeeId: req.user?.employeeId,
      role: req.user?.role,
      level: req.user?.level,
      error: err.message,
      stack: err.stack
    }
  );

  return res.status(500).json({
    success: false,
    message: "Server error while fetching DAOCS summary",
  });
}
});

//============================================================================================
//                     ADVANCES SUMMARY (Dynamic Mapping + Intelligent Grouping)
//============================================================================================

app.post("/get-advances-summary", authMiddleware, async (req, res) => {
const userId = req.user.employeeId;
const role = req.user.role;
const level = req.user.level;
const designation = req.user.designation;
  try {

console.log(
  "ADVANCES_SUMMARY_API_HIT",
  {
    employeeId: userId,
    role,
    level
  }
);
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
console.warn(
  "ADVANCES_SUMMARY_MAPPING_MISSING",
  {
    employeeId: userId,
    section: SECTION
  }
);
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
	
	// 🔒 Level 1 users cannot access cluster summaries
if (level === "Level 1") {
console.warn(
  "ADVANCES_SUMMARY_LEVEL1_ACCESS_DENIED",
  {
    employeeId: userId,
    role
  }
);
  return res.status(403).json({
    success: false,
    message: "Access denied. Cluster summary is not available for Level 1 users.",
  });
}

if (level === "Level 2") {
  let effectiveCluster = null;

  const match = designation?.match(/Cluster\s*Head\s*[-:]\s*(.*)/i);

  if (match && match[1]) {
    effectiveCluster = match[1].trim().toLowerCase();
  }

  if (!effectiveCluster) {
console.warn(
  "ADVANCES_SUMMARY_CLUSTER_MISSING",
  {
    employeeId: userId,
    role
  }
);
    return res.status(403).json({
      success: false,
      message: "Cluster information missing for Level 2 user.",
    });
  }

  query += ` AND LOWER(LTRIM(RTRIM([${clusterCol}]))) = ?`;
  params.push(effectiveCluster);
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
console.log(
  "ADVANCES_SUMMARY_QUERY_STARTED",
  {
    employeeId: userId,
    level,
    paramsCount: params.length
  }
);
    const rows = await queryUTIDatabase(query, params);
console.log(
  "ADVANCES_SUMMARY_QUERY_SUCCESS",
  {
    employeeId: userId,
    records: rows.length
  }
);
    if (!rows.length) {
console.warn(
  "ADVANCES_SUMMARY_NO_RECORDS",
  {
    employeeId: userId,
    level
  }
);

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
await logActivity(
  userId,
  role,
  "View Advances Summary",
  `Level=${level}`
);

console.log(
  "VIEW_ADVANCES_SUMMARY_ACTIVITY_LOGGED",
  {
    employeeId: userId
  }
);
console.log(
  "ADVANCES_SUMMARY_SUCCESS",
  {
    employeeId: userId,
    totalRecords
  }
);
    return res.json({
      success: true,
      headers: staticHeaders,
      groupedHeaders,
      data,
      meaningMap,
      totalRecords,
    });
  } catch (err) {
   console.error(
  "ADVANCES_SUMMARY_FAILED",
  {
    employeeId: userId,
    role,
    level,
    error: err.message
  }
);
    return res.status(500).json({
      success: false,
      error: "Server error while fetching Advances Summary",
    });
  }
});



//============================================================================================
//                                      NPA SUMMARY (Dynamic Mapping + Grouped Headers)
//============================================================================================

app.post("/get-npa-summary", authMiddleware, async (req, res) => {
    const userId = req.user.employeeId;
const role = req.user.role;
const level = req.user.level;
const designation = req.user.designation;
  try {
console.log(
  "NPA_SUMMARY_API_HIT",
  {
    employeeId: userId,
    role,
    level
  }
);
    const SECTION = "NPACS";

    // 🔒 Hidden meanings (same as /get-npa)
    const HIDDEN_NPACS_MEANINGS = new Set([
      "unstamped_npa_prev_accounts",
      "unstamped_npa_prev_balance",
      "npa_monthly_accounts",
      "npa_monthly_balance",
    ]);

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
console.warn(
  "NPA_SUMMARY_MAPPING_MISSING",
  {
    employeeId: userId,
    section: SECTION
  }
);
      return res.status(400).json({
        success: false,
        message: "⚠️ NPA summary mapping not found. Please upload NPACS MIS first.",
      });
    }

    // 🔥 Apply hiding at mapping level
    const mapping = mappingRows
      .map((m) => ({
        colNumber: Number(m.col_number),
        colName: `col${m.col_number}`,
        displayLabel: m.display_label,
        meaning: m.meaning || null,
      }))
      .filter(
        (m) =>
          !m.meaning || !HIDDEN_NPACS_MEANINGS.has(m.meaning)
      );

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

    const getColByMeaning = (meaning) =>
      byMeaning[meaning] || fallback[meaning] || null;

    const clusterCol = getColByMeaning("cluster") || "col1";

    // Select list based on filtered mapping
    const colList = mapping.map((m) => `[${m.colName}]`).join(", ");

    // 2️⃣ Build SQL + level filter
    let query = `SELECT ${colList} FROM [dbo].[NPACS] WHERE 1=1`;
    const params = [];
	
	// 🔒 Level 1 users cannot access cluster summaries
if (level === "Level 1") {
console.warn(
  "NPA_SUMMARY_LEVEL1_ACCESS_DENIED",
  {
    employeeId: userId,
    role
  }
);
  return res.status(403).json({
    success: false,
    message: "Access denied. Cluster summary is not available for Level 1 users.",
  });
}

if (level === "Level 2") {
  let effectiveCluster = null;

  const match = designation?.match(
    /Cluster\s*Head\s*[-:]\s*(.*)/i
  );

  if (match && match[1]) {
    effectiveCluster = match[1].trim().toLowerCase();
  }

  if (!effectiveCluster) {
console.warn(
  "NPA_SUMMARY_CLUSTER_MISSING",
  {
    employeeId: userId,
    role
  }
);
    return res.status(403).json({
      success: false,
      message: "Cluster information missing for Level 2 user.",
    });
  }

  query += ` AND LOWER(LTRIM(RTRIM([${clusterCol}]))) = ?`;
  params.push(effectiveCluster);
}

    // 3️⃣ Custom ordering (unchanged)
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
console.log(
  "NPA_SUMMARY_QUERY_STARTED",
  {
    employeeId: userId,
    level,
    paramsCount: params.length
  }
);
    const rows = await queryUTIDatabase(query, params);
console.log(
  "NPA_SUMMARY_QUERY_SUCCESS",
  {
    employeeId: userId,
    records: rows.length
  }
);
await logActivity(
  userId,
  role,
  "View NPA Summary",
  `Level:${level || "All"}`
);
console.log(
  "VIEW_NPA_SUMMARY_ACTIVITY_LOGGED",
  {
    employeeId: userId
  }
);
    if (!rows || rows.length === 0) {
console.warn(
  "NPA_SUMMARY_NO_RECORDS",
  {
    employeeId: userId,
    level
  }
);
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
        obj[m.displayLabel] =
          raw === null || raw === undefined || raw === "" ? "-" : raw;
      });
      return obj;
    });

    // 5️⃣ Static headers
    const staticSet = new Set(["cluster", "branch_code", "branch_name", "district"]);

    const staticHeaders = mapping
      .filter((m) => m.meaning && staticSet.has(m.meaning))
      .map((m) => m.displayLabel);

    // 6️⃣ Grouped headers
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

      if (!groupMap.has(parent)) {
        groupMap.set(parent, { title: parent, cols: [] });
      }

      groupMap.get(parent).cols.push({
        key: m.displayLabel,
        label: child,
      });
    });

    const groupedHeaders = Array.from(groupMap.values());

    // 7️⃣ Meaning map (only visible columns)
    const meaningMap = {};
    mapping.forEach((m) => {
      if (m.meaning) {
        meaningMap[m.meaning] = m.displayLabel;
      }
    });

    const totalRecords = rows.length;
console.log(
  "NPA_SUMMARY_SUCCESS",
  {
    employeeId: userId,
    totalRecords
  }
);
    return res.json({
      success: true,
      headers: staticHeaders,
      groupedHeaders,
      data,
      meaningMap,
      totalRecords,
    });
  } catch (err) {
    console.error(
  "NPA_SUMMARY_FAILED",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    level: req.user?.level,
    error: err.message
  }
);
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
app.post("/get-loans-sanctioned-summary", authMiddleware, async (req, res) => {
    const userId = req.user.employeeId;
const role = req.user.role;
const level = req.user.level;
const designation = req.user.designation;
  try {
console.log(
  "LOANS_SANCTIONED_SUMMARY_API_HIT",
  {
    employeeId: userId,
    role,
    level
  }
);
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
console.warn(
  "LOANS_SANCTIONED_SUMMARY_MAPPING_MISSING",
  {
    employeeId: userId,
    section: SECTION
  }
);
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
	
	// 🔒 Level 1 users cannot access cluster summaries
if (level === "Level 1") {
console.warn(
  "LOANS_SANCTIONED_SUMMARY_LEVEL1_ACCESS_DENIED",
  {
    employeeId: userId,
    role
  }
);
  return res.status(403).json({
    success: false,
    message: "Access denied. Cluster summary is not available for Level 1 users.",
  });
}

if (level === "Level 2") {
  let effectiveCluster = null;

  const match = designation?.match(
    /Cluster\s*Head\s*[-:]\s*(.*)/i
  );

  if (match && match[1]) {
    effectiveCluster = match[1].trim().toLowerCase();
  }

  if (!effectiveCluster) {
console.warn(
  "LOANS_SANCTIONED_SUMMARY_CLUSTER_MISSING",
  {
    employeeId: userId,
    role
  }
);
    return res.status(403).json({
      success: false,
      message: "Cluster information missing for Level 2 user.",
    });
  }

  query += ` AND LOWER(LTRIM(RTRIM([${clusterCol}]))) = ?`;
  params.push(effectiveCluster);
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
console.log(
  "LOANS_SANCTIONED_SUMMARY_QUERY_STARTED",
  {
    employeeId: userId,
    level,
    paramsCount: params.length
  }
);
    const rows = await queryUTIDatabase(query, params);
console.log(
  "LOANS_SANCTIONED_SUMMARY_QUERY_SUCCESS",
  {
    employeeId: userId,
    records: rows.length
  }
);
await logActivity(
  userId,
  role,
  "View Loans Sanctioned Summary",
  `Level=${level}`
);

console.log(
  "VIEW_LOANS_SANCTIONED_SUMMARY_ACTIVITY_LOGGED",
  {
    employeeId: userId
  }
);

    if (!rows || rows.length === 0) {
console.warn(
  "LOANS_SANCTIONED_SUMMARY_NO_RECORDS",
  {
    employeeId: userId,
    level
  }
);
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
console.log(
  "LOANS_SANCTIONED_SUMMARY_SUCCESS",
  {
    employeeId: userId,
    totalRecords
  }
);
    return res.json({
      success: true,
      headers: staticHeaders,
      groupedHeaders,
      data,
      meaningMap,
      totalRecords,
    });
  } catch (err) {
   console.error(
  "LOANS_SANCTIONED_SUMMARY_FAILED",
  {
    employeeId: userId,
    role,
    level,
    error: err.message
  }
);
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
app.post("/get-deposits-by-cluster", authMiddleware, async (req, res) => {
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

app.post("/get-sma-denormalized", authMiddleware, async (req, res) => {
  try {
let {
  filters = {},
  page = 1,
  pageSize = 10,
  sortBy,
  sortOrder = "ASC",
} = req.body;

let userId = req.user.employeeId;
const role = req.user.role;
let level = req.user.level;
const designation = req.user.designation;
const branchCodeFromToken = req.user.branchCode;
    console.log("🔹 /get-sma-denormalized called");
console.log(
  "GET_SMA_DENORMALIZED_API_HIT",
  {
    employeeId: userId,
    role,
    level,
    filters,
    page,
    pageSize,
    sortBy
  }
);
    if (userId) {
      userId = parseInt(userId, 10);
    }

    const SECTION = "SMA";

    // ------------------------------------------------------
    // Normalize Filters
    // ------------------------------------------------------
// ⭐ FIXED FILTER PARSING
let {
  branchCode,
  branchName,
  districtName,
  clusterName
} = parseFilters(filters);


    // ⭐⭐⭐ LEVEL-1 MUST GET branch_code FROM USERS TABLE
if (level === "Level 1") {
  branchCode = branchCodeFromToken;

  const num = parseFloat(branchCode);

  if (!isNaN(num)) {
    branchCode = parseInt(num, 10).toString();
  }

  console.log("✔ Level-1 Branch Code Applied:", branchCode);
}

console.log("JWT Branch Code:", branchCodeFromToken);
console.log("Final Branch Code:", branchCode);
console.log("Level:", level);

    level = level || "Level 1";
    page = Number(page);
    pageSize = Number(pageSize);
    const offset = (page - 1) * pageSize;

    // ⭐⭐⭐ LEVEL-1 FIX: Normalise branchCode "2.00" → "2"
    if (level === "Level 1" && branchCode) {
      const num = parseFloat(branchCode);
      if (!isNaN(num)) branchCode = parseInt(num, 10).toString();
    }
const HIDDEN_SMA_MEANINGS = new Set([
  "sma_fy_accounts",
  "sma_fy_balance",
  "sma_monthly_accounts",   // 👈 ADD THIS
  "sma_monthly_balance",    // 👈 ADD THIS
]);

    //========================================================================================
    // 1️⃣ LOAD SMA MAPPING (meaning → colX)
    //========================================================================================
    const mappingRows = await queryUTIDatabase(
  `
  SELECT col_number, display_label, meaning, detected_type
  FROM MIS.dbo.MIS_Column_Mapping
  WHERE section = ?
    AND col_number BETWEEN 1 AND 30
  ORDER BY col_number
  `,
  [SECTION]
);


    if (!mappingRows.length) {
      return res.status(400).json({
        success: false,
        message: "SMA mapping not found. Upload SMA MIS first.",
      });
    }
const mapping = mappingRows
  .map((m) => ({
    colNumber: Number(m.col_number),
    colName: `col${m.col_number}`,
    displayLabel: m.display_label,
    meaning: m.meaning,
    type: m.detected_type,
  }))
  .filter(
    (m) =>
      !m.meaning ||                 // keep columns without meaning
      !HIDDEN_SMA_MEANINGS.has(m.meaning) // 🔥 hide FY SMA columns
  );


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
      query += ` AND LOWER(LTRIM(RTRIM(${qualifiedCol}))) LIKE ?`;
countQuery += ` AND LOWER(LTRIM(RTRIM(${qualifiedCol}))) LIKE ?`;
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

    await logSecurityEvent(
      userId,
      "SMA Access Denied",
      JSON.stringify({
        reason: "Cluster missing for Level 2 user"
      })
    );

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
console.log(
  "SMA_SORT_REQUEST",
  {
    sortBy,
    allowedSortableMeaning
  }
);
let orderExpr =
`TRY_CAST(LTRIM(RTRIM(s.[${orderCol}])) AS INT)`;
   if (sortBy && allowedSortableMeaning.includes(sortBy)) {
  orderCol = getCol(sortBy);

  const m = mapping.find((x) => x.colName === orderCol);

  if (
    m &&
    ["decimal", "int", "number"].includes(
      String(m.type || "").toLowerCase()
    )
  ) {
    orderExpr = `
      TRY_CAST(LTRIM(RTRIM(s.[${orderCol}])) AS DECIMAL(18,4))
    `;
  } else {
    orderExpr = `s.[${orderCol}]`;
  }
}

    const orderDir =
      sortOrder.toUpperCase() === "DESC" ? "DESC" : "ASC";

   query += `
  ORDER BY ${orderExpr} ${orderDir}
`;

    //========================================================================================
    // 7️⃣ PAGINATION
    //========================================================================================
    query += ` OFFSET ? ROWS FETCH NEXT ? ROWS ONLY`;
    params.push(offset, pageSize);


console.log("SMA QUERY:");
console.log(query);
console.log("PARAMS:");
console.log(params);
console.log(
  "SMA_QUERY_EXECUTION_STARTED",
  {
    employeeId: userId
  }
);
    //========================================================================================
    // 8️⃣ EXECUTE
    //========================================================================================
    const rows = await queryUTIDatabase(query, params);
	console.log(
  "SMA_QUERY_EXECUTION_COMPLETED",
  {
    employeeId: userId,
    rowsReturned: rows.length
  }
);
    const totalRecords =
      (await queryUTIDatabase(countQuery, countParams))[0]?.totalRecords || 0;

const totalPages = Math.max(
  1,
  Math.ceil(totalRecords / pageSize)
);

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
console.log(
  "GET_SMA_DENORMALIZED_SUCCESS",
  {
    employeeId: userId,
    rows: data.length,
    totalRecords,
    totalPages
  }
);

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
app.post("/sma-total", authMiddleware, async (req, res) => {
  try {
    let { level, clusterName } = req.body;
const userId = req.user.employeeId;
const role = req.user.role;

console.log(
  "SMA_TOTAL_API_HIT",
  {
    employeeId: userId,
    role,
    level,
    clusterName
  }
);
// =====================================================
// JWT SCOPE ENFORCEMENT
// =====================================================
const userLevel = req.user.level;
const designation = req.user.designation;

if (userLevel === "Level 1") {
console.warn(
  "SMA_TOTAL_LEVEL1_ACCESS_DENIED",
  {
    employeeId: userId,
    role
  }
);
  return res.status(403).json({
    success: false,
    message: "Access denied",
  });
}

if (userLevel === "Level 2") {
  level = "Level 2";

  const match = designation?.match(
    /Cluster\s*Head\s*-\s*(.*)/i
  );

  if (match) {
    clusterName = match[1].trim();

    console.log(
      "SMA_TOTAL_LEVEL2_CLUSTER_RESOLVED",
      {
        employeeId: userId,
        clusterName
      }
    );
  }
} // LEVEL 2 BLOCK ENDS HERE

console.log("AFTER_LEVEL_CHECK", {
  userLevel,
  level,
  clusterName
});

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
console.warn(
  "SMA_TOTAL_MAPPING_MISSING",
  {
    employeeId: userId
  }
);
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
console.warn(
  "SMA_TOTAL_MEANING_MISSING",
  {
    employeeId: userId,
    meaning: m
  }
);
        return res.json({
          success: false,
          message: `Meaning '${m}' missing in SMA mapping.`,
        });
      }
    }

    const c = meaningToCol;
// ----------------------------------
// Identifier columns
// ----------------------------------
const clusterCol =
  "col" +
  mappingSMA.find(x => x.meaning === "cluster")?.col_number;

const branchCodeCol =
  "col" +
  mappingSMA.find(x => x.meaning === "branch_code")?.col_number;

console.log("cluster mapping =", mappingSMA.find(x => x.meaning === "cluster"));
console.log("branch mapping =", mappingSMA.find(x => x.meaning === "branch_code"));

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
console.warn(
  "SMA_TOTAL_ADVANCES_NOT_FOUND",
  {
    employeeId: userId,
    level: cleanLevel,
    cluster: cleanCluster
  }
);
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
console.log(
  "SMA_TOTAL_QUERY_STARTED",
  {
    employeeId: userId,
    level: cleanLevel,
    cluster: cleanCluster || "ALL"
  }
);

console.log("cleanLevel =", cleanLevel);
console.log("cleanCluster =", cleanCluster);
console.log("clusterCol =", clusterCol);
console.log("branchCodeCol =", branchCodeCol);

    let smaRows = [];

if (cleanLevel === "grandtotal") {
  const q = `
    SELECT ${allSums}
    FROM SMA
    WHERE ISNUMERIC([${branchCodeCol}]) = 1
  `;
      smaRows = await queryUTIDatabase(q);
    } else if (["level 2", "level 3"].includes(cleanLevel)) {
  const q = `
    SELECT ${allSums}
    FROM SMA
    WHERE UPPER(LTRIM(RTRIM([${clusterCol}]))) = UPPER(?)
      AND ISNUMERIC([${branchCodeCol}]) = 1
  `;
      smaRows = await queryUTIDatabase(q, [cleanCluster]);
    }

    if (!smaRows.length) {
console.warn(
  "SMA_TOTAL_NO_DATA_FOUND",
  {
    employeeId: userId,
    level: cleanLevel,
    cluster: cleanCluster
  }
);
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
await logActivity(
  userId,
  role,
  "View SMA Summary",
  `Level=${cleanLevel}, Cluster=${cleanCluster || "ALL"}`
);

console.log(
  "VIEW_SMA_SUMMARY_ACTIVITY_LOGGED",
  {
    employeeId: userId
  }
);
console.log(
  "SMA_TOTAL_SUCCESS",
  {
    employeeId: userId,
    level: cleanLevel,
    cluster: cleanCluster || "ALL"
  }
);
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
    console.error(
  "SMA_TOTAL_FAILED",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    level: req.body?.level,
    clusterName: req.body?.clusterName,
    error: err.message,
    stack: err.stack
  }
);
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
app.post("/update-daily-summary", authMiddleware, async (req, res) => {
  try {

    // =====================================================
    // ADMIN ONLY
    // =====================================================
    if (req.user.role?.toLowerCase() !== "admin") {
      return res.status(403).json({
        success: false,
        message: "Access denied",
      });
    }

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
app.post("/api/filters", authMiddleware, async (req, res) => {  
const { reportType } = req.body;
const userLevel = req.user.level;
const designation = req.user.designation;
const userBranchCode = req.user.branchCode;
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
app.post("/get-daily-summary", authMiddleware, async (req, res) => {
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
          FROM employees_master
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
app.get("/get-daily-summary-columns", authMiddleware, async (req, res) => {
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

app.post("/get-daily-summary-trend", authMiddleware, async (req, res) => {
  try {
const { filters = {} } = req.body;

// =====================================================
// JWT VALUES ONLY
// =====================================================
const userId = req.user.employeeId;
const level = req.user.level;
const designation = req.user.designation;
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
        FROM employees_master
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

app.post("/generate-loans-opened-summary", authMiddleware, async (req, res) => {
  try {

 let {
  branchCode,
  branchName,
  clusterName,
  districtName,
  page = 1,
  pageSize = 10,
  sortBy = null,
} = req.body;

const userId = req.user.employeeId;
const role = req.user.role;

console.log(
  "GENERATE_LOANS_OPENED_SUMMARY_API_HIT",
  {
    employeeId: userId,
    role,
    branchCode,
    branchName,
    clusterName,
    districtName,
    page,
    pageSize,
    sortBy
  }
);

// =====================================================
// JWT SCOPE ENFORCEMENT
// =====================================================
const userLevel = req.user.level;
const designation = req.user.designation;

if (userLevel === "Level 1") {
  branchCode = req.user.branchCode;

  console.log(
    "GENERATE_LOANS_OPENED_SUMMARY_LEVEL1_BRANCH_ENFORCED",
    {
      employeeId: userId,
      branchCode
    }
  );
}
if (userLevel === "Level 2") {
  const match = designation?.match(
    /Cluster\s*Head\s*-\s*(.*)/i
  );

if (match) {
  clusterName = match[1].trim();

  console.log(
    "GENERATE_LOANS_OPENED_SUMMARY_LEVEL2_CLUSTER_RESOLVED",
    {
      employeeId: userId,
      clusterName
    }
  );
}
}

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
console.log(
  "GENERATE_LOANS_OPENED_SUMMARY_MASTER_QUERY_STARTED",
  {
    employeeId: userId
  }
);
    const masterData = await queryUTIDatabase(masterQuery);

    if (!masterData || masterData.length === 0) {
console.warn(
  "GENERATE_LOANS_OPENED_SUMMARY_NO_BRANCHES_FOUND",
  {
    employeeId: userId,
    branchCode,
    branchName,
    clusterName,
    districtName
  }
);
await logActivity(
  userId,
  role,
  "View Loans Accounts Opened",
  "No branches found"
);
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
console.log(
  "GENERATE_LOANS_OPENED_SUMMARY_LOAN_QUERY_STARTED",
  {
    employeeId: userId
  }
);
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
console.log(
  "GENERATE_LOANS_OPENED_SUMMARY_SORT_APPLIED",
  {
    employeeId: userId,
    sortBy: sortBy || "Branch Code"
  }
);
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
await logActivity(
  userId,
  role,
  "View Loans Accounts Opened",
  `Records=${totalRecords}, Page=${page}`
);

console.log(
  "VIEW_LOANS_ACCOUNTS_OPENED_ACTIVITY_LOGGED",
  {
    employeeId: userId
  }
);
console.log(
  "GENERATE_LOANS_OPENED_SUMMARY_SUCCESS",
  {
    employeeId: userId,
    totalRecords,
    totalPages,
    page
  }
);
    return res.json({
      success: true,
      totalRecords,
      totalPages,
      loan_categories: [...loanCategories, "Total"],
      column_structure,
      data: pagedData,
    });

  } catch (err) {
   console.error(
  "GENERATE_LOANS_OPENED_SUMMARY_FAILED",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    page: req.body?.page,
    error: err.message,
    stack: err.stack
  }
);
    return res.status(500).json({
      success: false,
      message: "Server error generating summary",
      error: err.message,
    });
  }
});
//============================================================================================
//                                      BRANCH CONTACTS
//============================================================================================
app.post("/get-branch-contacts", authMiddleware, async (req, res) => {
  try {
    let { branchCodes = [], filters = {} } = req.body;

let {
  clusterName,
  branchCode,
  branchName,
  districtName,
} = filters;

    // ----------------------------------------------------
    // Normalize branchCodes → numeric only
    // ----------------------------------------------------
    if (!Array.isArray(branchCodes)) branchCodes = [branchCodes];

    branchCodes = branchCodes
      .map((c) => {
        if (!c && c !== 0) return null;

        if (typeof c === "object") {
          const keys = [
            "Branch Code",
            "BR Code",
            "Br Code",
            "branch_code",
            "Brcode",
          ];
          for (const key of keys) {
            if (c[key] !== undefined && c[key] !== null && c[key] !== "") {
              const n = parseInt(c[key], 10);
              return isNaN(n) ? null : n;
            }
          }
          return null;
        }

        const n = parseInt(c, 10);
        return isNaN(n) ? null : n;
      })
      .filter((v) => v !== null);
	  
// =====================================================
// JWT SCOPE ENFORCEMENT
// =====================================================
const userLevel = req.user.level;
const designation = req.user.designation;

console.log(
  "GET_BRANCH_CONTACTS_API_HIT",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    userLevel,
    branchCodes,
    filters
  }
);

// ----------------------------------------------------
// LEVEL 1
// ----------------------------------------------------
if (userLevel === "Level 1") {

    const authorizedBranch = Number(req.user.branchCode);

    const requestedBranches = Array.isArray(branchCodes)
        ? branchCodes.map(Number)
        : [Number(branchCodes)];

    const unauthorizedAttempt =
        requestedBranches.some(
            b => b !== authorizedBranch
        );

    if (unauthorizedAttempt) {

        await logActivity(
            req.user.employeeId,
            req.user.role,
            "Blocked Branch IDOR",
            `Requested=${requestedBranches.join(",")} Allowed=${authorizedBranch}`
        );

        return res.status(403).json({
            success: false,
            message: "Access denied"
        });
    }

    // Ignore client supplied values
    branchCodes = [authorizedBranch];
    branchCode = authorizedBranch;
}


// ----------------------------------------------------
// LEVEL 2
// ----------------------------------------------------
if (userLevel === "Level 2") {

    // Extract cluster from designation
    const match =
        designation?.match(
            /Cluster\s*Head\s*-\s*(.*)/i
        );

    if (!match) {

        console.warn(
            "LEVEL2_CLUSTER_NOT_FOUND",
            {
                employeeId: req.user.employeeId,
                designation
            }
        );

        return res.status(403).json({
            success: false,
            message: "Cluster mapping not found"
        });
    }

    const allowedCluster = match[1].trim();
	
	if (!allowedCluster) {
    return res.status(403).json({
        success:false,
        message:"Cluster mapping not found"
    });
}

    // Never trust request cluster
    clusterName = allowedCluster;

    // Fetch branches belonging to cluster
    const allowedBranches = await queryUTIDatabase(
        `
        SELECT branch_code
        FROM MIS.dbo.Branch_Cluster_Master
        WHERE cluster_name = ?
        `,
        [allowedCluster]
    );

    const allowedSet = new Set(
        allowedBranches.map(
            x => Number(x.branch_code)
        )
    );
// Don't allow branch_name-only requests
if (
    branch_name &&
    (
      branch_code === null ||
      branch_code === undefined ||
      branch_code === ""
    )
) {

    await logActivity(
        req.user.employeeId,
        req.user.role,
        "Blocked Glance IDOR",
        `BranchName=${branch_name}`
    );

    return res.status(403).json({
        success:false,
        message:"Access denied"
    });
}
    // Validate branchCodes[]
    if (branchCodes.length) {

        const unauthorized =
            branchCodes.some(
                b => !allowedSet.has(Number(b))
            );

        if (unauthorized) {

            await logActivity(
                req.user.employeeId,
                req.user.role,
                "Blocked Branch IDOR",
                `Requested=${branchCodes.join(",")}`
            );

            return res.status(403).json({
                success: false,
                message: "Access denied"
            });
        }
    }

    // Validate branchCode filter
    if (
        branchCode &&
        !allowedSet.has(Number(branchCode))
    ) {

        await logActivity(
            req.user.employeeId,
            req.user.role,
            "Blocked BranchCode IDOR",
            `Requested=${branchCode}`
        );

        return res.status(403).json({
            success: false,
            message: "Access denied"
        });
    }
}


    // ----------------------------------------------------
    // BASE QUERY (MASTER TABLE IS AUTHORITATIVE)
    // ----------------------------------------------------
    let query = `
      SELECT
        bc.[Br Code]         AS branch_code,
        bcm.cluster_name    AS cluster,
        bc.[Branch Name]    AS branch_name,
        bc.[Branch Manager] AS manager,
        bc.[Contact Number] AS contact_number
      FROM [MIS].[dbo].[branchContacts] bc
      INNER JOIN [MIS].[dbo].[Branch_Cluster_Master] bcm
        ON TRY_CAST(LTRIM(RTRIM(bc.[Br Code])) AS INT)
         = TRY_CAST(LTRIM(RTRIM(bcm.branch_code)) AS INT)
      WHERE 1 = 1
    `;

    const params = [];

    // ----------------------------------------------------
    // 🔐 FILTER PRIORITY (VERY IMPORTANT)
    // ----------------------------------------------------

    // 1️⃣ Cluster filter → HIGHEST PRIORITY
  if (clusterName && !branchCodes.length) {
      query += ` AND LTRIM(RTRIM(bcm.cluster_name)) = ?`;
      params.push(clusterName.trim());
    }
    // 2️⃣ Only if cluster NOT provided → use branchCodes
    else if (branchCodes.length) {
      const placeholders = branchCodes.map(() => "?").join(",");
      query += `
        AND TRY_CAST(LTRIM(RTRIM(bc.[Br Code])) AS INT)
            IN (${placeholders})
      `;
      params.push(...branchCodes);
    }
    // 3️⃣ Safety fallback
    else {
console.warn(
  "GET_BRANCH_CONTACTS_NO_FILTERS",
  {
    employeeId: req.user?.employeeId
  }
);
      return res.json({ data: [] });
    }

    // ----------------------------------------------------
    // ADDITIONAL FILTERS
    // ----------------------------------------------------
    if (branchCode) {
      query += `
        AND TRY_CAST(LTRIM(RTRIM(bc.[Br Code])) AS INT)
            = TRY_CAST(? AS INT)
      `;
      params.push(branchCode);
    }

    if (branchName) {
      query += `
        AND LOWER(LTRIM(RTRIM(bc.[Branch Name]))) LIKE ?
      `;
      params.push(`%${branchName.trim().toLowerCase()}%`);
    }

    if (districtName) {
      query += `
        AND LOWER(LTRIM(RTRIM(bcm.district))) LIKE ?
      `;
      params.push(`%${districtName.trim().toLowerCase()}%`);
    }

    // ----------------------------------------------------
    // ORDERING (NUMERIC SAFE)
    // ----------------------------------------------------
    query += `
      ORDER BY TRY_CAST(LTRIM(RTRIM(bc.[Br Code])) AS INT) ASC
    `;
	
    // ----------------------------------------------------
    // EXECUTE
    // ----------------------------------------------------
    let results = [];
    try {
console.log(
  "GET_BRANCH_CONTACTS_QUERY_STARTED",
  {
    employeeId: req.user?.employeeId,
    paramsCount: params.length,
    queryLength: query.length
  }
);
      results = await queryUTIDatabase(query, params);
console.log(
  "GET_BRANCH_CONTACTS_QUERY_SUCCESS",
  {
    employeeId: req.user?.employeeId,
    records: results.length
  }
);
    } catch (dbErr) {
     console.error(
  "GET_BRANCH_CONTACTS_DB_FAILED",
  {
    employeeId: req.user?.employeeId,
    error: dbErr.message
  }
);
      return res.status(500).json({ error: "Database query failed" });
    }
await logActivity(
  req.user.employeeId,
  req.user.role,
  "View Branch Contacts",
  `Records=${results.length}`
);

console.log(
  "VIEW_BRANCH_CONTACTS_ACTIVITY_LOGGED",
  {
    employeeId: req.user?.employeeId
  }
);
console.log(
  "GET_BRANCH_CONTACTS_SUCCESS",
  {
    employeeId: req.user?.employeeId,
    records: results.length
  }
);
    // ----------------------------------------------------
    // RETURN RAW DATA
    // ----------------------------------------------------
    return res.json({ data: results });

  } catch (err) {
   console.error(
  "GET_BRANCH_CONTACTS_FAILED",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    error: err.message
  }
);
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
  "deposits_all": "DOA_Yearly",
  "rbia audit": "RBIA_Audit_Report",
};

// CLUSTER SECTION → TABLE
const clusterTableMap = {
  deposits: "DCS",
  advances: "ACS",
  "accounts opened": "DAOCS",
  npa: "NPACS",
  loans: "LSCS",
  sma: "SMACS"
};

const fixedTableUploadMap = {
  "monthly loan summary": "Monthly_Loan_Summary_Stage",
  "yearly loan summary": "Yearly_Loan_Summary_Stage",
  "deposits_all": "DOA_Yearly",
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
  "FY SMA 31-03-2025 A/Cs":"sma_fy_accounts",
  "FY SMA 31-03-2025 Balance":"sma_fy_balance",
};

// -------- DEPOSITS ACCOUNTS OPENED --------
const DAO_MEANING_OVERRIDES = {
  "TDR (>=1LAKH ) A/Cs": "tdr_accounts",
  "RD A/Cs": "rd_accounts",
  "Savings A/Cs": "savings_accounts",
  "TDR (>=1LAKH ) Amount": "tdr_amount",
  "RD Amount": "rd_amount",
  "Savings Amount": "savings_amount",
  "Current A/Cs": "current_accounts",
  "Current Amount": "current_amount",
  Vouchers: "vouchers",
  "Vouchers Average": "vouchers_average",  
};

const DAOCS_MEANING_OVERRIDES = {
  "TDR(>=1 LAKH) A/Cs": "tdr_accounts",
  "TDR(>=1 LAKH) Amount": "tdr_amount",
  "RD A/Cs": "rd_accounts",
  "RD Amount": "rd_amount",
  "Savings A/Cs": "savings_accounts",
  "Savings Amount": "savings_amount",
  "Current A/Cs": "current_accounts",
  "Current Amount": "current_amount",
  Vouchers: "vouchers",
  "Vouchers Average": "vouchers_average",  
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
  "Savings A/Cs": "savings_accounts",
  "Current A/Cs": "current_accounts",
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

    "On _17-05-2026  A/Cs": "on_date_accounts",
    "On _17-05-2026  Amount": "on_date_amount",

    "During May-2026 A/Cs": "during_month_accounts",
    "During May-2026 Amount": "during_month_amount"
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
// ⭐ FY SNAPSHOT — As on 31-03-YYYY (Bank + Cluster + History)
// -------------------------------------------------------------------------
const lower = trimmed.toLowerCase();

// FY snapshot column
if (lower.startsWith("as on")) {

    const isFYEnd =
        /\b31[-\/]03[-\/]\d{4}\b/.test(lower) ||
        /\b31[-\/]3[-\/]\d{4}\b/.test(lower);

    if (isFYEnd) {

        if (
            tableName === "Deposits" ||
            tableName === "Deposits_history" ||
            tableName === "DCS"
        ) {
            return "as_on_fy_deposits";
        }

        if (
            tableName === "Advances" ||
            tableName === "Advances_history" ||
            tableName === "ACS"
        ) {
            return "as_on_fy_advances";
        }
    }
}

// ----------------------------------------------------
// ADVANCES MONTHLY SNAPSHOT
// ----------------------------------------------------
if (
    tableName === "Advances" ||
    tableName === "Advances_history" ||
    tableName === "ACS"
) {

    const hasDate =
        /\b\d{1,2}[-\/]\d{1,2}[-\/]\d{4}\b/.test(lower);

    const isFYEnd =
        /\b31[-\/]03[-\/]\d{4}\b/.test(lower);

    if (
        lower.startsWith("as on") &&
        hasDate &&
        !isFYEnd
    ) {
        return "as_on_date_advances";
    }

    if (lower === "amount") {
        return "amount_advances";
    }

    if (lower.includes("cy budgeted growth")) {
        return "cy_budgeted_growth_advances";
    }
}
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
// ⭐ SMA — MONTHLY SNAPSHOT (AS ON <DATE>)
// -------------------------------------------------------------------------
if (tableName === "SMA") {
  const lower = trimmed.toLowerCase().replace(/\s+/g, " ").trim();

  // detect any date like 31-12-2025 / 31/12/2025 / 31 12 2025
  const hasDate =
    /\b\d{1,2}[-\/\s\.]\d{1,2}[-\/\s\.]\d{2,4}\b/.test(lower);

  // must start with "as on"
  if (hasDate && lower.startsWith("as on")) {
    if (lower.includes("a/c")) {
      return "sma_monthly_accounts";
    }
    if (lower.includes("amount") || lower.includes("balance")) {
      return "sma_monthly_balance";
    }
  }
}

  // ⭐ SMA FY COLUMNS (Year-end snapshot)
if (tableName === "SMA") {
  const lower = trimmed.toLowerCase().replace(/\s+/g, " ").trim();

  // detect FY date like 31-03-2025
  const isFYDate =
    lower.includes("fy") ||
    lower.includes("31-03") ||
    lower.includes("31/03");

  if (isFYDate) {
    if (lower.includes("a/c")) return "sma_fy_accounts";
    if (lower.includes("balance") || lower.includes("amount"))
      return "sma_fy_balance";
  }
}

if (tableName === "Deposits" && trimmed === "Branch Code") {
  return "branch_code";
}

if (tableName === "Advances" && trimmed === "Branch Code") {
  return "branch_code";
}
// -------------------------------------------------------------------------
// ⭐ SMACS — COMPLETE MEANING RESOLUTION (CLUSTER SMA SUMMARY)
// -------------------------------------------------------------------------
if (tableName === "SMACS") {
  const lower = trimmed.toLowerCase().replace(/\s+/g, " ").trim();

  // -------------------------------
  // 1️⃣ BASIC IDENTIFIERS
  // -------------------------------
  if (lower === "cluster") return "cluster";
  if (lower === "district") return "district";

  // -------------------------------
  // 2️⃣ SMA CATEGORY COUNTS
  // -------------------------------
  if (lower === "sma 0 a/cs") return "sma0_accounts";
  if (lower === "sma 0 amount") return "sma0_amount";

  if (lower === "sma 1 a/cs") return "sma1_accounts";
  if (lower === "sma 1 amount") return "sma1_amount";

  if (lower === "sma 2 a/cs") return "sma2_accounts";
  if (lower === "sma 2 amount") return "sma2_amount";

  // -------------------------------
  // 3️⃣ TOTALS
  // -------------------------------
  if (lower === "total_ac" || lower === "total a/cs")
    return "total_accounts";

  if (lower === "total_balance" || lower === "total amount")
    return "total_balance";

  // -------------------------------
  // 4️⃣ FY SNAPSHOT (31-03-YYYY)
  // -------------------------------
  const isFY =
    lower.includes("fy") ||
    lower.includes("31-03") ||
    lower.includes("31/03");

  if (isFY) {
    if (lower.includes("a/c")) return "sma_fy_accounts";
    if (lower.includes("balance") || lower.includes("amount"))
      return "sma_fy_balance";
  }

  // -------------------------------
  // 5️⃣ MONTHLY SNAPSHOT (AS ON <DATE>)
  // -------------------------------
  const hasDate =
    /\b\d{1,2}[-\/\s\.]\d{1,2}[-\/\s\.]\d{2,4}\b/.test(lower);

  if (hasDate && lower.startsWith("as on")) {
    if (lower.includes("a/c")) return "sma_monthly_accounts";

    if (
      lower.includes("amount") ||
      lower.includes("amonut") || // typo-safe
      lower.includes("balance")
    )
      return "sma_monthly_balance";
  }

  // -------------------------------
  // 6️⃣ NO MATCH
  // -------------------------------
  return null;
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
// ⭐ Loans Sanctioned (Bank + Cluster)
if (
    tableName === "LSCS" ||
    tableName === "loanssanctioned"
) {

    const lower = trimmed
        .toLowerCase()
        .replace(/\s+/g, " ")
        .trim();

    const hasDate =
        /_?\d{1,2}[-\/\.]\d{1,2}[-\/\.]\d{2,4}/.test(lower);

    const hasMonthYear =
        /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*[- ]?\d{2,4}\b/
        .test(lower);

    // On <date>
    if (
        (lower.startsWith("on") || lower.startsWith("on _")) &&
        (hasDate || hasMonthYear)
    ) {

        if (lower.includes("a/c"))
            return "on_date_accounts";

        if (lower.includes("amount"))
            return "on_date_amount";
    }

    // During month
    if (
        lower.startsWith("during") &&
        hasMonthYear
    ) {

        if (lower.includes("a/c"))
            return "during_month_accounts";

        if (lower.includes("amount"))
            return "during_month_amount";
    }

    // FY
    if (lower.includes("over all fy")) {

        if (lower.includes("a/c"))
            return "over_all_fy_accounts";

        if (lower.includes("amount"))
            return "over_all_fy_amount";
    }
}
// -------------------------------------------------------------------------
// ⭐ CLUSTER SUMMARY — DCS (Deposits Cluster Summary)
// -------------------------------------------------------------------------
if (tableName === "DCS") {
  const dcsMap = {
    "As on Date": "as_on_date",
    "Change over PD": "change_over_pd_deposits",
    "GDM": "gdm_deposits",
    "GUM": "gum_deposits",
    "Budget Achieved %": "budget_achieved_deposits",
  };

  if (dcsMap[trimmed]) {
    return dcsMap[trimmed];
  }
}

// -------------------------------------------------------------------------
// ⭐ CLUSTER SUMMARY — ACS (Advances Cluster Summary)
// -------------------------------------------------------------------------
if (tableName === "ACS") {
  const acsMap = {
    "As on Date": "as_on_date",
    "Change over PD": "change_over_pd_advances",
    "GDM": "gdm_advances",
    "GUM": "gum_advances",
    "Budget Achieved %": "budget_achieved_advances",
  };

  if (acsMap[trimmed]) {
    return acsMap[trimmed];
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
// ⭐ 1️⃣ MONTHLY SNAPSHOT COLUMNS (AS ON <DATE>)
// -------------------------------
if (hasDate && lower.startsWith("as on")) {
  if (lower.includes("a/c")) {
    return "npa_monthly_accounts";
  }
  if (lower.includes("balance")) {
    return "npa_monthly_balance";
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

// ⭐ NPA — DATE-BASED COLUMNS
if (tableName === "NPA") {
  const lower = trimmed.toLowerCase().replace(/\s+/g, " ").trim();
  const hasDate = /\b\d{1,2}[-\/\s\.]\d{1,2}[-\/\s\.]\d{2,4}\b/.test(lower);

  // YEAR-END (31/03/yyyy)
  if (hasDate && lower.includes("31/03")) {
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

  // MONTHLY SNAPSHOT
  if (hasDate && lower.startsWith("as on")) {
    if (lower.includes("a/c")) {
      return "npa_monthly_accounts";
    }
    if (lower.includes("balance")) {
      return "npa_monthly_balance";
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
const storage = multer.memoryStorage();

// 👇 ADD HERE
const allowedExtensions = [".csv", ".xlsx"];

const ALLOWED_MIME_TYPES = [
  "text/csv",
  "text/comma-separated-values",
  "application/csv",
  "text/plain",
  "application/vnd.ms-excel",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
];

const upload = multer({
  storage,

  limits: {
    fileSize: 10 * 1024 * 1024,
  },

fileFilter: (req, file, cb) => {

  console.log("=================================");
  console.log("FILE NAME =", file.originalname);
  console.log("FILE MIME =", file.mimetype);
  console.log("=================================");

  const ext = path.extname(file.originalname).toLowerCase();

  if (!allowedExtensions.includes(ext)) {
    return cb(new Error("Only CSV and XLSX files are allowed"));
  }

  if (!ALLOWED_MIME_TYPES.includes(file.mimetype)) {
    return cb(new Error("Invalid file MIME type"));
  }

  cb(null, true);
},
});

const UPLOAD_SIGNATURES = {

 Deposits:{
  meanings:[
    "as_on_date_deposits",
    "change_over_pd_deposits",
    "gdm_deposits",
    "gum_deposits"
  ],
  identifiers:[
    "branch_code",
    "branch_name"
  ]
},
Advances:{
  meanings:[
    "as_on_date_advances",
    "amount_advances",
    "change_over_pd_advances",
    "gdm_advances",
    "gum_advances"
  ],
  identifiers:[
    "branch_code",
    "branch_name"
  ]
},
  deposits_accounts_opened: {
    meanings: ["tdr_accounts", "rd_accounts"],
    identifiers: ["branch_code", "branch_name"]
  },

  NPA: {
    meanings: ["stamped_npa_accounts"],
    identifiers: ["branch_code", "branch_name"]
  },

  SMA: {
    meanings: ["sma0_accounts", "sma1_accounts"],
    identifiers: ["branch_code", "branch_name"]
  },

  loanssanctioned: {
    meanings: ["over_all_fy_accounts"],
    identifiers: ["branch_code", "branch_name"]
  },

 DCS:{
  meanings:[
    "as_on_date",
    "change_over_pd_deposits",
    "gdm_deposits",
    "gum_deposits"
  ],
  identifiers:["cluster"]
},

 ACS:{
  meanings:[
    "as_on_date",
    "change_over_pd_advances",
    "gdm_advances",
    "gum_advances"
  ],
  identifiers:["cluster"]
},
  DAOCS: {
    meanings: ["tdr_accounts", "rd_accounts"],
    identifiers: ["cluster"]
  },

  NPACS: {
    meanings: ["stamped_npa_accounts"],
    identifiers: ["cluster"]
  },

  SMACS: {
    meanings: ["sma0_accounts", "sma1_accounts"],
    identifiers: ["cluster"]
  },

  LSCS: {
    meanings: ["over_all_fy_accounts"],
    identifiers: ["cluster"]
  }
};

// ============================================================================================
// MAIN UPLOAD ROUTE
// ============================================================================================
app.post(
"/upload-mis",
authMiddleware,
uploadLimiter,
upload.single("file"),
async (req,res)=>{

  const { section, clusterSection } = req.body;

  const userId = req.user.employeeId;
  const role = req.user.role;

console.log(
  "UPLOAD_MIS_API_HIT",
  {
    employeeId: userId,
    role,
    section,
    clusterSection,
    fileName: req.file?.originalname,
    fileSize: req.file?.size
  }
);

  // =====================================================
  // ROLE CHECK
  // =====================================================
  if (req.user.role?.toLowerCase() !== "admin") {
console.warn(
  "UPLOAD_MIS_ACCESS_DENIED",
  {
    employeeId: userId,
    role
  }
);

await logSecurityEvent(
  userId,
  "MIS Upload Access Denied",
  JSON.stringify({
    role
  })
);
    return res.status(403).json({
      success: false,
      message: "Access denied",
    });
  }

if (!req.file) {
  console.warn(
    "UPLOAD_MIS_NO_FILE",
    {
      employeeId: userId
    }
  );

  return res.status(400).json({
    message: "No file uploaded"
  });
}

// =====================================================
// FILE SIGNATURE VALIDATION
// =====================================================
const fileBuffer = req.file.buffer;

const detectedType = await fileTypeFromBuffer(fileBuffer);

console.log(
  "UPLOAD_MIS_FILE_VALIDATED",
  {
    employeeId: userId,
    fileName: req.file.originalname,
    fileSize: req.file.size,
    detectedMime: detectedType?.mime || "csv"
  }
);

// XLSX validation
if (
  req.file.originalname.toLowerCase().endsWith(".xlsx")
) {
  if (
    !detectedType ||
    detectedType.mime !==
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  ) {
   
console.warn(
  "UPLOAD_MIS_INVALID_XLSX_SIGNATURE",
  {
    employeeId: userId,
    fileName: req.file.originalname
  }
);

await logSecurityEvent(
  userId,
  "Invalid XLSX Upload",
  req.file.originalname
);
    return res.status(400).json({
      success: false,
      message: "Invalid XLSX file signature",
    });
  }
}

// CSV validation
if (
  req.file.originalname.toLowerCase().endsWith(".csv")
) {
  const textContent = fileBuffer
    .toString("utf8")
    .slice(0, 1000);

  const printableChars =
    textContent.match(/[\x20-\x7E\r\n\t]/g)?.length || 0;

  const ratio =
    printableChars / Math.max(textContent.length, 1);

  if (ratio < 0.8) {
console.warn(
  "UPLOAD_MIS_INVALID_CSV_SIGNATURE",
  {
    employeeId: userId,
    fileName: req.file.originalname
  }
);

await logSecurityEvent(
  userId,
  "Invalid CSV Upload",
  req.file.originalname
);
    return res.status(400).json({
      success: false,
      message: "Invalid CSV file signature",
    });
  }
}

  let tableName;
const normalizedSection = section?.toLowerCase().trim();

// 1️⃣ FIXED STRUCTURE TABLES (Monthly / Yearly Loan Summary)
if (normalizedSection && fixedTableUploadMap[normalizedSection]) {
  tableName = fixedTableUploadMap[normalizedSection];
}

// 2️⃣ NORMAL MIS SECTIONS
else if (normalizedSection && sectionTableMap[normalizedSection]) {
  tableName = sectionTableMap[normalizedSection];
}

// 3️⃣ CLUSTER SECTIONS
else if (clusterSection) {
  const normalizedCluster = clusterSection.toLowerCase().trim();
  tableName = clusterTableMap[normalizedCluster];
}

// 4️⃣ FINAL VALIDATION
if (!tableName) {
console.warn(
  "UPLOAD_MIS_INVALID_SECTION",
  {
    employeeId: userId,
    section,
    clusterSection
  }
);
  return res.status(400).json({ message: "Invalid section" });
}


  const historyTable = `${tableName}_history`;
// ==========================================================================================
// ⭐ FIXED STRUCTURE LOAN SUMMARY TABLES (DIRECT COLUMN INSERT)
// ==========================================================================================
if (fixedTableUploadMap[section?.toLowerCase()]) {
  const tableName = fixedTableUploadMap[section.toLowerCase()];

  try {
    const ext = req.file.originalname.toLowerCase();
  const fileContent = req.file.buffer;

    const workbook = XLSX.read(
      ext.endsWith(".csv") ? fileContent.toString("utf8") : fileContent,
      { type: ext.endsWith(".csv") ? "string" : "buffer" }
    );

    const rows = XLSX.utils.sheet_to_json(
      workbook.Sheets[workbook.SheetNames[0]]
    );

    if (!rows.length) {
      return res.status(400).json({ message: "File empty" });
    }

    const fileCols = Object.keys(rows[0]); // EXACT headers
	
	

    // 1️⃣ Clear stage table (snapshot table)
    await queryUTIDatabase(`DELETE FROM [${tableName}]`);

    // 2️⃣ Insert rows directly
    for (const row of rows) {
      const cols = fileCols.map(c => `[${c}]`).join(", ");
      const placeholders = fileCols.map(() => "?").join(", ");
      const values = fileCols.map(c => row[c] ?? null);

      await queryUTIDatabase(
        `INSERT INTO [${tableName}] (${cols}) VALUES (${placeholders})`,
        values
      );
    }

    await logActivity(
      userId,
      role,
      "Upload MIS Report",
      `Uploaded ${rows.length} rows → ${tableName}`
    );
console.log(
  "UPLOAD_MIS_FIXED_TABLE_SUCCESS",
  {
    employeeId: userId,
    tableName,
    uploadedRows: rows.length
  }
);
    return res.json({
      success: true,
      message: `${tableName} uploaded successfully`,
      uploaded: rows.length,
    });
  } catch (err) {
  console.error(
  "UPLOAD_FIXED_TABLE_FAILED",
  {
    employeeId: userId,
    tableName,
    error: err.message
  }
);

    return res.status(500).json({
      message: "Error uploading loan summary",
      error: err.message,
    });
  }
}

  // ==========================================================================================
  // SPECIAL CASE: LoansAccountsOpened (real columns, NOT col1..col30)
  // ==========================================================================================
  if (tableName === "LoansAccountsOpened") {
    try {
      const ext = req.file.originalname.toLowerCase();
      const fileContent = req.file.buffer;
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
console.log(
  "UPLOAD_LOANS_ACCOUNTS_OPENED_SUCCESS",
  {
    employeeId: userId,
    uploadedRows: rows.length
  }
);
      return res.json({
        success: true,
        message:
          "LoansAccountsOpened uploaded with real columns + sortable meanings",
      });
    } catch (err) {
  console.error(
  "UPLOAD_LOANS_ACCOUNTS_OPENED_FAILED",
  {
    employeeId: userId,
    error: err.message
  }
);

      return res.status(500).json({
        message: "Error uploading LoansAccountsOpened",
        error: err.message,
      });
    }
  }
// ==========================================================================================
// DOA YEARLY
// ==========================================================================================
if (tableName === "DOA_Yearly") {

  try {

    const ext = req.file.originalname.toLowerCase();

    const fileContent = req.file.buffer;

    const workbook = XLSX.read(
      ext.endsWith(".csv")
        ? fileContent.toString("utf8")
        : fileContent,
      {
        type: ext.endsWith(".csv")
          ? "string"
          : "buffer"
      }
    );

    const rows = XLSX.utils.sheet_to_json(
      workbook.Sheets[workbook.SheetNames[0]]
    );

    if (!rows.length) {
      return res.status(400).json({
        message: "File empty"
      });
    }

    await queryUTIDatabase(`
      DELETE FROM MIS.dbo.DOA_Yearly
    `);

    for (const row of rows) {

      await queryUTIDatabase(`
      INSERT INTO MIS.dbo.DOA_Yearly
      (
        [Branch Code],
        [Branch Name],
        [District],
        [Cluster],

        [TDR_A/Cs_Month],
        [TDR_Bal_Month],

        [RD_A/Cs_Month],
        [RD_Bal_Month],

        [SB_A/Cs_Month],
        [SB_Bal_Month],

        [CA_A/Cs_Month],
        [CA_Bal_Month],

        [TDR_A/Cs_FY],
        [TDR_Bal_FY],

        [RD_A/Cs_FY],
        [RD_Bal_FY],

        [SB_A/Cs_FY],
        [SB_Bal_FY],

        [CA_A/Cs_FY],
        [CA_Bal_FY]
      )
      VALUES
      (
        ?,?,?,?,?,?,
        ?,?,?,?,?,?,
        ?,?,?,?,?,?,
        ?,?
      )
      `,
      [
        row["Branch Code"],
        row["Branch Name"],
        row["District"],
        row["Cluster"],

        row["TDR_A/Cs_Month"],
        row["TDR_Bal_Month"],

        row["RD_A/Cs_Month"],
        row["RD_Bal_Month"],

        row["SB_A/Cs_Month"],
        row["SB_Bal_Month"],

        row["CA_A/Cs_Month"],
        row["CA_Bal_Month"],

        row["TDR_A/Cs_FY"],
        row["TDR_Bal_FY"],

        row["RD_A/Cs_FY"],
        row["RD_Bal_FY"],

        row["SB_A/Cs_FY"],
        row["SB_Bal_FY"],

        row["CA_A/Cs_FY"],
        row["CA_Bal_FY"]
      ]);
    }
console.log(
  "UPLOAD_DOA_YEARLY_SUCCESS",
  {
    employeeId: userId,
    uploadedRows: rows.length
  }
);
    return res.json({
      success: true,
      message: `DOA_Yearly uploaded successfully`,
      uploaded: rows.length
    });

  } catch (err) {
console.error(
  "UPLOAD_DOA_YEARLY_FAILED",
  {
    employeeId: userId,
    error: err.message
  }
);
    return res.status(500).json({
      message: "DOA upload failed",
      error: err.message
    });
  }
}

// ==========================================================================================
// RBIA AUDIT REPORT
// ==========================================================================================
// ==========================================================================================
// RBIA AUDIT REPORT
// ==========================================================================================
if (tableName === "RBIA_Audit_Report") {

  try {

    const ext = req.file.originalname.toLowerCase();

    const fileContent = req.file.buffer;

    const workbook = XLSX.read(
      ext.endsWith(".csv")
        ? fileContent.toString("utf8")
        : fileContent,
      {
        type: ext.endsWith(".csv")
          ? "string"
          : "buffer"
      }
    );

    const rows = XLSX.utils.sheet_to_json(
      workbook.Sheets[workbook.SheetNames[0]]
    );

    if (!rows.length) {
      return res.status(400).json({
        success: false,
        message: "File empty"
      });
    }

    // =====================================================
    // VALIDATE REQUIRED HEADERS
    // =====================================================

    const requiredColumns = [
      "Name of the branch",
      "Branch Code",
      "Date of Audit",
      "Business Risk",
      "Credit Risk",
      "Operational Risk",
      "Overall Score",
      "Rating"
    ];

    const fileColumns = Object.keys(rows[0]).map(column =>
      column
        .replace(/\r?\n/g, " ")
        .replace(/\s+/g, " ")
        .trim()
    );

    const missingColumns = requiredColumns.filter(
      column => !fileColumns.includes(column)
    );

    if (missingColumns.length > 0) {
      return res.status(400).json({
        success: false,
        message: `Wrong RBIA Audit file. Missing columns: ${missingColumns.join(", ")}`
      });
    }

    // =====================================================
    // CLEAR OLD DATA
    // =====================================================

    await queryUTIDatabase(`
      DELETE FROM MIS.dbo.RBIA_Audit_Report
    `);

    // =====================================================
    // INSERT NEW DATA
    // =====================================================

    for (const row of rows) {

      const normalizedRow = {};

      Object.keys(row).forEach(key => {
        normalizedRow[
          key
            .replace(/\r?\n/g, " ")
            .replace(/\s+/g, " ")
            .trim()
        ] = row[key];
      });

      // ===============================================
      // FORMAT DATE
      // Supports:
      // 04.08.2025
      // Excel Date object
      // ===============================================

      let auditDate = normalizedRow["Date of Audit"];

      if (auditDate instanceof Date) {

        auditDate = auditDate.toISOString().split("T")[0];

      } else if (typeof auditDate === "string") {

        const parts = auditDate.split(".");

        if (parts.length === 3) {
          auditDate = `${parts[2]}-${parts[1]}-${parts[0]}`;
        }

      }

      await queryUTIDatabase(
        `
        INSERT INTO MIS.dbo.RBIA_Audit_Report
        (
          [Name of the branch],
          [Branch Code],
          [Date of Audit],
          [Business Risk],
          [Credit Risk],
          [Operational Risk],
          [Overall Score],
          [Rating]
        )
        VALUES
        (
          ?, ?, ?, ?, ?, ?, ?, ?
        )
        `,
        [
          normalizedRow["Name of the branch"],
          normalizedRow["Branch Code"],
          auditDate,
          normalizedRow["Business Risk"],
          normalizedRow["Credit Risk"],
          normalizedRow["Operational Risk"],
          normalizedRow["Overall Score"],
          normalizedRow["Rating"]
        ]
      );

    }

    // =====================================================
    // SUCCESS LOGS
    // =====================================================

    await logActivity(
      userId,
      role,
      "Upload MIS Report",
      `Uploaded ${rows.length} rows → RBIA_Audit_Report`
    );

    console.log(
      "UPLOAD_RBIA_AUDIT_SUCCESS",
      {
        employeeId: userId,
        uploadedRows: rows.length
      }
    );

    return res.json({
      success: true,
      message: "RBIA Audit uploaded successfully",
      uploaded: rows.length
    });

  } catch (err) {

    console.error(
      "UPLOAD_RBIA_AUDIT_FAILED",
      {
        employeeId: userId,
        error: err.message
      }
    );

    return res.status(500).json({
      success: false,
      message: "RBIA Audit upload failed",
      error: err.message
    });

  }

}
  // ==========================================================================================
  // DEFAULT MODE: col1..col30 (Deposits, Advances, NPA, SMA, deposits_accounts_opened, etc.)
  // ==========================================================================================
  try {
    const ext = req.file.originalname.toLowerCase();
    const fileContent = req.file.buffer;
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
  
  if (
    section === "deposits" &&
    req.file.originalname.toLowerCase().includes("adv")
) {
    return res.status(400).json({
        success:false,
        message:"Advances file selected under Deposits"
    });
}

if (
    section === "advances" &&
    req.file.originalname.toLowerCase().includes("dp")
) {
    return res.status(400).json({
        success:false,
        message:"Deposits file selected under Advances"
    });
}

// ==========================================
// FILE VALIDATION
// ==========================================

const fileMeanings = [];

for (const col of fileCols) {
	
	console.log(
  "HEADER:",
  col,
  "=>",
  resolveFinalMeaning(col.trim(), tableName)
);

  const meaning = resolveFinalMeaning(
    col.trim(),
    tableName
  );

  if (meaning) {
    fileMeanings.push(meaning);
  }
}

console.log("FILE_MEANINGS =", fileMeanings);

const signature = UPLOAD_SIGNATURES[tableName];

if (signature) {

  const meaningsOk =
    signature.meanings.every(
      m => fileMeanings.includes(m)
    );

  const identifiersOk =
    signature.identifiers.every(
      m => fileMeanings.includes(m)
    );

  if (!meaningsOk || !identifiersOk) {

    return res.status(400).json({
      success: false,
      message: `Wrong file selected for ${tableName}`
    });

  }
}
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
const sortableMeaning = resolveFinalMeaning(label, tableName);

if (sortableMeaning) {

  sortablePromises.push(
    queryUTIDatabase(
      `
      INSERT INTO MIS.dbo.MIS_Sortable_Columns
      (
        section,
        display_label,
        meaning
      )
      VALUES (?, ?, ?)
      `,
      [
        tableName,
        label,
        sortableMeaning
      ]
    )
  );

}
      colIndex++;
    }

    await Promise.all([...mappingPromises, ...sortablePromises]);

    // Prepare rows in col1..col30 order
const dbRows = rows.map((r) => {
  const ordered = fileCols.map((c) => {
    let val = r[c];

    // convert dash / blank to NULL
    if (
      val === "-" ||
      val === "" ||
      val === undefined ||
      val === null ||
      val === " - "
    ) {
      return null;
    }

    // remove commas for decimal columns
    if (typeof val === "string") {
      val = val.replace(/,/g, "").trim();
    }

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
console.log(
  "UPLOAD_MIS_ACTIVITY_LOGGED",
  {
    employeeId: userId,
    tableName,
    uploadedRows: dbRows.length,
    fileName: req.file.originalname
  }
);
console.log(
  "UPLOAD_MIS_SUCCESS",
  {
    employeeId: userId,
    tableName,
    uploadedRows: dbRows.length
  }
);
    return res.json({
      success: true,
      message: "MIS uploaded successfully",
      uploaded: dbRows.length,
    });
  } catch (err) {
console.error(
  "UPLOAD_MIS_FAILED",
  {
    employeeId: userId,
    tableName,
    error: err.message
  }
);
  console.error(err);

return res.status(500).json({
  success:false,
  message:"Internal server error"
});
  }
});

// ============================================================================================
// PART D — UNIVERSAL SORTABLE COLUMNS ENDPOINT
// ============================================================================================
app.post("/get-sortable-columns", authMiddleware, async (req, res) => {
  try {
const { section } = req.body;

console.log(
  "GET_SORTABLE_COLUMNS_API_HIT",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    section
  }
);

if (!section) {

  console.warn(
    "GET_SORTABLE_COLUMNS_SECTION_MISSING",
    {
      employeeId: req.user?.employeeId,
      role: req.user?.role
    }
  );

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

let filteredRows = rows || [];

// ----------------------
// SMA
// ----------------------
if (tableName === "SMA") {
  filteredRows = filteredRows.filter(r =>
    [
      "sma0_accounts",
      "sma1_accounts",
      "sma2_accounts"
    ].includes(r.meaning)
  );
}

// ----------------------
// Deposits
// ----------------------
if (tableName === "Deposits") {
  filteredRows = filteredRows.filter(r =>
    [
      "change_over_pd_deposits",
      "gdm_deposits",
      "gum_deposits",
      "budget_achieved_deposits"
    ].includes(r.meaning)
  );
}

// ----------------------
// Deposits Accounts Opened
// ----------------------
if (tableName === "deposits_accounts_opened") {
  filteredRows = filteredRows.filter(r =>
    [
      "tdr_accounts",
      "rd_accounts",
      "savings_accounts",
      "current_accounts"
    ].includes(r.meaning)
  );
}

// ----------------------
// Advances
// ----------------------
if (tableName === "Advances") {
  filteredRows = filteredRows.filter(r =>
    [
      "change_over_pd_advances",
      "gdm_advances",
      "gum_advances",
      "budget_achieved_advances",
      "cy_budgeted_growth_advances"
    ].includes(r.meaning)
  );
}

// ----------------------
// NPA
// ----------------------
if (tableName === "NPA") {
  filteredRows = filteredRows.filter(r =>
    [
      "stamped_npa_accounts",
      "stamped_npa_balance"
    ].includes(r.meaning)
  );
}

// ----------------------
// Loans
// ----------------------
if (tableName === "loanssanctioned") {
  filteredRows = filteredRows.filter(r =>
    [
      "over_all_fy_accounts",
      "over_all_fy_amount"
    ].includes(r.meaning)
  );
}
const columns = filteredRows
  .filter(r => r.meaning && r.display_label)
  .map(r => ({
    label: r.display_label,
    meaning: r.meaning
  }));
console.log(
  "GET_SORTABLE_COLUMNS_SUCCESS",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    section: tableName,
    columnsCount: columns.length
  }
);
    return res.json({
      success: true,
      section: tableName,
      columns,
    });

  } catch (err) {
  console.error(
  "GET_SORTABLE_COLUMNS_FAILED",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    error: err.message
  }
);
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

app.post("/set-as-on-date", authMiddleware, async (req, res) => {
  try {

    if (req.user.role?.toLowerCase() !== "admin") {
      return res.status(403).json({
        success: false,
        message: "Access denied",
      });
    }

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

//=================GET AS ON DATE====================================
app.post("/get-as-on-date", authMiddleware, async (req, res) => {
  try {
console.log(
  "GET_AS_ON_DATE_API_HIT",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role
  }
);
    const rows = await queryUTIDatabase(
      `
        SELECT as_on_date
        FROM MIS.dbo.MIS_date
        WHERE id = 1
      `
    );

    if (!rows || rows.length === 0) {
console.warn(
  "GET_AS_ON_DATE_NOT_FOUND",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role
  }
);
      return res.status(404).json({
        success: false,
        message: "As-On Date not found",
      });
    }
console.log(
  "GET_AS_ON_DATE_SUCCESS",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    asOnDate: rows[0].as_on_date
  }
);
    return res.json({
      success: true,
      as_on_date: rows[0].as_on_date,
    });

  } catch (err) {
    console.error(
  "GET_AS_ON_DATE_FAILED",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    error: err.message
  }
);
    return res.status(500).json({
      success: false,
      message: "Server error fetching As-On Date",
    });
  }
});

// =====================================================================================
//                ✅ Fetch Activity Logs (with filters & pagination)
// ======================================================================================
app.post("/get-activity-logs", authMiddleware, async (req, res) => {
  try {
console.log(
  "GET_ACTIVITY_LOGS_API_HIT",
  {
    requestedBy: req.user?.employeeId,
    role: req.user?.role,
    filters: req.body
  }
);
    if (req.user.role?.toLowerCase() !== "admin") {

  console.warn(
    "GET_ACTIVITY_LOGS_ACCESS_DENIED",
    {
      employeeId: req.user?.employeeId,
      role: req.user?.role
    }
  );

      return res.status(403).json({
        success: false,
        message: "Access denied",
      });
    }

    const { userId, fromDate, toDate, role, action, page = 1, limit = 10 } = req.body;

    let sql = `
      SELECT a.*, e.[Emp No.] AS emp_no, e.[Employee Name] AS emp_name
      FROM [dbo].[activity_logs] a
      LEFT JOIN [dbo].[employees_master] e 
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
console.log(
  "GET_ACTIVITY_LOGS_SUCCESS",
  {
    requestedBy: req.user?.employeeId,
    records: logs.length,
    page: parseInt(page),
    totalPages
  }
);
    return res.json({ logs, page: parseInt(page), totalPages, total });
  } catch (err) {
  console.error(
  "GET_ACTIVITY_LOGS_FAILED",
  {
    requestedBy: req.user?.employeeId,
    error: err.message
  }
);
    return res.status(500).json({ message: "Server error while fetching activity logs" });
  }
});

// ======================================================================================
//                          ✅ Activity Stats (Date Filter)
// ======================================================================================
app.post("/get-activity-stats", authMiddleware, async (req, res) => {
  try {
console.log(
  "GET_ACTIVITY_STATS_API_HIT",
  {
    requestedBy: req.user?.employeeId,
    role: req.user?.role,
    filters: req.body
  }
);
    if (req.user.role?.toLowerCase() !== "admin") {
console.warn(
  "GET_ACTIVITY_STATS_ACCESS_DENIED",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role
  }
);
      return res.status(403).json({
        success: false,
        message: "Access denied",
      });
    }

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
console.log(
  "GET_ACTIVITY_STATS_SUCCESS",
  {
    requestedBy: req.user?.employeeId,
    totalLogins: stats.totalLogins || 0,
    failedLogins: stats.failedLogins || 0,
    misUploads: stats.misUploads || 0
  }
);
    res.json({
      totalLogins: stats.totalLogins || 0,
      failedLogins: stats.failedLogins || 0,
      misUploads: stats.misUploads || 0,
      passwordResets: stats.passwordResets || 0,
      lockedAccounts: stats.lockedAccounts || 0,
    });
  } catch (err) {
   console.error(
  "GET_ACTIVITY_STATS_FAILED",
  {
    requestedBy: req.user?.employeeId,
    error: err.message
  }
);
    res.status(500).json({ message: "Server error while fetching activity stats" });
  }
});

// ====================================================================================
//                         ✅ Fetch Last Activities per User
// ====================================================================================
app.post("/get-last-activities", authMiddleware, async (req, res) => {
  try {

    if (req.user.role?.toLowerCase() !== "admin") {
      return res.status(403).json({
        success: false,
        message: "Access denied",
      });
    }

    const { role } = req.body;

    let sql = `
      SELECT a1.*, e.[Emp No.] AS emp_no, e.[Employee Name] AS emp_name
      FROM [dbo].[activity_logs] a1
      INNER JOIN (
        SELECT user_id, MAX(created_at) AS max_date
        FROM [dbo].[activity_logs]
        GROUP BY user_id
      ) a2 ON a1.user_id = a2.user_id AND a1.created_at = a2.max_date
      JOIN [dbo].[employees_master] e ON a1.user_id = e.[Emp No.]
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

//===================== UPDATE ALL DEPOSITS ================================
app.post(
  "/update-deposits-all",
  authMiddleware,
  async (req, res) => {
    try {
console.log(
  "UPDATE_DEPOSITS_ALL_API_HIT",
  {
    employeeId: req.user.employeeId,
    role: req.user.role
  }
);
if (req.user.role !== "admin") {

  console.warn(
    "UPDATE_DEPOSITS_ALL_ACCESS_DENIED",
    {
      employeeId: req.user.employeeId,
      role: req.user.role
    }
  );

  await logSecurityEvent(
    req.user.employeeId,
    "Deposits Update Access Denied",
    JSON.stringify({
      role: req.user.role
    })
  );

  return res.status(403).json({
    success: false,
    message: "Access denied"
  });
}
console.log(
  "USP_DEPOSITS_UPDATE_STARTED",
  {
    employeeId: req.user.employeeId
  }
);
      await queryUTIDatabase(`
        EXEC usp_DepositsUpdate
      `);
	  
	  await logActivity(
  req.user.employeeId,
  req.user.role,
  "Deposits Update",
  "Executed usp_DepositsUpdate"
);
console.log(
  "USP_DEPOSITS_UPDATE_COMPLETED",
  {
    employeeId: req.user.employeeId
  }
);
      return res.json({
        success: true,
        message: "Deposits Update completed successfully"
      });

    } catch (err) {

      console.error(
  "UPDATE_DEPOSITS_ALL_FAILED",
  {
    employeeId: req.user.employeeId,
    error: err.message
  }
);

      return res.status(500).json({
        success: false,
        message: err.message
      });
    }
  }
);


//============================================================================
//                              AT A GLANCE – MAIN
//============================================================================
app.post("/api/glance", authMiddleware, async (req, res) => {
let {
  level,
  branch_code,
  branch_name,
  cluster_name
} = req.body;

const userLevel = req.user.level;
const designation = req.user.designation;

console.log(
  "GLANCE_API_HIT",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    userLevel,
    level,
    branchCode: branch_code,
    branchName: branch_name,
    clusterName: cluster_name
  }
);

  try {
// ==========================================
// LEVEL 1 - FORCE USER'S OWN BRANCH
// ==========================================
const branchCodeFromToken = req.user.branchCode;
const branchNameFromToken = req.user.branchName;

if (userLevel === "Level 1") {
    level = "BRANCH";
    branch_code = branchCodeFromToken;
    branch_name = branchNameFromToken;
    cluster_name = null;

    if (branch_code !== null && branch_code !== undefined) {
        const numeric = parseFloat(branch_code);

        if (!Number.isNaN(numeric)) {
            branch_code = parseInt(numeric, 10);
        }
    }
}
if (userLevel === "Level 2") {

    const match =
        designation?.match(
            /Cluster\s*Head\s*-\s*(.*)/i
        );

    if (!match) {

        return res.status(403).json({
            success:false,
            message:"Cluster mapping not found"
        });
    }

    const allowedCluster = match[1].trim();

    cluster_name = allowedCluster;

    // fetch branches of cluster
    const pool = await poolPromise;

    const allowedBranchesResult =
        await pool.request()
            .input(
                "cluster_name",
                sql.VarChar(50),
                allowedCluster
            )
            .query(`
                SELECT branch_code
                FROM MIS.dbo.Branch_Cluster_Master
                WHERE cluster_name = @cluster_name
            `);

    const allowedSet =
        new Set(
            allowedBranchesResult.recordset.map(
                x => Number(x.branch_code)
            )
        );

    // Validate branch_code
    if (
        branch_code &&
        !allowedSet.has(Number(branch_code))
    ) {

        await logActivity(
            req.user.employeeId,
            req.user.role,
            "Blocked Glance IDOR",
            `Requested=${branch_code}`
        );

        return res.status(403).json({
            success:false,
            message:"Access denied"
        });
    }

    // no branch selected => cluster view
    if (!branch_code && !branch_name) {

        level = "CLUSTER";
    }
}
    const pool = await poolPromise;
	const request = pool.request();

const r = applyLevel1Restriction({
  userLevel: req.user.level,
  requestedLevel: level,
  branch_code,
  branch_name,
  cluster_name,
  designation: req.user.designation,
});
console.log(
  "GLANCE_SCOPE_RESOLVED",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    requestedLevel: level,
    resolvedLevel: r.level,
    branchCode: r.branch_code,
    branchName: r.branch_name,
    clusterName: r.cluster_name
  }
);
   request.input("level", sql.VarChar(10), r.level);
request.input("branch_code", sql.Int, r.branch_code ?? null);
request.input("branch_name", sql.VarChar(100), r.branch_name ?? null);

request.input(
  "cluster_name",
  sql.VarChar(50),
  r.level === "CLUSTER"
    ? r.cluster_name
    : null
);
console.log(
  "GLANCE_SP_EXECUTION_STARTED",
  {
    employeeId: req.user?.employeeId,
    procedure: "get_glance_data",
    resolvedLevel: r.level
  }
);
    const result = await request.execute("get_glance_data");
console.log(
  "GLANCE_SP_EXECUTION_SUCCESS",
  {
    employeeId: req.user?.employeeId,
    card1Rows: result.recordsets?.[0]?.length || 0,
    card2Rows: result.recordsets?.[1]?.length || 0,
    card3Rows: result.recordsets?.[2]?.length || 0
  }
);
console.log(
  "GLANCE_SUCCESS",
  {
    employeeId: req.user?.employeeId,
    resolvedLevel: r.level,
    card1Available: !!result.recordsets?.[0]?.[0],
    card2Available: !!result.recordsets?.[1]?.[0],
    card3Available: !!result.recordsets?.[2]?.[0]
  }
);
await logActivity(
  req.user.employeeId,
  req.user.role,
  "View At A Glance",
  `Level=${r.level}, Branch=${r.branch_code || "-"}, Cluster=${r.cluster_name || "-"}`
);

console.log(
  "VIEW_AT_A_GLANCE_ACTIVITY_LOGGED",
  {
    employeeId: req.user?.employeeId
  }
);
    res.json({
      success: true,
resolvedLevel: r.level,
      data: {
        card1: result.recordsets?.[0]?.[0] ?? null,
        card2: result.recordsets?.[1]?.[0] ?? null,
        card3: result.recordsets?.[2]?.[0] ?? null,
      },
    });
  } catch (err) {

    console.error(
      "GLANCE_FAILED",
      {
        employeeId: req.user?.employeeId,
        role: req.user?.role,
        level,
        error: err.message
      }
    );

    res.status(500).json({
      success: false
    });
}
});


// ===============================
// GLANCE – NPA
// ===============================
app.post("/api/glance/npa", authMiddleware, async (req, res) => {
  let {
    level,
    branch_code,
    branch_name,
    cluster_name,
  } = req.body;

  const userLevel = req.user.level;
  const designation = req.user.designation;

  // Branch from authenticated user
  const branchCodeFromToken = req.user.branchCode;
  const branchNameFromToken = req.user.branchName;

  console.log("GLANCE_NPA_API_HIT", {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    level,
    branchCode: branch_code,
    branchName: branch_name,
    clusterName: cluster_name,
  });

  // ==========================================
  // LEVEL 1 - FORCE USER'S OWN BRANCH
  // ==========================================
  if (userLevel === "Level 1") {
    branch_code = branchCodeFromToken;
    branch_name = branchNameFromToken;
    level = "BRANCH";

    if (branch_code !== null && branch_code !== undefined) {
      const numeric = parseFloat(branch_code);

      if (!Number.isNaN(numeric)) {
        branch_code = parseInt(numeric, 10);
      }
    }
  }

  // ==========================================
  // LEVEL 2 CLUSTER RESOLUTION
  // ==========================================
  if (userLevel === "Level 2") {
    const match = designation?.match(/Cluster\s*Head\s*-\s*(.*)/i);

    if (match) {
      cluster_name = match[1].trim();

      console.log("GLANCE_NPA_CLUSTER_RESOLVED", {
        employeeId: req.user?.employeeId,
        clusterName: cluster_name,
      });
    }

    // Force CLUSTER only when branch is not selected
    if (
      (branch_code === null ||
        branch_code === undefined ||
        branch_code === "") &&
      !branch_name
    ) {
      level = "CLUSTER";
    }
  }

  try {
    const pool = await poolPromise;

    const r = applyLevel1Restriction({
      userLevel,
      requestedLevel: level,
      branch_code,
      branch_name,
      cluster_name,
      designation,
    });

    console.log("GLANCE_NPA_SCOPE_RESOLVED", {
      employeeId: req.user?.employeeId,
      resolvedLevel: r.level,
      branchCode: r.branch_code,
      branchName: r.branch_name,
      clusterName: r.cluster_name,
    });

    const request = pool.request();

    request.input("level", sql.VarChar(10), r.level);
    request.input(
      "branch_code",
      sql.VarChar(20),
      r.branch_code?.toString() ?? null
    );
    request.input("branch_name", sql.VarChar(100), r.branch_name ?? null);
    request.input(
      "cluster_name",
      sql.VarChar(50),
      r.level === "CLUSTER" ? r.cluster_name : null
    );

    console.log("GLANCE_NPA_SP_STARTED", {
      employeeId: req.user?.employeeId,
      procedure: "get_glance_npa_data",
    });

    const result = await request.execute("get_glance_npa_data");

    console.log("GLANCE_NPA_SUCCESS", {
      employeeId: req.user?.employeeId,
      rows: result.recordset?.length || 0,
    });

    await logActivity(
      req.user.employeeId,
      req.user.role,
      "View NPA At A Glance",
      `Level=${r.level}, Branch=${r.branch_code || "-"}, Cluster=${r.cluster_name || "-"}`
    );

    console.log("VIEW_NPA_ACTIVITY_LOGGED", {
      employeeId: req.user?.employeeId,
    });

    res.json({
      success: true,
      data: result.recordset || [],
    });

  } catch (err) {
    console.error("GLANCE_NPA_FAILED", {
      employeeId: req.user?.employeeId,
      role: req.user?.role,
      error: err.message,
    });

    res.status(500).json({
      success: false,
      message: err.message,
    });
  }
});

// ===============================
// GLANCE – SMA
// ===============================
app.post("/api/glance/sma", authMiddleware, async (req, res) => {
  let {
    level,
    branch_code,
    branch_name,
    cluster_name,
  } = req.body;

  const userLevel = req.user.level;
  const designation = req.user.designation;

  // Branch from authenticated user
  const branchCodeFromToken = req.user.branchCode;
  const branchNameFromToken = req.user.branchName;

  console.log("GLANCE_SMA_API_HIT", {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    level,
    branchCode: branch_code,
    branchName: branch_name,
    clusterName: cluster_name,
  });

  // ==========================================
  // LEVEL 1 - FORCE USER'S OWN BRANCH
  // ==========================================
  if (userLevel === "Level 1") {
    branch_code = branchCodeFromToken;
    branch_name = branchNameFromToken;
    level = "BRANCH";

    if (branch_code !== null && branch_code !== undefined) {
      const numeric = parseFloat(branch_code);

      if (!Number.isNaN(numeric)) {
        branch_code = parseInt(numeric, 10);
      }
    }
  }

  // ==========================================
  // LEVEL 2 CLUSTER RESOLUTION
  // ==========================================
  if (userLevel === "Level 2") {
    const match = designation?.match(/Cluster\s*Head\s*-\s*(.*)/i);

    if (match) {
      cluster_name = match[1].trim();

      console.log("GLANCE_SMA_CLUSTER_RESOLVED", {
        employeeId: req.user?.employeeId,
        clusterName: cluster_name,
      });
    }

    // Force CLUSTER only when branch is not selected
    if (
      (branch_code === null ||
        branch_code === undefined ||
        branch_code === "") &&
      !branch_name
    ) {
      level = "CLUSTER";
    }
  }

  try {
    const pool = await poolPromise;

    const r = applyLevel1Restriction({
      userLevel,
      requestedLevel: level,
      branch_code,
      branch_name,
      cluster_name,
      designation,
    });

    console.log("GLANCE_SMA_SCOPE_RESOLVED", {
      employeeId: req.user?.employeeId,
      resolvedLevel: r.level,
      branchCode: r.branch_code,
      branchName: r.branch_name,
      clusterName: r.cluster_name,
    });

    const request = pool.request();

    request.input("level", sql.VarChar(10), r.level);
    request.input("branch_code", sql.Int, r.branch_code ?? null);
    request.input("branch_name", sql.VarChar(100), r.branch_name ?? null);
    request.input(
      "cluster_name",
      sql.VarChar(50),
      r.level === "CLUSTER" ? r.cluster_name : null
    );

    console.log("GLANCE_SMA_SP_STARTED", {
      employeeId: req.user?.employeeId,
      procedure: "get_glance_sma_data",
    });

    const result = await request.execute("get_glance_sma_data");

    console.log("GLANCE_SMA_SUCCESS", {
      employeeId: req.user?.employeeId,
      rows: result.recordset?.length || 0,
    });

    await logActivity(
      req.user.employeeId,
      req.user.role,
      "View SMA At A Glance",
      `Level=${r.level}, Branch=${r.branch_code || "-"}, Cluster=${r.cluster_name || "-"}`
    );

    console.log("VIEW_SMA_ACTIVITY_LOGGED", {
      employeeId: req.user?.employeeId,
    });

    res.json({
      success: true,
      data: result.recordset || [],
    });

  } catch (err) {
    console.error("GLANCE_SMA_FAILED", {
      employeeId: req.user?.employeeId,
      role: req.user?.role,
      error: err.message,
    });

    res.status(500).json({
      success: false,
      message: err.message,
    });
  }
});

// ===============================
// DOA – AT A GLANCE (CARD 4) – FINAL & STABLE
// ===============================
// ===============================
// GLANCE – DOA CARD 4
// ===============================
app.post("/api/glance/doa-card4", authMiddleware, async (req, res) => {
  let {
    level,
    branch_code,
    branch_name,
    cluster_name,
  } = req.body;

  const userLevel = req.user.level;
  const designation = req.user.designation;

  // Branch from authenticated user
  const branchCodeFromToken = req.user.branchCode;
  const branchNameFromToken = req.user.branchName;

  console.log("GLANCE_DOA_API_HIT", {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    level,
    branchCode: branch_code,
    clusterName: cluster_name,
  });

  // ==========================================
  // LEVEL 1 - FORCE USER'S OWN BRANCH
  // ==========================================
  if (userLevel === "Level 1") {
    branch_code = branchCodeFromToken;
    branch_name = branchNameFromToken;
    level = "BRANCH";

    if (branch_code !== null && branch_code !== undefined) {
      const numeric = parseFloat(branch_code);

      if (!Number.isNaN(numeric)) {
        branch_code = parseInt(numeric, 10);
      }
    }
  }

  // ==========================================
  // LEVEL 2 CLUSTER RESOLUTION
  // ==========================================
  if (userLevel === "Level 2") {
    const match = designation?.match(/Cluster\s*Head\s*-\s*(.*)/i);

    if (match) {
      cluster_name = match[1].trim();

      console.log("GLANCE_DOA_CLUSTER_RESOLVED", {
        employeeId: req.user?.employeeId,
        clusterName: cluster_name,
      });
    }

    // Force CLUSTER only when branch is not selected
    if (
      (branch_code === null ||
        branch_code === undefined ||
        branch_code === "") &&
      !branch_name
    ) {
      level = "CLUSTER";
    }
  }

  try {
    const pool = await poolPromise;

    const r = applyLevel1Restriction({
      userLevel,
      requestedLevel: level,
      branch_code,
      branch_name,
      cluster_name,
      designation,
    });

    console.log("GLANCE_DOA_SCOPE_RESOLVED", {
      employeeId: req.user?.employeeId,
      resolvedLevel: r.level,
      branchCode: r.branch_code,
      branchName: r.branch_name,
      clusterName: r.cluster_name,
    });

    const request = pool.request();

    request.input("level", sql.VarChar(10), r.level);
    request.input("branch_code", sql.Int, r.branch_code ?? null);
    request.input("branch_name", sql.NVarChar(100), r.branch_name ?? null);
    request.input(
      "cluster_name",
      sql.VarChar(50),
      r.level === "CLUSTER" ? r.cluster_name : null
    );

    console.log("GLANCE_DOA_SP_STARTED", {
      employeeId: req.user?.employeeId,
      procedure: "get_glance_doa_card4",
    });

    const result = await request.execute("get_glance_doa_card4");

    const rows = result.recordsets
      ? result.recordsets.flat()
      : result.recordset || [];

    // ===============================
    // RESHAPE FOR UI
    // ===============================
    const data = {
      TDR: { type: "TDR(≥ 1 LAKH)" },
      RD: { type: "RD" },
      Savings: { type: "Savings" },
      Current: { type: "Current" },
    };

    rows.forEach((row) => {
      if (row.section === "As on Date") {
        data.TDR.on_acs = row.tdr_acs;
        data.TDR.on_amt = row.tdr_amt;
        data.RD.on_acs = row.rd_acs;
        data.RD.on_amt = row.rd_amt;
        data.Savings.on_acs = row.sb_acs;
        data.Savings.on_amt = row.sb_amt;
        data.Current.on_acs = row.ca_acs;
        data.Current.on_amt = row.ca_amt;
      }

      if (row.section === "Month To Date") {
        data.TDR.mtd_acs = row.tdr_acs;
        data.TDR.mtd_amt = row.tdr_amt;
        data.RD.mtd_acs = row.rd_acs;
        data.RD.mtd_amt = row.rd_amt;
        data.Savings.mtd_acs = row.sb_acs;
        data.Savings.mtd_amt = row.sb_amt;
        data.Current.mtd_acs = row.ca_acs;
        data.Current.mtd_amt = row.ca_amt;
      }

      if (row.section === "Previous Year") {
        data.TDR.ytd_acs = row.tdr_acs;
        data.TDR.ytd_amt = row.tdr_amt;
        data.RD.ytd_acs = row.rd_acs;
        data.RD.ytd_amt = row.rd_amt;
        data.Savings.ytd_acs = row.sb_acs;
        data.Savings.ytd_amt = row.sb_amt;
        data.Current.ytd_acs = row.ca_acs;
        data.Current.ytd_amt = row.ca_amt;
      }
    });

    console.log("GLANCE_DOA_SUCCESS", {
      employeeId: req.user?.employeeId,
      rows: rows.length,
    });

    await logActivity(
      req.user.employeeId,
      req.user.role,
      "View DOA At A Glance",
      `Level=${r.level}, Branch=${r.branch_code || "-"}, Cluster=${r.cluster_name || "-"}`
    );

    console.log("VIEW_DOA_ACTIVITY_LOGGED", {
      employeeId: req.user?.employeeId,
    });

    return res.status(200).json({
      success: true,
      data: Object.values(data),
    });

  } catch (err) {
    console.error("GLANCE_DOA_FAILED", {
      employeeId: req.user?.employeeId,
      role: req.user?.role,
      error: err.message,
    });

    return res.status(500).json({
      success: false,
      message: "Failed to fetch DOA Card 4 data",
    });
  }
});

// ===============================
// GLANCE – LOANS OPENED
// ===============================
app.post("/api/glance/loans-opened", authMiddleware, async (req, res) => {
  let {
    level,
    branch_code,
    branch_name,
    cluster_name,
  } = req.body;

  const userLevel = req.user.level;
  const designation = req.user.designation;

  // Branch from authenticated user
  const branchCodeFromToken = req.user.branchCode;
  const branchNameFromToken = req.user.branchName;

  console.log("GLANCE_LOANS_OPENED_API_HIT", {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    level,
    branchCode: branch_code,
    branchName: branch_name,
    clusterName: cluster_name,
  });

  // ==========================================
  // LEVEL 1 - FORCE USER'S OWN BRANCH
  // ==========================================
  if (userLevel === "Level 1") {
    branch_code = branchCodeFromToken;
    branch_name = branchNameFromToken;
    level = "BRANCH";

    if (branch_code !== null && branch_code !== undefined) {
      const numeric = parseFloat(branch_code);

      if (!Number.isNaN(numeric)) {
        branch_code = parseInt(numeric, 10);
      }
    }
  }

  // ==========================================
  // LEVEL 2 CLUSTER RESOLUTION
  // ==========================================
  if (userLevel === "Level 2") {
    const match = designation?.match(/Cluster\s*Head\s*-\s*(.*)/i);

    if (match) {
      cluster_name = match[1].trim();

      console.log("GLANCE_LOANS_OPENED_CLUSTER_RESOLVED", {
        employeeId: req.user?.employeeId,
        clusterName: cluster_name,
      });
    }

    // Force CLUSTER only when branch is not selected
    if (
      (branch_code === null ||
        branch_code === undefined ||
        branch_code === "") &&
      !branch_name
    ) {
      level = "CLUSTER";
    }
  }

  try {
    const pool = await poolPromise;

    const r = applyLevel1Restriction({
      userLevel,
      requestedLevel: level,
      branch_code,
      branch_name,
      cluster_name,
      designation,
    });

    console.log("GLANCE_LOANS_OPENED_SCOPE_RESOLVED", {
      employeeId: req.user?.employeeId,
      resolvedLevel: r.level,
      branchCode: r.branch_code,
      branchName: r.branch_name,
      clusterName: r.cluster_name,
    });

    const request = pool.request();

    request.input("level", sql.VarChar(10), r.level);
    request.input("branch_code", sql.Int, r.branch_code ?? null);
    request.input("branch_name", sql.NVarChar(100), r.branch_name ?? null);
    request.input(
      "cluster_name",
      sql.VarChar(50),
      r.level === "CLUSTER" ? r.cluster_name : null
    );

    console.log("GLANCE_LOANS_OPENED_SP_STARTED", {
      employeeId: req.user?.employeeId,
      procedure: "get_glance_loans_card5",
    });

    const result = await request.execute("get_glance_loans_card5");

    await logActivity(
      req.user.employeeId,
      req.user.role,
      "View Loans Opened At A Glance",
      `Level=${r.level}, Branch=${r.branch_code || "-"}, Cluster=${r.cluster_name || "-"}`
    );

    console.log("VIEW_LOANS_OPENED_ACTIVITY_LOGGED", {
      employeeId: req.user?.employeeId,
    });

    console.log("GLANCE_LOANS_OPENED_SUCCESS", {
      employeeId: req.user?.employeeId,
      rows: result.recordset?.length || 0,
    });

    res.json({
      success: true,
      rows: result.recordset || [],
    });

  } catch (err) {
    console.error("GLANCE_LOANS_OPENED_FAILED", {
      employeeId: req.user?.employeeId,
      role: req.user?.role,
      error: err.message,
    });

    res.status(500).json({
      success: false,
      message: err.message,
    });
  }
});


// ===============================
// GLANCE – EMPLOYEE EFFICIENCY
// ===============================
app.post("/api/glance/employee-efficiency", authMiddleware, async (req, res) => {
  let {
    level,
    branch_code,
    branch_name,
    cluster_name,
  } = req.body;

  const userLevel = req.user.level;
  const designation = req.user.designation;

  // Branch from authenticated user
  const branchCodeFromToken = req.user.branchCode;
  const branchNameFromToken = req.user.branchName;

  console.log("GLANCE_EMPLOYEE_EFFICIENCY_API_HIT", {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    level,
    branchCode: branch_code,
    branchName: branch_name,
    clusterName: cluster_name,
  });

  // ==========================================
  // LEVEL 1 - FORCE USER'S OWN BRANCH
  // ==========================================
  if (userLevel === "Level 1") {
    branch_code = branchCodeFromToken;
    branch_name = branchNameFromToken;
    level = "BRANCH";

    if (branch_code !== null && branch_code !== undefined) {
      const numeric = parseFloat(branch_code);

      if (!Number.isNaN(numeric)) {
        branch_code = parseInt(numeric, 10);
      }
    }
  }

  // ==========================================
  // LEVEL 2 CLUSTER RESOLUTION
  // ==========================================
  if (userLevel === "Level 2") {
    const match = designation?.match(/Cluster\s*Head\s*-\s*(.*)/i);

    if (match) {
      cluster_name = match[1].trim();

      console.log("GLANCE_EMPLOYEE_EFFICIENCY_CLUSTER_RESOLVED", {
        employeeId: req.user?.employeeId,
        clusterName: cluster_name,
      });
    }

    // Force CLUSTER only when branch is not selected
    if (
      (branch_code === null ||
        branch_code === undefined ||
        branch_code === "") &&
      !branch_name
    ) {
      level = "CLUSTER";
    }
  }

  try {
    const pool = await poolPromise;

    const r = applyLevel1Restriction({
      userLevel,
      requestedLevel: level,
      branch_code,
      branch_name,
      cluster_name,
      designation,
    });

    console.log("GLANCE_EMPLOYEE_EFFICIENCY_SCOPE_RESOLVED", {
      employeeId: req.user?.employeeId,
      resolvedLevel: r.level,
      branchCode: r.branch_code,
      branchName: r.branch_name,
      clusterName: r.cluster_name,
    });

    const request = pool.request();

    request.input("level", sql.VarChar(10), r.level);
    request.input("branch_code", sql.Int, r.branch_code ?? null);
    request.input("branch_name", sql.VarChar(100), r.branch_name ?? null);
    request.input(
      "cluster_name",
      sql.VarChar(50),
      r.level === "CLUSTER" ? r.cluster_name : null
    );

    console.log("GLANCE_EMPLOYEE_EFFICIENCY_SP_STARTED", {
      employeeId: req.user?.employeeId,
      procedure: "get_glance_employee_efficiency",
    });

    const result = await request.execute(
      "get_glance_employee_efficiency"
    );

    console.log("GLANCE_EMPLOYEE_EFFICIENCY_SUCCESS", {
      employeeId: req.user?.employeeId,
    });

    await logActivity(
      req.user.employeeId,
      req.user.role,
      "View Employee Efficiency At A Glance",
      `Level=${r.level}, Branch=${r.branch_code || "-"}, Cluster=${r.cluster_name || "-"}`
    );

    console.log("VIEW_EMPLOYEE_EFFICIENCY_ACTIVITY_LOGGED", {
      employeeId: req.user?.employeeId,
    });

    res.json({
      success: true,
      data: result.recordset?.[0] || null,
    });

  } catch (err) {
    console.error("GLANCE_EMPLOYEE_EFFICIENCY_FAILED", {
      employeeId: req.user?.employeeId,
      role: req.user?.role,
      error: err.message,
    });

    res.status(500).json({
      success: false,
      message: err.message,
    });
  }
});

// ===============================
// GLANCE – BUSINESS CORRESPONDENTS
// ===============================
app.post("/api/glance/business-correspondents", authMiddleware, async (req, res) => {
  let {
    level,
    branch_code,
    branch_name,
    cluster_name,
  } = req.body;

  const userLevel = req.user.level;
  const designation = req.user.designation;

  // Branch from authenticated user
  const branchCodeFromToken = req.user.branchCode;
  const branchNameFromToken = req.user.branchName;

  // ==========================================
  // LEVEL 1 - FORCE USER'S OWN BRANCH
  // ==========================================
  if (userLevel === "Level 1") {
    branch_code = branchCodeFromToken;
    branch_name = branchNameFromToken;
    level = "BRANCH";

    if (branch_code !== null && branch_code !== undefined) {
      const numeric = parseFloat(branch_code);

      if (!Number.isNaN(numeric)) {
        branch_code = parseInt(numeric, 10);
      }
    }
  }

  // ==========================================
  // LEVEL 2 CLUSTER RESOLUTION
  // ==========================================
  if (userLevel === "Level 2") {
    const match = designation?.match(/Cluster\s*Head\s*-\s*(.*)/i);

    if (match) {
      cluster_name = match[1].trim();
    }

    if (
      (branch_code === null ||
        branch_code === undefined ||
        branch_code === "") &&
      !branch_name
    ) {
      level = "CLUSTER";
    }
  }

  const r = applyLevel1Restriction({
    userLevel,
    requestedLevel: level,
    branch_code,
    branch_name,
    cluster_name,
    designation,
  });

  console.log("GLANCE_BC_API_HIT", {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    resolvedLevel: r.level,
    branchCode: r.branch_code,
    branchName: r.branch_name,
    clusterName: r.cluster_name,
  });

  try {
    const pool = await poolPromise;
    const request = pool.request();

    request.input("level", sql.VarChar(10), r.level);
    request.input("branch_code", sql.Int, r.branch_code ?? null);
    request.input("branch_name", sql.VarChar(100), r.branch_name ?? null);
    request.input(
      "cluster_name",
      sql.VarChar(50),
      r.level === "CLUSTER" ? r.cluster_name : null
    );

    console.log("GLANCE_BC_SP_STARTED", {
      employeeId: req.user?.employeeId,
      procedure: "get_glance_business_correspondents",
    });

    const result = await request.execute(
      "get_glance_business_correspondents"
    );

    await logActivity(
      req.user.employeeId,
      req.user.role,
      "View BC At A Glance",
      `Level=${r.level}, Branch=${r.branch_code || "-"}, Cluster=${r.cluster_name || "-"}`
    );

    console.log("VIEW_BC_ACTIVITY_LOGGED", {
      employeeId: req.user?.employeeId,
    });

    console.log("GLANCE_BC_SUCCESS", {
      employeeId: req.user?.employeeId,
    });

    res.json({
      success: true,
      data: result.recordset?.[0] || null,
    });

  } catch (err) {
    console.error("GLANCE_BC_FAILED", {
      employeeId: req.user?.employeeId,
      role: req.user?.role,
      error: err.message,
    });

    res.status(500).json({
      success: false,
      message: err.message,
    });
  }
});
// ===============================
// ATMs – CARD 1 (REUSABLE)
// ===============================
app.post("/api/glance/atms", authMiddleware, async (req, res) => {
  let {
    level,
    branch_code,
    branch_name,
    cluster_name,
  } = req.body;

  const userLevel = req.user.level;
  const designation = req.user.designation;

  // Branch from authenticated user
  const branchCodeFromToken = req.user.branchCode;
  const branchNameFromToken = req.user.branchName;

  // ==========================================
  // LEVEL 1 - FORCE USER'S OWN BRANCH
  // ==========================================
  if (userLevel === "Level 1") {
    branch_code = branchCodeFromToken;
    branch_name = branchNameFromToken;
    level = "BRANCH";

    if (branch_code !== null && branch_code !== undefined) {
      const numeric = parseFloat(branch_code);

      if (!Number.isNaN(numeric)) {
        branch_code = parseInt(numeric, 10);
      }
    }
  }

  // ==========================================
  // LEVEL 2 CLUSTER RESOLUTION
  // ==========================================
  if (userLevel === "Level 2") {
    const match = designation?.match(/Cluster\s*Head\s*-\s*(.*)/i);

    if (match) {
      cluster_name = match[1].trim();
    }

    if (
      (branch_code === null ||
        branch_code === undefined ||
        branch_code === "") &&
      !branch_name
    ) {
      level = "CLUSTER";
    }
  }

  const r = applyLevel1Restriction({
    userLevel,
    requestedLevel: level,
    branch_code,
    branch_name,
    cluster_name,
    designation,
  });

  console.log("GLANCE_ATM_API_HIT", {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    resolvedLevel: r.level,
    branchCode: r.branch_code,
    branchName: r.branch_name,
    clusterName: r.cluster_name,
  });

  try {
    const pool = await poolPromise;

    const request = pool.request();

    request.input("level", sql.VarChar(10), r.level);
    request.input("branch_code", sql.Int, r.branch_code ?? null);
    request.input("branch_name", sql.VarChar(100), r.branch_name ?? null);
    request.input(
      "cluster_name",
      sql.VarChar(50),
      r.level === "CLUSTER" ? r.cluster_name : null
    );

    console.log("GLANCE_ATM_SP_STARTED", {
      employeeId: req.user?.employeeId,
      procedure: "get_glance_atms",
    });

    const result = await request.execute("get_glance_atms");

    await logActivity(
      req.user.employeeId,
      req.user.role,
      "View ATM At A Glance",
      `Level=${r.level}, Branch=${r.branch_code || "-"}, Cluster=${r.cluster_name || "-"}`
    );

    console.log("VIEW_ATM_ACTIVITY_LOGGED", {
      employeeId: req.user?.employeeId,
    });

    console.log("GLANCE_ATM_SUCCESS", {
      employeeId: req.user?.employeeId,
    });

    res.json({
      success: true,
      data: result.recordset?.[0] || null,
    });

  } catch (err) {
    console.error("GLANCE_ATM_FAILED", {
      employeeId: req.user?.employeeId,
      role: req.user?.role,
      error: err.message,
    });

    res.status(500).json({
      success: false,
      message: err.message,
    });
  }
});
// ================================================
// 🔐 VALIDATE BRANCH ↔ CLUSTER (LEVEL 2 SECURITY)
// ================================================
app.post("/api/validate-branch-cluster", authMiddleware, async (req, res) => {
let {
  level,
  branch_code,
  branch_name,
  cluster_name
} = req.body;

const userLevel = req.user.level;
const designation = req.user.designation;

console.log(
  "VALIDATE_BRANCH_CLUSTER_API_HIT",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    userLevel,
    branchCode: branch_code,
    branchName: branch_name,
    clusterName: cluster_name
  }
);

if (userLevel === "Level 2") {
  const match = designation?.match(
    /Cluster\s*Head\s*-\s*(.*)/i
  );

if (match) {

    cluster_name = match[1].trim();

    console.log(
      "VALIDATE_BRANCH_CLUSTER_CLUSTER_RESOLVED",
      {
        employeeId: req.user?.employeeId,
        clusterName: cluster_name
      }
    );
}

  // Force CLUSTER only when branch is not selected
  if (
    (branch_code === null || branch_code === undefined || branch_code === "") &&
    !branch_name
  ) {
    level = "CLUSTER";
  }
}
  if (!cluster_name || (branch_code == null && !branch_name)) {
console.warn(
  "VALIDATE_BRANCH_CLUSTER_INVALID_INPUT",
  {
    employeeId: req.user?.employeeId,
    branchCode: branch_code,
    branchName: branch_name,
    clusterName: cluster_name
  }
);
    return res.status(400).json({ success: false, message: "Invalid input" });
  }

  try {
    const pool = await poolPromise;
    const request = pool.request();

    let query = `
      SELECT COUNT(*) AS cnt
      FROM Branch_Cluster_Master
      WHERE LOWER(LTRIM(RTRIM(cluster_name)))
            = LOWER(LTRIM(RTRIM(@cluster_name)))
    `;

    request.input("cluster_name", sql.VarChar(50), cluster_name);

    if (branch_code !== null && branch_code !== undefined) {
      query += `
        AND TRY_CAST(LTRIM(RTRIM(branch_code)) AS INT)
            =
            TRY_CAST(@branch_code AS INT)
      `;
const safeBranchCode = String(branch_code).replace(/,/g, "").trim();

request.input("branch_code", sql.VarChar(50), safeBranchCode);
    } else {
      query += `
        AND LOWER(LTRIM(RTRIM(branch_name)))
            =
            LOWER(LTRIM(RTRIM(@branch_name)))
      `;
      request.input("branch_name", sql.VarChar(100), branch_name);
    }

    const result = await request.query(query);
    const count = result.recordset[0].cnt;

console.log(
  "VALIDATE_BRANCH_CLUSTER_QUERY_RESULT",
  {
    employeeId: req.user?.employeeId,
    count
  }
);

    if (count === 0) {
console.warn(
  "VALIDATE_BRANCH_CLUSTER_ACCESS_DENIED",
  {
    employeeId: req.user?.employeeId,
    branchCode: branch_code,
    branchName: branch_name,
    clusterName: cluster_name
  }
);
      return res.status(403).json({
        success: false,
        message: "Branch does not belong to your cluster",
      });
    }
console.log(
  "VALIDATE_BRANCH_CLUSTER_SUCCESS",
  {
    employeeId: req.user?.employeeId,
    branchCode: branch_code,
    branchName: branch_name,
    clusterName: cluster_name
  }
);
    return res.json({ success: true });

  } catch (err) {
   console.error(
  "VALIDATE_BRANCH_CLUSTER_FAILED",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    error: err.message
  }
);
    return res.status(500).json({ success: false, message: "Server error" });
  }
});

app.post("/get-branches-by-cluster", authMiddleware, async (req, res) => {
let { cluster_name } = req.body;

const userLevel = req.user.level;
const designation = req.user.designation;

console.log(
  "GET_BRANCHES_BY_CLUSTER_API_HIT",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    clusterName: cluster_name
  }
);

if (userLevel === "Level 2") {
  const match = designation?.match(
    /Cluster\s*Head\s*-\s*(.*)/i
  );

  if (match) {
    cluster_name = match[1].trim();
console.log(
  "GET_BRANCHES_BY_CLUSTER_CLUSTER_RESOLVED",
  {
    employeeId: req.user?.employeeId,
    clusterName: cluster_name
  }
);
  }
}

  try {
    const pool = await poolPromise;

    const result = await pool
      .request()
      .input("cluster", sql.VarChar(50), cluster_name)
      .query(`
        SELECT branch_code, branch_name
        FROM Branch_Cluster_Master
        WHERE cluster_name = @cluster
        ORDER BY branch_name
      `);
console.log(
  "GET_BRANCHES_BY_CLUSTER_SUCCESS",
  {
    employeeId: req.user?.employeeId,
    clusterName: cluster_name,
    branchesCount: result.recordset?.length || 0
  }
);
    return res.json({
      success: true,
      branches: result.recordset,
    });
  } catch (err) {
    console.error(
  "GET_BRANCHES_BY_CLUSTER_FAILED",
  {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    error: err.message
  }
);
    return res.status(500).json({
      success: false,
      message: "Server error",
    });
  }
});

// ===============================
// GLANCE – VOUCHERS CARD
// ===============================
// ===============================
// GLANCE – VOUCHERS
// ===============================
app.post("/api/glance/vouchers", authMiddleware, async (req, res) => {
  let {
    level,
    branch_code,
    branch_name,
    cluster_name,
  } = req.body;

  const userLevel = req.user.level;
  const designation = req.user.designation;

  // Branch from authenticated user
  const branchCodeFromToken = req.user.branchCode;
  const branchNameFromToken = req.user.branchName;

  console.log("GLANCE_VOUCHERS_API_HIT", {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    level,
    branchCode: branch_code,
    branchName: branch_name,
    clusterName: cluster_name,
  });

  // ==========================================
  // LEVEL 1 - FORCE USER'S OWN BRANCH
  // ==========================================
  if (userLevel === "Level 1") {
    branch_code = branchCodeFromToken;
    branch_name = branchNameFromToken;
    level = "BRANCH";

    if (branch_code !== null && branch_code !== undefined) {
      const numeric = parseFloat(branch_code);

      if (!Number.isNaN(numeric)) {
        branch_code = parseInt(numeric, 10);
      }
    }
  }

  // ==========================================
  // LEVEL 2 CLUSTER RESOLUTION
  // ==========================================
  if (userLevel === "Level 2") {
    const match = designation?.match(/Cluster\s*Head\s*-\s*(.*)/i);

    if (match) {
      cluster_name = match[1].trim();

      console.log("GLANCE_VOUCHERS_CLUSTER_RESOLVED", {
        employeeId: req.user?.employeeId,
        clusterName: cluster_name,
      });
    }

    // Force CLUSTER only when branch is not selected
    if (
      (branch_code === null ||
        branch_code === undefined ||
        branch_code === "") &&
      !branch_name
    ) {
      level = "CLUSTER";
    }
  }

  try {
    const pool = await poolPromise;

    const r = applyLevel1Restriction({
      userLevel,
      requestedLevel: level,
      branch_code,
      branch_name,
      cluster_name,
      designation,
    });

    console.log("GLANCE_VOUCHERS_SCOPE_RESOLVED", {
      employeeId: req.user?.employeeId,
      resolvedLevel: r.level,
      branchCode: r.branch_code,
      branchName: r.branch_name,
      clusterName: r.cluster_name,
    });

    const request = pool.request();

    request.input("level", sql.VarChar(10), r.level);
    request.input("branch_code", sql.Int, r.branch_code ?? null);
    request.input("branch_name", sql.VarChar(100), r.branch_name ?? null);
    request.input(
      "cluster_name",
      sql.VarChar(50),
      r.level === "CLUSTER" ? r.cluster_name : null
    );

    console.log("GLANCE_VOUCHERS_SP_STARTED", {
      employeeId: req.user?.employeeId,
      procedure: "get_glance_vouchers",
    });

    const result = await request.execute("get_glance_vouchers");

    console.log("GLANCE_VOUCHERS_SUCCESS", {
      employeeId: req.user?.employeeId,
    });

    await logActivity(
      req.user.employeeId,
      req.user.role,
      "View Vouchers At A Glance",
      `Level=${r.level}, Branch=${r.branch_code || "-"}, Cluster=${r.cluster_name || "-"}`
    );

    console.log("VIEW_VOUCHERS_ACTIVITY_LOGGED", {
      employeeId: req.user?.employeeId,
    });

    res.json({
      success: true,
      data: result.recordset?.[0] || null,
    });

  } catch (err) {
    console.error("GLANCE_VOUCHERS_FAILED", {
      employeeId: req.user?.employeeId,
      role: req.user?.role,
      error: err.message,
    });

    res.status(500).json({
      success: false,
      message: err.message,
    });
  }
});

// ===============================
// GLANCE – RBIA AUDIT CARD
// ===============================
app.post("/api/glance/rbia-audit", authMiddleware, async (req, res) => {
  let {
    level,
    branch_code,
    branch_name,
    cluster_name,
  } = req.body;

  const userLevel = req.user.level;
  const designation = req.user.designation;

  // Branch details from authenticated user
  const branchCodeFromToken = req.user.branchCode;
  const branchNameFromToken = req.user.branchName;

  console.log("GLANCE_RBIA_AUDIT_API_HIT", {
    employeeId: req.user?.employeeId,
    role: req.user?.role,
    level,
    branchCode: branch_code,
    branchName: branch_name,
    clusterName: cluster_name,
  });

  // ==========================================
  // LEVEL 1 - FORCE USER'S OWN BRANCH
  // ==========================================
  if (userLevel === "Level 1") {
    branch_code = branchCodeFromToken;
    branch_name = branchNameFromToken;
    level = "BRANCH";

    if (branch_code !== null && branch_code !== undefined) {
      const numeric = parseFloat(branch_code);

      if (!Number.isNaN(numeric)) {
        branch_code = parseInt(numeric, 10);
      }
    }
  }

  // ==========================================
  // LEVEL 2 CLUSTER RESOLUTION
  // ==========================================
  if (userLevel === "Level 2") {
    const match = designation?.match(/Cluster\s*Head\s*-\s*(.*)/i);

    if (match) {
      cluster_name = match[1].trim();

      console.log("GLANCE_RBIA_AUDIT_CLUSTER_RESOLVED", {
        employeeId: req.user?.employeeId,
        clusterName: cluster_name,
      });
    }

    if (
      (branch_code === null ||
        branch_code === undefined ||
        branch_code === "") &&
      !branch_name
    ) {
      level = "CLUSTER";
    }
  }

  try {
    const pool = await poolPromise;

    const r = applyLevel1Restriction({
      userLevel,
      requestedLevel: level,
      branch_code,
      branch_name,
      cluster_name,
      designation,
    });

    console.log("GLANCE_RBIA_AUDIT_SCOPE_RESOLVED", {
      employeeId: req.user?.employeeId,
      resolvedLevel: r.level,
      branchCode: r.branch_code,
      branchName: r.branch_name,
      clusterName: r.cluster_name,
    });

    const request = pool.request();

    request.input("level", sql.VarChar(20), r.level);
    request.input("branch_code", sql.Int, r.branch_code ?? null);
    request.input("branch_name", sql.VarChar(150), r.branch_name ?? null);
    request.input(
      "cluster_name",
      sql.VarChar(150),
      r.level === "CLUSTER" ? r.cluster_name : null
    );

    console.log("GLANCE_RBIA_AUDIT_SP_STARTED", {
      employeeId: req.user?.employeeId,
      procedure: "get_glance_rbia_audit",
    });

    const result = await request.execute("get_glance_rbia_audit");

    console.log("GLANCE_RBIA_AUDIT_SUCCESS", {
      employeeId: req.user?.employeeId,
      records: result.recordset?.length || 0,
    });

    await logActivity(
      req.user.employeeId,
      req.user.role,
      "View RBIA Audit At A Glance",
      `Level=${r.level}, Branch=${r.branch_code || "-"}, Cluster=${r.cluster_name || "-"}`
    );

    console.log("VIEW_RBIA_AUDIT_ACTIVITY_LOGGED", {
      employeeId: req.user?.employeeId,
    });

    res.json({
      success: true,
      data: result.recordset || [],
    });

  } catch (err) {
    console.error("GLANCE_RBIA_AUDIT_FAILED", {
      employeeId: req.user?.employeeId,
      role: req.user?.role,
      error: err.message,
    });

    res.status(500).json({
      success: false,
      message: err.message,
    });
  }
});
// =====================================================================
//                              Data Upload
// =======================================================================

app.post(
"/api/data-upload",
authMiddleware,
uploadLimiter,
upload.single("file"),
  async (req, res) => {

const userId = req.user.employeeId;
const role = req.user.role;

let transaction = null;

console.log(
  "EMPLOYEE_UPLOAD_API_HIT",
  {
    employeeId: userId,
    role,
    fileName: req.file?.originalname,
    fileSize: req.file?.size
  }
);

if (role?.toLowerCase() !== "admin") {

  await logSecurityEvent(
    userId,
    "Employees Upload Access Denied",
    JSON.stringify({ role })
  );

  return res.status(403).json({
    success: false,
    message: "Access denied"
  });
}

    try {

      if (!req.file) {
        return res.status(400).json({
          success: false,
          message: "Please select a file"
        });
      }

      const fileBuffer = req.file.buffer;

const detectedType =
  await fileTypeFromBuffer(fileBuffer);

if (
  req.file.originalname
    .toLowerCase()
    .endsWith(".xlsx")
) {

  if (
    !detectedType ||
    detectedType.mime !==
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  ) {

    await logSecurityEvent(
      userId,
      "Invalid Employee XLSX Upload",
      req.file.originalname
    );

    return res.status(400).json({
      success: false,
      message: "Invalid XLSX file"
    });
  }
}

      const workbook = XLSX.read(
  fileBuffer,
  { type: "buffer" }
);

      const sheet =
        workbook.Sheets[
          workbook.SheetNames[0]
        ];

      const jsonData =
        XLSX.utils.sheet_to_json(sheet);

      if (jsonData.length === 0) {
        return res.status(400).json({
          success: false,
          message: "Excel file is empty"
        });
      }

      // ===========================
      // COLUMN VALIDATION
      // ===========================

      const excelColumns =
  Object.keys(jsonData[0]).map(col =>
    col
      .replace(/\n/g, " ")
      .replace(/\s+/g, " ")
      .trim()
  );

      const requiredColumns = [
  "Emp no",
  "Employee Name",
  "Br Code",
  "Branch Name",
  "Designation",
  "Contact Number",
  "Cluster"
];

console.log(
  "EXCEL COLUMNS =",
  excelColumns
);

      const missingColumns =
        requiredColumns.filter(
          column =>
            !excelColumns.includes(column)
        );

      if (missingColumns.length > 0) {

        return res.status(400).json({
          success: false,
          message:
            "Missing Columns : " +
            missingColumns.join(", ")
        });

      }

    const pool =
  await poolPromise;

transaction =
  new sql.Transaction(pool);

await transaction.begin();

// ===========================
// DELETE OLD DATA
// ===========================

console.log("DELETE STARTED");

const deleteResult =
  await transaction
  .request()
  .query(`
    DELETE FROM employees_master
  `);

console.log("DELETE COMPLETED");

const countAfterDelete =
  await transaction.request().query(`
    SELECT COUNT(*) AS total
    FROM employees_master
  `);

console.log(
  "COUNT AFTER DELETE =",
  countAfterDelete.recordset[0].total
);

let insertedCount = 0;

      // ===========================
      // INSERT NEW DATA
      // ===========================

      for (const row of jsonData) {

  const cleanRow = {};

  Object.keys(row).forEach(key => {

    const cleanKey = key
      .replace(/\n/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    cleanRow[cleanKey] = row[key];

  });

  const empId =
  String(cleanRow["Emp no"] || "").trim();

  if (!empId) continue;

  const empName =
    String(
      cleanRow["Employee Name"] || ""
    ).trim();

  const branchCode =
    String(
      cleanRow["Br Code"] || ""
    ).trim();

  const branchName =
    String(
      cleanRow["Branch Name"] || ""
    ).trim();

  let designation =
  String(
    cleanRow["Designation"] || ""
  ).trim();

// Force Cluster Head designations
const clusterHeadMap = {
  "65": "Cluster Head- Krishna",
  "223": "Cluster Head- Visakhapatnam",
  "878": "Cluster Head- West Godavari",
  "446": "Cluster Head- Guntur"
};

if (clusterHeadMap[empId]) {
  designation = clusterHeadMap[empId];
}

  let contactNumber =
  String(cleanRow["Contact Number"] || "")
    .replace(/\s+/g, "")
    .replace(/[-()+]/g, "")
    .trim();

// Remove country code 91 if present
if (
  contactNumber.length > 10 &&
  contactNumber.startsWith("91")
) {
  contactNumber = contactNumber.substring(2);
}

  let cluster =
  String(
    cleanRow["Cluster"] || ""
  )
  .trim()
  .replace(/\s+/g, "")
  .toUpperCase();

const clusterMap = {
  K: "Krishna",
  G: "Guntur",
  V: "Visakhapatnam",
  W: "West Godavari",
  WG: "West Godavari"
};

cluster =
  clusterMap[cluster] || cluster;

        await transaction.request()

          .input(
  "empId",
  sql.Int,
  parseInt(empId)
)

          .input(
            "empName",
            sql.VarChar(200),
            empName
          )

          .input(
  "branchCode",
  sql.Int,
  parseInt(branchCode)
)

          .input(
            "branchName",
            sql.VarChar(200),
            branchName
          )

          .input(
            "designation",
            sql.VarChar(100),
            designation
          )

          .input(
            "contactNumber",
            sql.VarChar(20),
            contactNumber
          )

          .input(
            "cluster",
            sql.VarChar(100),
            cluster
          )

          .query(`
  INSERT INTO employees_master
  (
    [Emp No.],
    [Employee Name],
    [Br Code],
    [Branch Name],
    [Designation],
    [Mobile number],
    [Cluster]
  )
  VALUES
  (
    @empId,
    @empName,
    @branchCode,
    @branchName,
    @designation,
    @contactNumber,
    @cluster
  )
`);

        insertedCount++;

      }

      await transaction.commit();

await logActivity(
  userId,
  role,
  "Upload Employees Master",
  `Uploaded ${insertedCount} rows`
);

console.log(
  "EMPLOYEE_UPLOAD_SUCCESS",
  {
    employeeId: userId,
    uploadedRows: insertedCount
  }
);

      return res.json({
  success: true,
  message:
    "File Uploaded Successfully",
  insertedCount
});

    } catch (err) {

  if (transaction) {

    try {

      await transaction.rollback();

    } catch (rollbackError) {

      console.error(
        "ROLLBACK_FAILED",
        rollbackError.message
      );

    }

  }

  console.error(
    "EMPLOYEE_UPLOAD_FAILED",
    {
      employeeId: userId,
      error: err.message
    }
  );

  try {

    await logSecurityEvent(
      userId,
      "Employees Upload Failed",
      err.message
    );

  } catch (logErr) {

    console.error(
      "SECURITY_LOG_FAILED",
      logErr.message
    );

  }

return res.status(500).json({
    success: false,
    message: err.message
  });

}

  }
);

// ======================================================
// SYNC EMPLOYEES MASTER → AUTH
// ======================================================
app.post(
  "/sync-employees",
  authMiddleware,
  async (req, res) => {
    try {

      // Only admin can run
      if (req.user.role?.toLowerCase() !== "admin") {
        return res.status(403).json({
          success: false,
          message: "Access denied"
        });
      }

      console.log(
        "SYNC_EMPLOYEES_API_HIT",
        {
          employeeId: req.user.employeeId
        }
      );

      await queryUTIDatabase(`
        EXEC sync_employees_master_and_auth
      `);

      await logActivity(
        req.user.employeeId,
        req.user.role,
        "Sync Employees",
        "Executed sync_employees_master_and_auth"
      );

      console.log(
        "SYNC_EMPLOYEES_SUCCESS",
        {
          employeeId: req.user.employeeId
        }
      );

      return res.json({
        success: true,
        message: "Employees synchronized successfully"
      });

    } catch (err) {

      console.error(
        "SYNC_EMPLOYEES_FAILED",
        {
          employeeId: req.user?.employeeId,
          error: err.message
        }
      );

      return res.status(500).json({
        success: false,
        message: "Sync failed",
        error: err.message
      });
    }
  }
);

app.use((err, req, res, next) => {

  console.error(
    "SERVER_ERROR",
    {
      message: err.message
    }
  );

  if (err instanceof multer.MulterError) {
    return res.status(400).json({
      success:false,
      message:"File upload failed"
    });
  }

  if (
    err.message === "Only CSV and XLSX files are allowed" ||
    err.message === "Invalid file MIME type"
  ) {
    return res.status(400).json({
      success:false,
      message:err.message
    });
  }

  return res.status(500).json({
    success:false,
    message:"Internal server error"
  });

});



// ==========================================================================================================================================================
//                                                                        DASHBOARD
// ==========================================================================================================================================================


app.post("/dashboard-login", async (req, res) => {
  try {

    const {
  employeeId,
  password,
  googleCode,
  forceLogin = false
} = req.body;

    const rows = await queryUTIDatabase(
      `
SELECT
    m.[Emp No.],
    m.[Employee Name],
    m.[Br Code],
    m.[Branch Name],
    m.[Cluster],
    m.[Designation],

    a.[Password],
    a.[role],
    a.[Level],
    a.[Approval status],
    a.[GA_Secret],

    ISNULL(a.DashboardSessionVersion, 0)
      AS DashboardSessionVersion

FROM employees_auth a
INNER JOIN employees_master m
ON a.[Emp No.] = m.[Emp No.]

WHERE a.[Emp No.] = ?
      `,
      [employeeId]
    );

    if (!rows.length) {
      return res.status(404).json({
        success: false,
        message: "User not found"
      });
    }

    const user = rows[0];

    // Approval Check
    if (
      user["Approval status"] !==
      "approved"
    ) {
      return res.status(403).json({
        success: false,
        message: "Account not approved"
      });
    }

    // Password Check
    if (
      user.Password !== password
    ) {
      return res.status(401).json({
        success: false,
        message: "Invalid password"
      });
    }

    // Google Authenticator Validation
if (!user.GA_Secret) {
  return res.status(400).json({
    success: false,
    message: "Google Authenticator is not configured."
  });
}

if (!googleCode) {
  return res.status(400).json({
    success: false,
    message: "Google Authenticator code is required."
  });
}

const result = speakeasy.totp.verifyDelta({
  secret: user.GA_Secret,
  encoding: "base32",
  token: googleCode,
  window: 0
});

if (!result || result.delta !== 0) {
  return res.status(401).json({
    success: false,
    message: "Google Authenticator code has expired or is invalid."
  });
}

    console.log(
      "Dashboard Login",
      {
        employeeId,
        forceLogin,
        sessionVersion:
          user.DashboardSessionVersion
      }
    );

    // Already Logged In
    if (
      user.DashboardSessionVersion > 0 &&
      !forceLogin
    ) {

      return res.status(200).json({
        success: false,
        alreadyLoggedIn: true,
        message:
          "This account is already logged in on another device. Do you want to continue?"
      });

    }

    // Create New Session Version
    const newVersion =
      user.DashboardSessionVersion + 1;

    await queryUTIDatabase(
      `
      UPDATE employees_auth
      SET
        DashboardSessionVersion = ?,
        DashboardLastLogin = GETDATE()
      WHERE [Emp No.] = ?
      `,
      [
        newVersion,
        employeeId
      ]
    );

    const token = jwt.sign(
      {
        employeeId:
          user["Emp No."].toString(),

        role:
          user.role,

        level:
          user.Level,

        branchCode:
          user["Br Code"],

        designation:
          user.Designation,

        dashboard: true,

        sessionVersion:
          newVersion
      },
      process.env.JWT_SECRET,
      {
        expiresIn: "8h"
      }
    );

    return res.status(200).json({
  success: true,
  token,

  userId: user["Emp No."],

  role: user.role,

  employeeName: user["Employee Name"],

  branchCode: user["Br Code"],

  branchName: user["Branch Name"],

  clusterName: user["Cluster"],

  designation: user.Designation,

  level: user.Level,

  message: `Welcome ${user["Employee Name"]}`
});

  } catch (err) {

    console.error(
      "Dashboard Login Error:",
      err
    );

    return res.status(500).json({
      success: false,
      message: "Server Error"
    });

  }
});


app.post(
  "/dashboard-logout",
  async (req, res) => {

    try {

      const authHeader =
        req.headers.authorization;

      if (!authHeader) {
        return res.status(401).json({
          message: "Login required"
        });
      }

      const token =
        authHeader.split(" ")[1];

      const decoded =
        jwt.verify(
          token,
          process.env.JWT_SECRET
        );

      await queryUTIDatabase(
        `
        UPDATE employees_auth
        SET DashboardSessionVersion = 0
        WHERE [Emp No.] = ?
        `,
        [
          decoded.employeeId
        ]
      );

      console.log(
        "DASHBOARD_LOGOUT_SUCCESS",
        decoded.employeeId
      );

      return res.json({
        success: true
      });

    } catch (err) {

      console.error(
        "DASHBOARD_LOGOUT_FAILED",
        err
      );

      return res.status(401).json({
        message: "Logout failed"
      });

    }
  }
);


//------------------------------------------------------
// At A Glance Export Excel
//------------------------------------------------------
app.post("/api/glance/export", authMiddleware, async (req, res) => {

    let {
        level,
        branch_code,
        branch_name,
        cluster_name
    } = req.body;

    try {

        //-----------------------------------------
        // Apply Existing Role Restriction
        //-----------------------------------------

        const r = applyLevel1Restriction({

            userLevel: req.user.level,

            requestedLevel: level,

            branch_code,

            branch_name,

            cluster_name,

            designation: req.user.designation

        });

        console.log(
            "GLANCE_EXPORT_STARTED",
            {
                employeeId: req.user.employeeId,
                role: req.user.role,
                level: r.level,
                branch: r.branch_code,
                cluster: r.cluster_name
            }
        );

        //-----------------------------------------
        // Database Connection
        //-----------------------------------------

        const pool = await poolPromise;

        //-----------------------------------------
        // Create Excel Workbook
        //-----------------------------------------

        const workbook = new ExcelJS.Workbook();
		
		//------------------------------------------------------
// EXCEL FORMAT HELPERS
//------------------------------------------------------

function formatNumber(value) {

    if (
        value === null ||
        value === undefined ||
        value === "" ||
        value === "-"
    ) {
        return "-";
    }

    const num =
Number(String(value).replace(/,/g,""));

if(isNaN(num) || num===0){
    return "-";
}

    if (isNaN(num)) {
        return value;
    }

    return num.toLocaleString("en-IN", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });

}

function formatAccounts(value) {

    if (
        value === null ||
        value === undefined ||
        value === "" ||
        value === "-"
    ) {
        return "-";
    }

    const num = parseInt(value, 10);
    if(isNaN(num) || num===0){
    return "-";
}

    if (isNaN(num)) {
        return "-";
    }

    return num.toLocaleString("en-IN");

}

function formatSmaAmount(value){

    if(
        value===null ||
        value===undefined ||
        value===""){
        return "-";
    }

    const num =
    Number(String(value).replace(/,/g,""));

    if(isNaN(num) || num===0){
    return "-";
}

    if(isNaN(num))
        return "-";

    return (num/100000)
    .toLocaleString("en-IN",{

        minimumFractionDigits:2,
        maximumFractionDigits:2

    });

}

function formatChangeOverPdBalance(value){

    if(
        value===null||
        value===undefined||
        value===""
    ){
        return "-";
    }

    let num=Math.abs(
        Number(
            String(value).replace(/,/g,"")
        )
    );

    if(isNaN(num) || num===0){
    return "-";
}

    if(isNaN(num)){
        return "-";
    }

    return (num/100000).toLocaleString(
        "en-IN",
        {
            minimumFractionDigits:2,
            maximumFractionDigits:2
        }
    );
}

function isNegative(value){

    const num =
    Number(String(value).replace(/,/g,""));

    return !isNaN(num) && num<0;

}

function isPositive(value){

    const num =
    Number(String(value).replace(/,/g,""));

    return !isNaN(num) && num>0;

}

        workbook.creator = "Coastal Local Area Bank Ltd.";

        workbook.created = new Date();

//------------------------------------------------------
// Execute Main Stored Procedure
//------------------------------------------------------

const request =
    pool.request();

request.input(
    "level",
    sql.VarChar(10),
    r.level
);

request.input(
    "branch_code",
    sql.Int,
    r.branch_code ?? null
);

request.input(
    "branch_name",
    sql.VarChar(100),
    r.branch_name ?? null
);

request.input(
    "cluster_name",
    sql.VarChar(50),
    r.level === "CLUSTER"
        ? r.cluster_name
        : null
);

const result =
    await request.execute(
        "get_glance_data"
    );

const card1 =
    result.recordsets?.[0]?.[0] || {};

const card2 =
    result.recordsets?.[1]?.[0] || {};
	
	
	//------------------------------------------------------
// Information Sheet
//------------------------------------------------------

const infoSheet =
workbook.addWorksheet(
    r.level === "BRANCH"
        ? "Branch Information"
        : "Cluster Information"
);

infoSheet.columns = [

{header:"Field",key:"field",width:30},

{header:"Value",key:"value",width:40}

];

if(r.level==="BRANCH"){

infoSheet.addRows([

{
field:"Branch Code",
value:card1.branch_code
},

{
field:"Branch Name",
value:card1.branch_name
},

{
field:"Cluster",
value:card1.cluster_name
},

{
field:"Branch Manager",
value:card1.branch_manager_name
},

{
field:"BM Mobile",
value:card1.branch_manager_contact
},

{
field:"No Of Staff",
value:card1.branch_staff_count
}

]);

}
else{

infoSheet.addRows([

{
field:"Cluster",
value:card1.cluster_name
},

{
field:"Cluster Head",
value:card1.cluster_head_name
},

{
field:"Mobile",
value:card1.cluster_head_contact
},

{
field:"Branches",
value:card1.cluster_branch_count
},

{
field:"Staff",
value:card1.cluster_staff_count
}

]);

}

//------------------------------------------------------
// RBIA AUDIT REPORT
// BRANCH LEVEL ONLY
//------------------------------------------------------

if (r.level === "BRANCH") {

    //------------------------------------------------------
    // FETCH RBIA AUDIT DATA
    //------------------------------------------------------

    const rbiaRequest = pool.request();

    rbiaRequest.input(
        "level",
        sql.VarChar(10),
        r.level
    );

    rbiaRequest.input(
        "branch_code",
        sql.Int,
        r.branch_code ?? null
    );

    rbiaRequest.input(
        "branch_name",
        sql.VarChar(100),
        r.branch_name ?? null
    );

    rbiaRequest.input(
        "cluster_name",
        sql.VarChar(50),
        null
    );

    const rbiaResult =
        await rbiaRequest.execute(
            "get_glance_rbia_audit"
        );

    const rbiaData =
        rbiaResult.recordset || [];

    //------------------------------------------------------
    // CREATE RBIA AUDIT SHEET
    //------------------------------------------------------

    if (rbiaData.length > 0) {

        const rbiaSheet =
            workbook.addWorksheet(
                "RBIA Audit Report"
            );

        //------------------------------------------------------
        // TITLE
        //------------------------------------------------------

        rbiaSheet.mergeCells("A1:B1");

        const rbiaTitle =
            rbiaSheet.getCell("A1");

        rbiaTitle.value =
            "RBIA Audit Report";

        rbiaTitle.font = {
            bold: true,
            size: 16,
            color: {
                argb: "FFFFFFFF"
            }
        };

        rbiaTitle.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: {
                argb: "1E40AF"
            }
        };

        rbiaTitle.alignment = {
            horizontal: "center",
            vertical: "middle"
        };

        rbiaSheet.getRow(1).height = 24;

        //------------------------------------------------------
        // HEADER
        //------------------------------------------------------

        rbiaSheet.addRow([
            "Particulars",
            "Details"
        ]);

        const rbiaHeader =
            rbiaSheet.getRow(2);

        rbiaHeader.font = {
            bold: true,
            color: {
                argb: "FFFFFFFF"
            }
        };

        rbiaHeader.eachCell(cell => {

            cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: {
                    argb: "1E40AF"
                }
            };

            cell.alignment = {
                horizontal: "center",
                vertical: "middle"
            };

        });

        //------------------------------------------------------
        // DATA
        //------------------------------------------------------

        rbiaData.forEach(row => {

            //--------------------------------------------------
            // FORMAT AUDIT DATE
            //--------------------------------------------------

            let auditDate =
                row["Date of Audit"];

            if (auditDate) {

                const date =
                    new Date(auditDate);

                if (!isNaN(date.getTime())) {

                    auditDate =
                        date.toLocaleDateString(
                            "en-GB"
                        );

                }

            } else {

                auditDate = "-";

            }

            //--------------------------------------------------
            // BRANCH NAME / CODE FALLBACK
            //--------------------------------------------------

            const branchName =
                row["Name of the branch"] ||
                row["Name of the Branch"] ||
                row.branch_name ||
                card1.branch_name ||
                "-";

            const branchCode =
                row["Branch Code"] ??
                row.branch_code ??
                card1.branch_code ??
                r.branch_code ??
                "-";

            //--------------------------------------------------
            // RBIA ROWS
            //--------------------------------------------------

            const rbiaRows = [

                [
                    "Name of the Branch",
                    branchName
                ],

                [
                    "Branch Code",
                    branchCode
                ],

                [
                    "Date of Audit",
                    auditDate
                ],

                [
                    "Business Risk",
                    row["Business Risk"] ?? "-"
                ],

                [
                    "Credit Risk",
                    row["Credit Risk"] ?? "-"
                ],

                [
                    "Operational Risk",
                    row["Operational Risk"] ?? "-"
                ],

                [
                    "Overall Score",
                    row["Overall Score"] ?? "-"
                ],

                [
                    "Rating",
                    row["Rating"] || "-"
                ]

            ];

            //--------------------------------------------------
            // ADD ROWS TO SHEET
            //--------------------------------------------------

            rbiaRows.forEach(item => {

                const excelRow =
                    rbiaSheet.addRow(item);

                // Particulars
                excelRow.getCell(1).alignment = {
                    horizontal: "left",
                    vertical: "middle"
                };

                excelRow.getCell(1).font = {
                    bold: true
                };

                // Details
                excelRow.getCell(2).alignment = {
                    horizontal: "right",
                    vertical: "middle"
                };

                excelRow.getCell(2).font = {
                    bold: true
                };

            });

        });

        //------------------------------------------------------
        // COLUMN WIDTHS
        //------------------------------------------------------

        rbiaSheet.getColumn(1).width = 30;

        rbiaSheet.getColumn(2).width = 25;

        //------------------------------------------------------
        // BORDERS
        //------------------------------------------------------

        rbiaSheet.eachRow(row => {

            row.eachCell(cell => {

                cell.border = {

                    top: {
                        style: "thin"
                    },

                    left: {
                        style: "thin"
                    },

                    bottom: {
                        style: "thin"
                    },

                    right: {
                        style: "thin"
                    }

                };

            });

        });

    }

}

//------------------------------------------------------
// DEPOSITS & ADVANCES
//------------------------------------------------------

const depositsSheet =
workbook.addWorksheet("Deposits & Advances");

// Title

depositsSheet.mergeCells("A1:C1");

const depositsTitle =
depositsSheet.getCell("A1");

depositsTitle.value = "Deposits & Advances";

depositsTitle.font = {
    bold: true,
    size: 16,
    color: {
        argb: "FFFFFFFF"
    }
};

depositsTitle.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: {
        argb: "1E40AF"
    }
};

depositsTitle.alignment = {
    horizontal: "center",
    vertical: "middle"
};

depositsSheet.getRow(1).height = 24;

//------------------------------------------------------
// Header
//------------------------------------------------------

depositsSheet.addRow([
    "Sections",
    "Deposits",
    "Advances"
]);

const header =
depositsSheet.getRow(2);

header.height = 22;

header.font = {
    bold: true,
    color: {
        argb: "FFFFFFFF"
    }
};

header.eachCell(cell=>{

cell.fill={

type:"pattern",

pattern:"solid",

fgColor:{argb:"1E40AF"}

};

});

header.alignment = {
    horizontal: "center",
    vertical: "middle"
};

//------------------------------------------------------
// Data
//------------------------------------------------------

const depositRows = [

{
section:"As On Date",
deposit:card2.as_on_date_deposits,
advance:card2.as_on_date_advances
},

{
section:"Change Over PD",
deposit:card2.change_over_pd_deposits,
advance:card2.change_over_pd_advances
},

{
section:"GDM",
deposit:card2.gdm_deposits,
advance:card2.gdm_advances
},

{
section:"GUM",
deposit:card2.gum_deposits,
advance:card2.gum_advances
},

{
section:"Budget %",
deposit:card2.budget_achieved_deposits,
advance:card2.budget_achieved_advances
}

];

depositRows.forEach(r=>{

const row =
depositsSheet.addRow([

r.section,

formatNumber(r.deposit),

formatNumber(r.advance)

]);

row.getCell(1).alignment={
horizontal:"center"
};

row.getCell(2).alignment={
horizontal:"right"
};

row.getCell(3).alignment={
horizontal:"right"
};

// Deposit Color

if(Number(r.deposit)<0){

row.getCell(2).font={

color:{
argb:"DC2626"
},

bold:true

};

}else{

row.getCell(2).font={

color:{
argb:"111827"
},

bold:true

};

}

// Advance Color

if(Number(r.advance)<0){

row.getCell(3).font={

color:{
argb:"DC2626"
},

bold:true

};

}else{

row.getCell(3).font={

color:{
argb:"111827"
},

bold:true

};

}

});

//------------------------------------------------------
// Column Width
//------------------------------------------------------

depositsSheet.getColumn(1).width=28;

depositsSheet.getColumn(2).width=20;

depositsSheet.getColumn(3).width=20;

//------------------------------------------------------
// Borders
//------------------------------------------------------

depositsSheet.eachRow(row=>{

row.eachCell(cell=>{

cell.border={

top:{style:"thin"},

left:{style:"thin"},

bottom:{style:"thin"},

right:{style:"thin"}

};

});

});

//------------------------------------------------------
// NPA DATA
//------------------------------------------------------

const npaRequest = pool.request();

npaRequest.input("level", sql.VarChar(10), r.level);
npaRequest.input("branch_code", sql.VarChar(20), r.branch_code?.toString() ?? null);
npaRequest.input("branch_name", sql.VarChar(100), r.branch_name ?? null);
npaRequest.input(
    "cluster_name",
    sql.VarChar(50),
    r.level === "CLUSTER"
        ? r.cluster_name
        : null
);

const npaResult =
await npaRequest.execute(
    "get_glance_npa_data"
);

const npaData =
npaResult.recordset || [];

//------------------------------------------------------
// SMA DATA
//------------------------------------------------------

const smaRequest = pool.request();

smaRequest.input("level", sql.VarChar(10), r.level);
smaRequest.input("branch_code", sql.Int, r.branch_code ?? null);
smaRequest.input("branch_name", sql.VarChar(100), r.branch_name ?? null);

smaRequest.input(
    "cluster_name",
    sql.VarChar(50),
    r.level === "CLUSTER"
        ? r.cluster_name
        : null
);

const smaResult =
await smaRequest.execute(
    "get_glance_sma_data"
);

const smaData =
smaResult.recordset || [];

//------------------------------------------------------
// Convert NPA
//------------------------------------------------------

const card3 = {

    asOnAccounts:"-",
    asOnBalance:"-",

    pdAccounts:"-",
    pdBalance:"-",

    monthAccounts:"-",
    monthBalance:"-",

    prevAccounts:"-",
    prevBalance:"-",

    percentage:"-"

};

npaData.forEach(r=>{

    switch(r.section){

        case "As on Date":

            card3.asOnAccounts=r.npa_accounts;
            card3.asOnBalance=r.npa_balance;

            break;

        case "Change over PD":

            card3.pdAccounts=r.npa_accounts;
            card3.pdBalance=r.npa_balance;

            break;

        case "Month To Date":

            card3.monthAccounts=r.npa_accounts;
            card3.monthBalance=r.npa_balance;

            break;

        case "Previous Year":

            card3.prevAccounts=r.npa_accounts;
            card3.prevBalance=r.npa_balance;

            break;

        case "% of Advances (As on Date)":

            card3.percentage=r.npa_balance;

            break;

    }

});

//------------------------------------------------------
// Convert SMA
//------------------------------------------------------

const sma = {

    asOnAccounts:"-",
    asOnBalance:"-",

    pdAccounts:"-",
    pdBalance:"-",

    monthAccounts:"-",
    monthBalance:"-",

    prevAccounts:"-",
    prevBalance:"-",

    percentage:"-"

};

smaData.forEach(r=>{

    switch(r.section){

        case "As on Date":

            sma.asOnAccounts=r.acc;
            sma.asOnBalance=r.bal;

            break;

        case "Change over PD":

            sma.pdAccounts=r.acc;
            sma.pdBalance=r.bal;

            break;

        case "Month To Date":

            sma.monthAccounts=r.acc;
            sma.monthBalance=r.bal;

            break;

        case "Previous Year":

            sma.prevAccounts=r.acc;
            sma.prevBalance=r.bal;

            break;

        case "% of Advances (As on Date)":

            sma.percentage=r.bal;

            break;

    }

});

//------------------------------------------------------
// NPA & SMA SHEET
//------------------------------------------------------

const npaSheet =
workbook.addWorksheet("NPA & SMA");

//-----------------------------------------
// TITLE
//-----------------------------------------

npaSheet.mergeCells("A1:E1");

const npaTitle =
npaSheet.getCell("A1");

npaTitle.value = "NPA & SMA";

npaTitle.font = {
    bold: true,
    size: 16,
    color: { argb: "FFFFFFFF" }
};

npaTitle.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "1E40AF" }
};

npaTitle.alignment = {
    horizontal: "center",
    vertical: "middle"
};

//-----------------------------------------
// HEADER
//-----------------------------------------

npaSheet.mergeCells("B2:C2");
npaSheet.mergeCells("D2:E2");

npaSheet.getCell("A2").value = "Sections";
npaSheet.getCell("B2").value = "NPA";
npaSheet.getCell("D2").value = "SMA (1+2)";

["A2","B2","D2"].forEach(c=>{

const cell=npaSheet.getCell(c);

cell.font={
bold:true,
color:{argb:"FFFFFFFF"}
};

cell.fill={
type:"pattern",
pattern:"solid",
fgColor:{argb:"1E40AF"}
};

cell.alignment={
horizontal:"center",
vertical:"middle"
};

});

npaSheet.getRow(3).values=[
"",
"A/Cs",
"Balance",
"A/Cs",
"Balance"
];

npaSheet.getRow(3).eachCell(cell=>{

cell.font={
bold:true,
color:{argb:"FFFFFFFF"}
};

cell.fill={
type:"pattern",
pattern:"solid",
fgColor:{argb:"1E40AF"}
};

cell.alignment={
horizontal:"center"
};

});

//-----------------------------------------
// DATA
//-----------------------------------------

const rows=[

{
section:"As on Date",
npaAcs:card3.asOnAccounts,
npaBal:card3.asOnBalance,
smaAcs:sma.asOnAccounts,
smaBal:sma.asOnBalance
},

{
section:"Change over PD",
npaAcs:card3.pdAccounts,
npaBal:card3.pdBalance,
smaAcs:sma.pdAccounts,
smaBal:sma.pdBalance
},

{
section:"Month To Date",
npaAcs:card3.monthAccounts,
npaBal:card3.monthBalance,
smaAcs:sma.monthAccounts,
smaBal:sma.monthBalance
},

{
section:"Previous Year",
npaAcs:card3.prevAccounts,
npaBal:card3.prevBalance,
smaAcs:sma.prevAccounts,
smaBal:sma.prevBalance
}

];

rows.forEach(r=>{

const isChangeOverPD = r.section === "Change over PD";

const row = npaSheet.addRow([

    r.section,

    Number(r.npaAcs || 0),

    Number(r.npaBal || 0),

    Number(r.smaAcs || 0),

    Number(r.smaBal || 0)

]);

if (isChangeOverPD) {

    row.getCell(2).value = Math.abs(Number(r.npaAcs || 0));

    row.getCell(3).value =
        Number((Math.abs(Number(r.npaBal || 0)) / 100000).toFixed(2));

    row.getCell(4).value = Math.abs(Number(r.smaAcs || 0));

    row.getCell(5).value =
        Number((Math.abs(Number(r.smaBal || 0)) / 100000).toFixed(2));

}
else {

    row.getCell(5).value =
        Number((Number(r.smaBal || 0) / 100000).toFixed(2));

}

row.getCell(1).alignment={horizontal:"center"};

for(let i=2;i<=5;i++){

row.getCell(i).alignment={
horizontal:"right"
};

}

// NPA Balance
row.getCell(3).numFmt = '#,##0.00';

// SMA Balance
row.getCell(5).numFmt = '#,##0.00';

const pdValues = {
    2: Number(r.npaAcs || 0),
    3: Number(r.npaBal || 0),
    4: Number(r.smaAcs || 0),
    5: Number(r.smaBal || 0)
};

[2,3,4,5].forEach(col => {

    const value = pdValues[col];

    let color = "111827";

    if (isChangeOverPD) {

        if (value > 0) {
            color = "DC2626";      // RED (same as frontend)
        }
        else if (value < 0) {
            color = "16A34A";      // GREEN (same as frontend)
        }
        else {
            color = "111827";
        }

    }

    row.getCell(col).font = {
        bold: true,
        name: "Calibri",
        size: 11,
        color: {
            argb: color
        }
    };

});

if (r.section === "Change over PD") {

    row.getCell(2).value = Math.abs(Number(r.npaAcs || 0));
    row.getCell(3).value =
    Math.abs(Number(r.npaBal || 0)) < 1
        ? 0
        : Number(
            (
                Math.abs(Number(r.npaBal)) /
                100000
            ).toFixed(2)
        );

    row.getCell(4).value = Math.abs(Number(r.smaAcs || 0));
    row.getCell(5).value =
    Math.abs(Number(r.smaBal || 0)) < 1
        ? 0
        : Number(
            (
                Math.abs(Number(r.smaBal)) /
                100000
            ).toFixed(2)
        );

}

});

//-----------------------------------------
// % OF ADVANCES
//-----------------------------------------

const pctRow=
npaSheet.addRow([]);

pctRow.getCell(1).value="% of Advances (As on Date)";

npaSheet.mergeCells(`B${pctRow.number}:C${pctRow.number}`);
npaSheet.mergeCells(`D${pctRow.number}:E${pctRow.number}`);

pctRow.getCell(2).value =
`${formatNumber(card3.percentage)}%`;

pctRow.getCell(4).value =
`${formatNumber(sma.percentage)}%`;

pctRow.getCell(2).alignment={
horizontal:"center"
};

pctRow.getCell(4).alignment={
horizontal:"center"
};

pctRow.getCell(2).font={
bold:true
};

pctRow.getCell(4).font={
bold:true
};

//-----------------------------------------
// WIDTHS
//-----------------------------------------

npaSheet.getColumn(1).width=30;
npaSheet.getColumn(2).width=14;
npaSheet.getColumn(3).width=18;
npaSheet.getColumn(4).width=14;
npaSheet.getColumn(5).width=18;

//-----------------------------------------
// BORDERS
//-----------------------------------------

npaSheet.eachRow(row=>{

row.eachCell(cell=>{

cell.border={

top:{style:"thin"},

left:{style:"thin"},

bottom:{style:"thin"},

right:{style:"thin"}

};

});

});

//------------------------------------------------------
// DEPOSIT ACCOUNTS OPENED
//------------------------------------------------------

const doaRequest = pool.request();

doaRequest.input(
    "level",
    sql.VarChar(10),
    r.level
);

doaRequest.input(
    "branch_code",
    sql.Int,
    r.branch_code ?? null
);

doaRequest.input(
    "branch_name",
    sql.NVarChar(100),
    r.branch_name ?? null
);

doaRequest.input(
    "cluster_name",
    sql.VarChar(50),
    r.level === "CLUSTER"
        ? r.cluster_name
        : null
);

const doaResult =
await doaRequest.execute(
    "get_glance_doa_card4"
);

const doaRows =
doaResult.recordsets
? doaResult.recordsets.flat()
: doaResult.recordset || [];

//------------------------------------------------------
// TRANSFORM DOA
//------------------------------------------------------

const depositData = {

TDR:{type:"TDR"},

RD:{type:"RD"},

Savings:{type:"Savings Account"},

Current:{type:"Current Account"}

};

doaRows.forEach(row=>{

if(row.section==="As on Date"){

depositData.TDR.odAcs=row.tdr_acs;
depositData.TDR.odAmt=row.tdr_amt;

depositData.RD.odAcs=row.rd_acs;
depositData.RD.odAmt=row.rd_amt;

depositData.Savings.odAcs=row.sb_acs;
depositData.Savings.odAmt=row.sb_amt;

depositData.Current.odAcs=row.ca_acs;
depositData.Current.odAmt=row.ca_amt;

}

if(row.section==="Month To Date"){

depositData.TDR.mtdAcs=row.tdr_acs;
depositData.TDR.mtdAmt=row.tdr_amt;

depositData.RD.mtdAcs=row.rd_acs;
depositData.RD.mtdAmt=row.rd_amt;

depositData.Savings.mtdAcs=row.sb_acs;
depositData.Savings.mtdAmt=row.sb_amt;

depositData.Current.mtdAcs=row.ca_acs;
depositData.Current.mtdAmt=row.ca_amt;

}

if(row.section==="Previous Year"){

depositData.TDR.ytdAcs=row.tdr_acs;
depositData.TDR.ytdAmt=row.tdr_amt;

depositData.RD.ytdAcs=row.rd_acs;
depositData.RD.ytdAmt=row.rd_amt;

depositData.Savings.ytdAcs=row.sb_acs;
depositData.Savings.ytdAmt=row.sb_amt;

depositData.Current.ytdAcs=row.ca_acs;
depositData.Current.ytdAmt=row.ca_amt;

}

});

//------------------------------------------------------
// DEPOSIT ACCOUNTS SHEET
//------------------------------------------------------

const depositSheet =
workbook.addWorksheet("Deposit Accounts Opened");

//--------------------------------------
// TITLE
//--------------------------------------

depositSheet.mergeCells("A1:G1");

const depositTitle =
depositSheet.getCell("A1");

depositTitle.value =
"Deposit Accounts Opened";

depositTitle.font = {
    bold: true,
    size: 16,
    color: {
        argb: "FFFFFFFF"
    }
};

depositTitle.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: {
        argb: "1E40AF"
    }
};

depositTitle.alignment = {
    horizontal: "center",
    vertical: "middle"
};

//--------------------------------------
// HEADER ROW 1
//--------------------------------------

depositSheet.mergeCells("A2:A3");
depositSheet.mergeCells("B2:C2");
depositSheet.mergeCells("D2:E2");
depositSheet.mergeCells("F2:G2");

depositSheet.getCell("A2").value = "Section";
depositSheet.getCell("B2").value = "On Date";
depositSheet.getCell("D2").value = "Month To Date";
depositSheet.getCell("F2").value = "Year To Date";

["A2","B2","D2","F2"].forEach(c=>{

const cell=
depositSheet.getCell(c);

cell.font={
bold:true,
color:{argb:"FFFFFFFF"}
};

cell.fill={
type:"pattern",
pattern:"solid",
fgColor:{argb:"1E40AF"}
};

cell.alignment={
horizontal:"center",
vertical:"middle"
};

});

//--------------------------------------
// HEADER ROW 2
//--------------------------------------

depositSheet.getRow(3).values=[

"",
"A/Cs",
"Amount",
"A/Cs",
"Amount",
"A/Cs",
"Amount"

];

depositSheet.getRow(3).eachCell(cell=>{

cell.font={

bold:true,

color:{
argb:"FFFFFFFF"
}

};

cell.fill={

type:"pattern",

pattern:"solid",

fgColor:{
argb:"1E40AF"
}

};

cell.alignment={

horizontal:"center"

};

});

//--------------------------------------
// DATA
//--------------------------------------

Object.values(depositData).forEach(r=>{

const row=
depositSheet.addRow([

r.type,

formatAccounts(r.odAcs),

formatNumber(r.odAmt),

formatAccounts(r.mtdAcs),

formatNumber(r.mtdAmt),

formatAccounts(r.ytdAcs),

formatNumber(r.ytdAmt)

]);

// Alignment

row.getCell(1).alignment={
horizontal:"center"
};

for(let i=2;i<=7;i++){

row.getCell(i).alignment={
horizontal:"right"
};

}

});

//--------------------------------------
// WIDTH
//--------------------------------------

depositSheet.getColumn(1).width=28;

depositSheet.getColumn(2).width=14;
depositSheet.getColumn(3).width=18;

depositSheet.getColumn(4).width=14;
depositSheet.getColumn(5).width=18;

depositSheet.getColumn(6).width=14;
depositSheet.getColumn(7).width=18;

//--------------------------------------
// BORDERS
//--------------------------------------

depositSheet.eachRow(row=>{

row.eachCell(cell=>{

cell.border={

top:{style:"thin"},

left:{style:"thin"},

bottom:{style:"thin"},

right:{style:"thin"}

};

});

});

//------------------------------------------------------
// LOAN ACCOUNTS
//------------------------------------------------------

const loanRequest = pool.request();

loanRequest.input(
"level",
sql.VarChar(10),
r.level
);

loanRequest.input(
"branch_code",
sql.Int,
r.branch_code ?? null
);

loanRequest.input(
"branch_name",
sql.NVarChar(100),
r.branch_name ?? null
);

loanRequest.input(
"cluster_name",
sql.VarChar(50),
r.level==="CLUSTER"
? r.cluster_name
: null
);

const loanResult =
await loanRequest.execute(
"get_glance_loans_card5"
);

const loanRows =
loanResult.recordset || [];

//------------------------------------------------------
// TRANSFORM LOANS
//------------------------------------------------------

const loanMap = {};

loanRows.forEach(row=>{

if(!loanMap[row.loan_group]){

loanMap[row.loan_group]={

loanType:row.loan_group

};

}

if(row.period_type==="ON_DATE"){

loanMap[row.loan_group].odAcs=row.acs;
loanMap[row.loan_group].odAmt=row.amount;

}

if(row.period_type==="MONTH"){

loanMap[row.loan_group].mtdAcs=row.acs;
loanMap[row.loan_group].mtdAmt=row.amount;

}

if(row.period_type==="YEAR"){

loanMap[row.loan_group].ytdAcs=row.acs;
loanMap[row.loan_group].ytdAmt=row.amount;

}

});

//------------------------------------------------------
// LOAN ACCOUNTS OPENED SHEET
//------------------------------------------------------

const loanSheet =
workbook.addWorksheet("Loan Accounts Opened");

//--------------------------------------
// TITLE
//--------------------------------------

loanSheet.mergeCells("A1:G1");

const loanTitle =
loanSheet.getCell("A1");

loanTitle.value =
"Loan Accounts Opened";

loanTitle.font = {
    bold: true,
    size: 16,
    color: {
        argb: "FFFFFFFF"
    }
};

loanTitle.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: {
        argb: "1E40AF"
    }
};

loanTitle.alignment = {
    horizontal: "center",
    vertical: "middle"
};

//--------------------------------------
// HEADER
//--------------------------------------

loanSheet.mergeCells("A2:A3");
loanSheet.mergeCells("B2:C2");
loanSheet.mergeCells("D2:E2");
loanSheet.mergeCells("F2:G2");

loanSheet.getCell("A2").value="Loan Type";
loanSheet.getCell("B2").value="On Date";
loanSheet.getCell("D2").value="Month To Date";
loanSheet.getCell("F2").value="Year To Date";

["A2","B2","D2","F2"].forEach(c=>{

const cell=
loanSheet.getCell(c);

cell.font={
bold:true,
color:{
argb:"FFFFFFFF"
}
};

cell.fill={
type:"pattern",
pattern:"solid",
fgColor:{
argb:"1E40AF"
}
};

cell.alignment={
horizontal:"center",
vertical:"middle"
};

});

//--------------------------------------
// SUB HEADER
//--------------------------------------

loanSheet.getRow(3).values=[

"",
"A/Cs",
"Amount",
"A/Cs",
"Amount",
"A/Cs",
"Amount"

];

loanSheet.getRow(3).eachCell(cell=>{

cell.font={
bold:true,
color:{
argb:"FFFFFFFF"
}
};

cell.fill={
type:"pattern",
pattern:"solid",
fgColor:{
argb:"1E40AF"
}
};

cell.alignment={
horizontal:"center"
};

});

//--------------------------------------
// SORT LOAN GROUPS
//--------------------------------------

const LOAN_GROUP_ORDER=[

"GOLD LOANS",

"LOAN AGAINST DEPOSIT",

"PRAGATHI LOANS",

"TERM LOAN (NON-PRIORITY)",

"HOME LOANS",

"CRE",

"AGRI LOANS",

"VEHICLE LOANS",

"TERM LOAN (PRIORITY)",

"OTHERS"

];

const sortedLoans=
Object.values(loanMap).sort((a,b)=>{

const ia=
LOAN_GROUP_ORDER.indexOf(a.loanType);

const ib=
LOAN_GROUP_ORDER.indexOf(b.loanType);

return(
(ia===-1?999:ia)-
(ib===-1?999:ib)
);

});

//--------------------------------------
// DATA
//--------------------------------------

sortedLoans.forEach(r=>{

const row=
loanSheet.addRow([

r.loanType,

formatAccounts(r.odAcs),

formatNumber(r.odAmt),

formatAccounts(r.mtdAcs),

formatNumber(r.mtdAmt),

formatAccounts(r.ytdAcs),

formatNumber(r.ytdAmt)

]);

row.getCell(1).alignment={
horizontal:"center",
vertical:"middle",
wrapText:true
};

for(let i=2;i<=7;i++){

row.getCell(i).alignment={
horizontal:"right"
};

}

});

//--------------------------------------
// COLUMN WIDTHS
//--------------------------------------

loanSheet.getColumn(1).width=35;

loanSheet.getColumn(2).width=14;
loanSheet.getColumn(3).width=18;

loanSheet.getColumn(4).width=14;
loanSheet.getColumn(5).width=18;

loanSheet.getColumn(6).width=14;
loanSheet.getColumn(7).width=18;

//--------------------------------------
// BORDERS
//--------------------------------------

loanSheet.eachRow(row=>{

row.eachCell(cell=>{

cell.border={

top:{style:"thin"},

left:{style:"thin"},

bottom:{style:"thin"},

right:{style:"thin"}

};

});

});

//------------------------------------------------------
// BUSINESS CORRESPONDENTS
//------------------------------------------------------

const bcRequest = pool.request();

bcRequest.input("level", sql.VarChar(10), r.level);

bcRequest.input("branch_code", sql.Int, r.branch_code ?? null);

bcRequest.input("branch_name", sql.VarChar(100), r.branch_name ?? null);

bcRequest.input(
    "cluster_name",
    sql.VarChar(50),
    r.level === "CLUSTER"
        ? r.cluster_name
        : null
);

const bcResult =
await bcRequest.execute(
    "get_glance_business_correspondents"
);

const bcData =
bcResult.recordset?.[0] || {};

//------------------------------------------------------
// ATM COUNT
//------------------------------------------------------

const atmRequest = pool.request();

atmRequest.input("level", sql.VarChar(10), r.level);

atmRequest.input("branch_code", sql.Int, r.branch_code ?? null);

atmRequest.input("branch_name", sql.VarChar(100), r.branch_name ?? null);

atmRequest.input(
    "cluster_name",
    sql.VarChar(50),
    r.level==="CLUSTER"
        ? r.cluster_name
        : null
);

const atmResult =
await atmRequest.execute(
    "get_glance_atms"
);

const atmData =
atmResult.recordset?.[0] || {};

//------------------------------------------------------
// EMPLOYEE EFFICIENCY
//------------------------------------------------------

const empRequest = pool.request();

empRequest.input("level", sql.VarChar(10), r.level);

empRequest.input("branch_code", sql.Int, r.branch_code ?? null);

empRequest.input("branch_name", sql.VarChar(100), r.branch_name ?? null);

empRequest.input(
    "cluster_name",
    sql.VarChar(50),
    r.level==="CLUSTER"
        ? r.cluster_name
        : null
);

const empResult =
await empRequest.execute(
    "get_glance_employee_efficiency"
);

const empData =
empResult.recordset?.[0] || {};

//------------------------------------------------------
// EMPLOYEE EFFICIENCY SHEET
//------------------------------------------------------

const employeeSheet =
workbook.addWorksheet("Employee Efficiency");

//--------------------------------------
// TITLE
//--------------------------------------

employeeSheet.mergeCells("A1:C1");

const empTitle =
employeeSheet.getCell("A1");

empTitle.value =
"Employee Efficiency";

empTitle.font = {
    bold: true,
    size: 16,
    color: { argb: "FFFFFFFF" }
};

empTitle.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "1E40AF" }
};

empTitle.alignment = {
    horizontal: "center",
    vertical: "middle"
};

//--------------------------------------
// HEADER
//--------------------------------------

employeeSheet.addRow([
    "Section",
    "On Date",
    "Previous Year"
]);

const empHeader =
employeeSheet.getRow(2);

empHeader.font = {
    bold: true,
    color: { argb: "FFFFFFFF" }
};

empHeader.eachCell(cell=>{

cell.fill={

type:"pattern",

pattern:"solid",

fgColor:{argb:"1E40AF"}

};

});

empHeader.alignment = {
    horizontal: "center",
    vertical: "middle"
};

//--------------------------------------
// DATA
//--------------------------------------

employeeSheet.addRow([

"Number of Staff\n(Excluding Sub Staff)",

formatAccounts(empData.no_of_staff),

formatAccounts(empData.no_of_staff_previous_year)

]);

employeeSheet.addRow([

"Business Per Employee",

formatNumber(empData.business_per_employee),

formatNumber(empData.business_per_employee_previous_year)

]);

//--------------------------------------
// FORMATTING
//--------------------------------------

employeeSheet.getColumn(1).width = 40;
employeeSheet.getColumn(2).width = 20;
employeeSheet.getColumn(3).width = 20;

employeeSheet.eachRow((row,rowNumber)=>{

row.eachCell(cell=>{

cell.border={

top:{style:"thin"},

left:{style:"thin"},

bottom:{style:"thin"},

right:{style:"thin"}

};

if(rowNumber>=3){

if(cell.col===1){

cell.alignment={
horizontal:"center",
vertical:"middle",
wrapText:true
};

}else{

cell.alignment={
horizontal:"right"
};

}

}

});

});

//------------------------------------------------------
// VOUCHERS
//------------------------------------------------------

const voucherRequest = pool.request();

voucherRequest.input("level", sql.VarChar(10), r.level);

voucherRequest.input("branch_code", sql.Int, r.branch_code ?? null);

voucherRequest.input("branch_name", sql.VarChar(100), r.branch_name ?? null);

voucherRequest.input(
"cluster_name",
sql.VarChar(50),
r.level==="CLUSTER"
? r.cluster_name
: null
);

const voucherResult =
await voucherRequest.execute(
"get_glance_vouchers"
);

const voucher =
voucherResult.recordset?.[0] || {};

//------------------------------------------------------
// VOUCHERS SHEET
//------------------------------------------------------

const voucherSheet =
workbook.addWorksheet("Vouchers");

//--------------------------------------
// TITLE
//--------------------------------------

voucherSheet.mergeCells("A1:B1");

const voucherTitle =
voucherSheet.getCell("A1");

voucherTitle.value =
"Vouchers";

voucherTitle.font = {
    bold: true,
    size: 16,
    color: {
        argb:"FFFFFFFF"
    }
};

voucherTitle.fill = {
    type:"pattern",
    pattern:"solid",
    fgColor:{
        argb:"1E40AF"
    }
};

voucherTitle.alignment = {
    horizontal:"center",
    vertical:"middle"
};

//--------------------------------------
// HEADER
//--------------------------------------

voucherSheet.addRow([

"On Date",

"Average"

]);

const voucherHeader =
voucherSheet.getRow(2);

voucherHeader.font = {

bold:true,

color:{
argb:"FFFFFFFF"
}

};

voucherHeader.eachCell(cell=>{

cell.fill={

type:"pattern",

pattern:"solid",

fgColor:{argb:"1E40AF"}

};

});

voucherHeader.alignment = {

horizontal:"center"

};

//--------------------------------------
// DATA
//--------------------------------------

voucherSheet.addRow([

formatAccounts(voucher.vouchers),

formatAccounts(voucher.vouchers_average)

]);

//--------------------------------------
// FORMAT
//--------------------------------------

voucherSheet.getColumn(1).width=20;
voucherSheet.getColumn(2).width=20;

voucherSheet.eachRow((row,rowNumber)=>{

row.eachCell(cell=>{

cell.border={

top:{style:"thin"},

left:{style:"thin"},

bottom:{style:"thin"},

right:{style:"thin"}

};

if(rowNumber>=3){

cell.alignment={
horizontal:"center"
};

}

});

});

        //-----------------------------------------
        // Return Excel
        //-----------------------------------------

        res.setHeader(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        );

        res.setHeader(
            "Content-Disposition",
            "attachment; filename=At_A_Glance.xlsx"
        );

        await workbook.xlsx.write(res);

        res.end();

    }

    catch (err) {

        console.error(
            "GLANCE_EXPORT_FAILED",
            err
        );

        res.status(500).json({

            success: false,

            message: err.message

        });

    }

});



app.get(
  "/dashboard-session-check",
  dashboardAuthMiddleware,
  (req, res) => {

    res.json({
      success: true
    });

  }
);







// ======================
// FRONTEND LOGGING API
// ======================

app.post("/api/frontend-log", (req, res) => {

  console.log(
    "FRONTEND_LOG_API_HIT",
    req.body
  );

  try {

    const {
      source,
      level,
      message
    } = req.body;

    if (!message) {

      console.warn(
        "FRONTEND_LOG_EMPTY_MESSAGE"
      );

      return res.status(400).json({
        success: false
      });

    }

    writeDailyLog(
      `frontend/${source}`,
      `[${level}] ${message}`
    );

    console.log(
      "FRONTEND_LOG_WRITTEN"
    );

    return res.json({
      success: true
    });

  } catch (err) {

    console.error(
      "FRONTEND_LOG_ERROR",
      err
    );

    return res.status(500).json({
      success: false
    });

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


