// importSMA.js
const fs = require("fs");
const path = require("path");
const csv = require("csv-parser");
const mysql = require("mysql2/promise");

const csvFilePath = path.join(__dirname, "uploads", "SMA.csv");

// Normalize header keys (remove BOM, trim, lowercase)
function normalizeKey(header) {
  return header.replace(/^\uFEFF/, "").trim().toLowerCase();
}

// Clean branch code: remove unwanted leading zeros, ensure string
function cleanBranchCode(code) {
  if (!code) return "";
  return code.toString().trim().replace(/^0+/, ''); // Remove leading zeros
}

async function importSMA() {
  try {
    if (!fs.existsSync(csvFilePath)) {
      console.error("❌ CSV file not found at:", csvFilePath);
      process.exit(1);
    }

    console.log(`📂 Reading CSV: ${csvFilePath}`);

    // Connect to MySQL
    const connection = await mysql.createConnection({
      host: "misdashboard.c7o6omumk3c5.ap-south-1.rds.amazonaws.com",
      user: "admin",
      password: "Clab#2025",
      database: "MISdb",
    });

    const rows = [];
    let headersPrinted = false;

    const parser = csv({
      mapHeaders: ({ header }) => normalizeKey(header),
      skipLines: 0,
    });

    fs.createReadStream(csvFilePath)
      .on("error", (err) => {
        console.error("❌ File read error:", err.message);
        process.exit(1);
      })
      .pipe(parser)
      .on("headers", (headers) => {
        if (!headersPrinted) {
          console.log("📑 CSV Headers Detected:", headers);
          headersPrinted = true;
        }
      })
      .on("data", (row) => {
        // Clean branch code properly
        const branchCode = cleanBranchCode(row["branch code"] || row["br code"]);
        const branchName = (row["branch name"] || "").trim();
        const district = (row["district"] || "").trim();
        const cluster = (row["cluster"] || "").trim();

        const ac_00 = row["ac_00"] ? parseInt(row["ac_00"], 10) : 0;
        const balance_00 = row["balance_00"] ? parseFloat(row["balance_00"]) : 0;
        const ac_01 = row["ac_01"] ? parseInt(row["ac_01"], 10) : 0;
        const balance_01 = row["balance_01"] ? parseFloat(row["balance_01"]) : 0;
        const ac_02 = row["ac_02"] ? parseInt(row["ac_02"], 10) : 0;
        const balance_02 = row["balance_02"] ? parseFloat(row["balance_02"]) : 0;
        const ac_03 = row["ac_03"] ? parseInt(row["ac_03"], 10) : 0;
        const balance_03 = row["balance_03"] ? parseFloat(row["balance_03"]) : 0;
        const ac_04 = row["ac_04"] ? parseInt(row["ac_04"], 10) : 0;
        const balance_04 = row["balance_04"] ? parseFloat(row["balance_04"]) : 0;
        const total_ac = row["total_ac"] ? parseInt(row["total_ac"], 10) : 0;
        const total_balance = row["total_balance"] ? parseFloat(row["total_balance"]) : 0;

        // Skip invalid rows
        if (!branchCode) {
          console.warn("⚠️ Skipping row (missing Branch Code):", row);
          return;
        }

        rows.push([
          branchCode,
          branchName,
          district,
          cluster,
          ac_00,
          balance_00,
          ac_01,
          balance_01,
          ac_02,
          balance_02,
          ac_03,
          balance_03,
          ac_04,
          balance_04,
          total_ac,
          total_balance,
        ]);
      })
      .on("end", async () => {
        console.log(`✅ Parsed ${rows.length} valid rows`);

        if (rows.length === 0) {
          console.error("❌ No valid rows to insert. Check CSV headers.");
          process.exit(1);
        }

        const insertSql = `
          INSERT INTO SMA (\`Branch Code\`, \`Branch Name\`, District, Cluster,
            ac_00, balance_00, ac_01, balance_01,
            ac_02, balance_02, ac_03, balance_03,
            ac_04, balance_04, Total_ac, Total_balance)
          VALUES ?
        `;

        try {
          const [result] = await connection.query(insertSql, [rows]);
          console.log(`🎉 Inserted ${result.affectedRows} rows into SMA`);
        } catch (err) {
          console.error("❌ Error inserting rows:", err.message);
        }

        await connection.end();
        process.exit(0);
      });
  } catch (err) {
    console.error("❌ Unexpected error:", err);
    process.exit(1);
  }
}

importSMA();

