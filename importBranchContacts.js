// importBranchContacts.js
const fs = require("fs");
const path = require("path");
const csv = require("csv-parser");
const mysql = require("mysql2/promise");

const csvFilePath = path.join(__dirname, "uploads", "Branch_Contact.csv");

(async () => {
  try {
    if (!fs.existsSync(csvFilePath)) {
      console.error("❌ CSV file not found at:", csvFilePath);
      process.exit(1);
    }

    // ---------- Skip initial blank or garbage lines ----------
    const raw = fs.readFileSync(csvFilePath, "utf8");
    const lines = raw.split(/\r?\n/);
    let skipLines = 0;
    for (let i = 0; i < lines.length; i++) {
      const lineTrim = lines[i].trim();
      if (lineTrim === "") {
        skipLines++;
        continue;
      }
      // Line only has separators
      if (/^[,;\t\s]+$/.test(lineTrim)) {
        skipLines++;
        continue;
      }
      break;
    }

    console.log(`📂 Reading CSV file: ${csvFilePath}`);
    if (skipLines > 0) {
      console.log(`ℹ️ Skipping ${skipLines} leading blank/comma-only line(s)`);
    }

    // ---------- MySQL Connection ----------
    const connection = await mysql.createConnection({
      host: "misdashboard.c7o6omumk3c5.ap-south-1.rds.amazonaws.com",
      user: "admin",
      password: "Clab#2025",
      database: "MISdb",
    });

    const rows = [];
    const skippedRows = [];
    let parsedCount = 0;

    const parser = csv({
      skipLines,
      mapHeaders: ({ header }) => {
        if (!header) return header;
        return header.replace(/^\uFEFF/, "").trim(); // Remove BOM and trim
      },
    });

    fs.createReadStream(csvFilePath)
      .on("error", (err) => {
        console.error("❌ File read error:", err.message);
        process.exit(1);
      })
      .pipe(parser)
      .on("data", (row) => {
        parsedCount++;

        // Support multiple possible header variations
        const branchCode = (row["Br Code"] ?? row["Branch_Code"] ?? row["branch code"] ?? row["BranchCode"] ?? "")
          .toString()
          .trim();

        const cluster = (row["Cluster"] ?? row["cluster"] ?? row["Cluster Name"] ?? "")
          .toString()
          .trim();

        const branchName = (row["Branch Name"] ?? row["Branch_Name"] ?? row["branch name"] ?? row["BranchName"] ?? "")
          .toString()
          .trim();

        const branchManager = (row["Branch Manager"] ?? row["Branch_Manager"] ?? row["branch manager"] ?? row["BranchManager"] ?? "")
          .toString()
          .trim();

        const contactNumber = (row["Contact Number"] ?? row["Contact"] ?? row["Contact Info"] ?? row["contact info"] ?? "")
          .toString()
          .trim();

        // Detect accidental header-like rows
        const isHeaderLike = [branchCode, cluster, branchName, branchManager, contactNumber].some((v) =>
          String(v).toLowerCase().includes("branch code") ||
          String(v).toLowerCase().includes("cluster") ||
          String(v).toLowerCase().includes("branch name") ||
          String(v).toLowerCase().includes("branch manager") ||
          String(v).toLowerCase().includes("contact")
        );

        if (isHeaderLike) {
          skippedRows.push({ reason: "header-like row", raw: row });
          return;
        }

        // Clean contact number: keep only digits and optional leading "+"
        const normalizedContact = contactNumber.replace(/[^\d+]/g, "");

        // Validate required fields
        if (!branchCode || !cluster || !branchName || !branchManager || !normalizedContact) {
          skippedRows.push({ reason: "missing-required-field", raw: row });
          return;
        }

        rows.push({
          branchCode: parseInt(branchCode, 10),
          cluster,
          branchName,
          branchManager,
          contactNumber: normalizedContact,
        });
      })
      .on("end", async () => {
        console.log(`✅ Parsed ${parsedCount} rows from CSV.`);
        console.log(`ℹ️ Valid rows to import: ${rows.length}. Skipped rows: ${skippedRows.length}.`);

        // Insert SQL for the table
        const insertSql = `
          INSERT INTO Branch_Contact (\`Br Code\`, \`Cluster\`, \`Branch Name\`, \`Branch Manager\`, \`Contact Number\`)
          VALUES (?, ?, ?, ?, ?)
        `;

        let inserted = 0;
        for (const r of rows) {
          try {
            await connection.execute(insertSql, [
              r.branchCode,
              r.cluster,
              r.branchName,
              r.branchManager,
              r.contactNumber,
            ]);
            inserted++;
          } catch (err) {
            if (err.code === "ER_DUP_ENTRY") {
              console.warn(`⚠️ Duplicate entry for Branch Code ${r.branchCode}, skipping.`);
            } else {
              console.error("❌ Error inserting row:", r, err.message);
            }
          }
        }

        console.log(`🎉 Inserted ${inserted} rows into Branch_Contact.`);

        if (skippedRows.length) {
          console.log("⚠️ Skipped rows sample (first 5):", skippedRows.slice(0, 5));
        }

        await connection.end();
        process.exit(0);
      });
  } catch (err) {
    console.error("❌ Unexpected error:", err);
    process.exit(1);
  }
})();


