const fs = require("fs");
const path = require("path");

function sanitize(input) {
  return String(input || "")
    .replace(/[<>:"\\|?*\x00-\x1F]/g, "_")
    .replace(/\.\./g, "_")
    .replace(/\s+/g, "_")
    .substring(0, 50);
}

function writeDailyLog(type, message) {
  try {

    const now = new Date();

    const today = now.toLocaleDateString(
      "en-CA",
      {
        timeZone: "Asia/Kolkata"
      }
    );

    const time = now.toLocaleString(
      "en-IN",
      {
        timeZone: "Asia/Kolkata"
      }
    );

    // Split path
    const parts = String(type)
      .split("/")
      .map(sanitize);

    // Folder path
    const logDir = path.join(
      __dirname,
      "logs",
      ...parts
    );

    if (!fs.existsSync(logDir)) {
      fs.mkdirSync(
        logDir,
        { recursive: true }
      );
    }

    // File name = last folder name
    const fileName = parts[parts.length - 1];

    const logFile = path.join(
      logDir,
      `${today}.${fileName}.log`
    );

    const safeMessage = String(message || "")
      .replace(/\x00/g, "")
      .substring(0, 5000);

    fs.appendFileSync(
      logFile,
      `${time} | ${safeMessage}\n`
    );

  } catch (err) {

    console.error(
      "LOGGER ERROR:",
      err.message
    );

  }
}

module.exports = writeDailyLog;