// backend/db.js
const mysql = require('mysql2');

const connection = mysql.createConnection({
  host: 'localhost',          // or your DB host
  user: 'root',               // your MySQL username
  password: 'Clab',  // your MySQL password
  database: 'MIS', // the name of your DB
  multipleStatements: true,
  localInfile: true,
});

connection.connect((err) => {
  if (err) {
    console.error('❌ Database connection failed: ' + err.stack);
    return;
  }
  console.log('✅ Connected to MySQL as id ' + connection.threadId);
});

module.exports = connection;
