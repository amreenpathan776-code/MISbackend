const jwt = require("jsonwebtoken");
const sql = require("mssql");

const dbConfig = {
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  server: process.env.DB_SERVER,
  database: process.env.DB_NAME,
  options: {
    encrypt: false,
    trustServerCertificate: true
  }
};

async function dashboardAuthMiddleware(
  req,
  res,
  next
) {

  const authHeader =
    req.headers.authorization;

  if (!authHeader) {
    return res.status(401).json({
      message: "Login required"
    });
  }

  try {

    const token =
      authHeader.split(" ")[1];

    const decoded =
      jwt.verify(
        token,
        process.env.JWT_SECRET
      );

    const pool =
      await sql.connect(dbConfig);

    const result =
      await pool
        .request()
        .input(
          "employeeId",
          sql.VarChar,
          decoded.employeeId
        )
        .query(`
          SELECT
            DashboardSessionVersion
          FROM employees_auth
          WHERE [Emp No.] = @employeeId
        `);

    const rows =
      result.recordset;

    if (!rows.length) {

      return res.status(401).json({
        message: "User not found"
      });

    }


    if (
      Number(
        rows[0]
          .DashboardSessionVersion
      ) !==
      Number(
        decoded.sessionVersion
      )
    ) {


      return res.status(401).json({
        forceLogout: true,
        message:
          "This account was logged in from another device."
      });

    }

    req.user = decoded;

    next();

  } catch (err) {

    console.error(
      "DASHBOARD_AUTH_ERROR",
      err
    );

    return res.status(401).json({
      message: err.message
    });

  }
}

module.exports =
  dashboardAuthMiddleware;