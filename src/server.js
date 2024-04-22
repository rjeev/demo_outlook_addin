const express = require("express");
const { Pool } = require("pg");
const cors = require("cors"); // Import the cors middleware

const app = express();
const port = 3009;

// Create a pool to handle database connections
const pool = new Pool({
  user: "postgres",
  host: "localhost",
  database: "emaildump",
  password: "lfkoP@ssw0rd",
  port: 5432,
});

// Middleware to parse JSON request bodies
app.use(express.json());

// Enable CORS for all routes
app.use(cors());

app.post("/emails", async (req, res) => {
  const {
    subject,
    senderEmail,
    senderName,
    ccRecipients,
    bccRecipients,
    body,
  } = req.body;

  try {
    const client = await pool.connect();
    const queryText =
      "INSERT INTO emails (subject, sender_email, sender_name, cc_recipients, bcc_recipients, body) VALUES ($1, $2, $3, $4, $5, $6)";
    const ccRecipientsArray = ccRecipients
      .map((recipient) => `{${recipient}}`)
      .join(",");
    const bccRecipientsArray = bccRecipients
      .map((recipient) => `{${recipient}}`)
      .join(",");
    const values = [
      subject,
      senderEmail,
      senderName,
      `{${ccRecipientsArray}}`,
      `{${bccRecipientsArray}}`,
      body,
    ];
    await client.query(queryText, values);
    client.release();
    res.status(201).send("Email information stored successfully.");
  } catch (error) {
    console.error("Error storing email information:", error); // Log the error

    // Send a more detailed error message to the client
    res.status(500).send(`Internal Server Error: ${error.message}`);
  }
});

app.get("/personal-data/:email", async (req, res) => {
  const { email } = req.params;

  try {
    const client = await pool.connect();
    const queryText = "SELECT * FROM persona WHERE email = $1";
    const { rows } = await client.query(queryText, [email]);
    client.release();

    if (rows.length === 0) {
      // If no data found for the provided email, send a 404 Not Found response
      return res
        .status(404)
        .send("Personal data not found for the provided email.");
    }

    // Send the personal data as a JSON response
    res.status(200).json(rows[0]);
  } catch (error) {
    console.error("Error fetching personal data:", error);
    res.status(500).send(`Internal Server Error: ${error.message}`);
  }
});

app.listen(port, () => {
  console.log(`Server is listening at http://localhost:${port}`);
});
